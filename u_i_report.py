"""
Uturn(180°円弧+両端2直線) と Iturn(U内のコの字：直角2箇所の3直線) を
スケッチから抽出 → 寸法計測 → .xlsx 出力（1シート：左=Results, 右=Snapshots）
さらに、各パターンを選択して ReframeOnSelection で矩形拡大 → PNGスナップ取得。

要件まとめ
----------
- Uturn の円弧半径: 1〜2 mm
- U/I 抽出時、受理済みパターン（Uの円弧中心）との最小間隔: 100 mm
- U の直線（2本）と I の直線（3本）とのセグメント最短距離を出力
- Excelは1シートで、左側が計測表（Results）、右側がスナップ（Snapshots）

注意
----
- 単位は「スケッチ作成の単位（通常は mm）」を前提に計算。
- このスクリプトは Windows の COM 経由で CATIA と Excel を操作。
  Mac では動作しないため、必要に応じて仮想環境やリモートWindowsを用意。
"""

from __future__ import annotations

import math
import fnmatch
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Tuple

try:
    # pywin32 の COM ラッパー。CATIA/Excel を操作するために必須。
    # pip install pywin32
    import win32com.client as win32
except ImportError as e:
    raise SystemExit(
        "pywin32 が必要です。`pip install pywin32` を実行してください。"
    ) from e


# =============================================================================
# 設定（プロジェクトごとに調整する想定）
# =============================================================================
# 名前フィルタ（スケッチ名 or 要素名が該当する時に優先的に抽出）
NAME_LIKE_U = "Uturn*"           # "" にすると名前フィルタを無効化
NAME_LIKE_I = "Iturn*"

# Uturn 円弧の半径範囲（mm）
R_MIN = 1.0
R_MAX = 2.0

# 幾何判定の許容
TOL_ANG_DEG = 0.5                # U 円弧角が 180°±TOL_ANG_DEG
TOL_XY = 0.01                    # 端点一致（mm）
TOL_RIGHT = 3.0                  # 直角判定（90°±TOL_RIGHT）

# 多数・連続適合時の「採用間引き」：受理済みパターンとの最小距離（mm）
# （U の円弧中心座標どうしの距離で判定）
MIN_SPACING = 100.0

# 出力関連
OUT_PATH = r"C:\temp\U_I_report.xlsx"     # Excel 出力先
SNAP_DIR = r"C:\temp\ui_snaps"            # スナップ画像の保存ディレクトリ
SNAPSHOT_HEIGHT = 240                     # Excel 貼付時の画像高さ（px）
START_COL_SNAPSHOT = 22                   # 右側の画像カラム開始位置（V列=22）

# =============================================================================
# データ構造（読みやすさ・型安全性向上のため dataclass を使用）
# =============================================================================


@dataclass
class Line2D:
    """
    スケッチ上の 2D 直線を表す。
    - (x1, y1) と (x2, y2) は端点座標
    - length は自動計算された長さ
    - obj は元の CATIA ジオメトリ（COM オブジェクト）
    - name は CATIA 上の要素名（任意）
    """
    x1: float
    y1: float
    x2: float
    y2: float
    length: float
    obj: Any
    name: str = ""


@dataclass
class Arc2D:
    """
    スケッチ上の 2D 円弧を表す。
    - (cx, cy) は中心座標
    - r は半径
    - a0, a1 は開始/終了角（ラジアン）
    - deg_span は円弧角（度）
    - obj は元の CATIA ジオメトリ（COM オブジェクト）
    - name は CATIA 上の要素名（任意）
    """
    cx: float
    cy: float
    r: float
    a0: float
    a1: float
    deg_span: float
    obj: Any
    name: str = ""


# =============================================================================
# CATIA 接続
# =============================================================================

def connect() -> Tuple[Any, Any]:
    """
    CATIA.Application に接続し、アクティブドキュメントを取得する。

    Returns
    -------
    (app, doc): (CATIA.Application, Document)
        app: CATIA アプリケーション COM オブジェクト
        doc: アクティブドキュメント（Part または Product）

    Raises
    ------
    RuntimeError
        アクティブドキュメントが無い場合
    """
    # 既存の CATIA セッションにアタッチ（起動していなければエラー）
    app = win32.Dispatch("CATIA.Application")
    doc = app.ActiveDocument
    if doc is None:
        raise RuntimeError("CATIAで対象ドキュメントを開いてください。")
    return app, doc


# =============================================================================
# 幾何ユーティリティ（座標・角度・距離などの計算）
# =============================================================================

def dist2d(x1: float, y1: float, x2: float, y2: float) -> float:
    """2点間距離（2D のユークリッド距離）を返す。"""
    return math.hypot(x2 - x1, y2 - y1)


def nearly(a: float, b: float, tol: float) -> bool:
    """a と b が tol 以内で等しいと見なす。"""
    return abs(a - b) <= tol


def normalize_angle(rad: float) -> float:
    """
    角度 rad（ラジアン）を [-pi, pi] に畳み込み、その絶対値を返す。
    円弧角の比較を方向性に依存せず行うために使用。
    """
    while rad > math.pi:
        rad -= 2 * math.pi
    while rad < -math.pi:
        rad += 2 * math.pi
    return abs(rad)


def arc_endpoints(cx: float, cy: float, r: float,
                  a0: float, a1: float) -> Tuple[float, float, float, float]:
    """
    円弧の開始・終了点（x, y）を返す。
    CATIA の Arc2D/Circle2D の StartAngle/EndAngle はラジアン想定。
    """
    sx = cx + r * math.cos(a0)
    sy = cy + r * math.sin(a0)
    ex = cx + r * math.cos(a1)
    ey = cy + r * math.sin(a1)
    return sx, sy, ex, ey


def angle_between(ax: float, ay: float, bx: float, by: float) -> float:
    """
    ベクトル (ax, ay) と (bx, by) のなす角度（度）を返す。
    0〜180°の範囲に正規化される。
    """
    da = math.hypot(ax, ay)
    db = math.hypot(bx, by)
    if da == 0 or db == 0:
        return 0.0
    d = (ax * bx + ay * by) / (da * db)
    d = max(-1.0, min(1.0, d))  # 数値誤差で範囲外に出ないようクリップ
    deg = math.degrees(math.atan2(math.sqrt(1 - d * d), d))
    return deg if deg >= 0 else deg + 180.0


def inside_circle(cx: float, cy: float, r: float, x: float, y: float) -> bool:
    """点 (x, y) が半径 r の円（中心 (cx, cy)）の内側にあるか判定。境界含む。"""
    return dist2d(cx, cy, x, y) <= r


def dist_pt_to_seg(px: float, py: float,
                   x1: float, y1: float, x2: float, y2: float) -> float:
    """
    点 (px, py) と線分 [(x1, y1) - (x2, y2)] の最短距離を返す。
    投影点が線分外なら端点までの距離になる。
    """
    vx, vy = x2 - x1, y2 - y1
    wx, wy = px - x1, py - y1
    vv = vx * vx + vy * vy
    if vv == 0:
        # 長さ0（実質点）へのガード
        return math.hypot(px - x1, py - y1)
    t = (wx * vx + wy * vy) / vv  # 最近接点の線分上パラメータ
    t = 0.0 if t < 0 else (1.0 if t > 1 else t)  # 0〜1にクリップ
    cx, cy = x1 + t * vx, y1 + t * vy
    return math.hypot(px - cx, py - cy)


def dist_seg_to_seg(la: Line2D, lb: Line2D) -> float:
    """
    2つの線分 la, lb の最短距離を返す。
    端点→相手線分への距離の最小値で十分（2Dではこれで最短距離が得られる）。
    """
    d1 = dist_pt_to_seg(la.x1, la.y1, lb.x1, lb.y1, lb.x2, lb.y2)
    d2 = dist_pt_to_seg(la.x2, la.y2, lb.x1, lb.y1, lb.x2, lb.y2)
    d3 = dist_pt_to_seg(lb.x1, lb.y1, la.x1, la.y1, la.x2, la.y2)
    d4 = dist_pt_to_seg(lb.x2, lb.y2, la.x1, la.y1, la.x2, la.y2)
    return min(d1, d2, d3, d4)


def min3(a: float, b: float, c: float) -> float:
    """3つの値の最小値。可読性のための薄いラッパー。"""
    return min(a, b, c)


# =============================================================================
# 2D 要素抽出（CATIA の COM オブジェクトを軽く抽象化）
# =============================================================================

def is_line2d(g: Any) -> bool:
    """
    GeometricElement g が Line2D 相当かおおまかに判定。
    - StartPoint/EndPoint が取れれば Line2D と見なす。
    """
    try:
        _ = g.StartPoint
        _ = g.EndPoint
        return True
    except Exception:
        return False


def is_arc2d(g: Any) -> bool:
    """
    GeometricElement g が Arc2D/Circle2D 相当かおおまかに判定。
    - CenterPoint/Radius/StartAngle/EndAngle が取れれば円弧と見なす。
    """
    try:
        _ = g.CenterPoint
        _ = g.Radius
        _ = g.StartAngle
        _ = g.EndAngle
        return True
    except Exception:
        return False


def fill_line2d(g: Any) -> Line2D | None:
    """
    CATIA の Line2D から Line2D dataclass を生成。
    エラー時は None（例えば投影や定義不足など）。
    """
    try:
        p1, p2 = g.StartPoint, g.EndPoint
        x1, y1 = float(p1.X), float(p1.Y)
        x2, y2 = float(p2.X), float(p2.Y)
        return Line2D(
            x1=x1, y1=y1, x2=x2, y2=y2,
            length=dist2d(x1, y1, x2, y2),
            obj=g, name=getattr(g, "Name", "")
        )
    except Exception:
        return None


def fill_arc2d(g: Any) -> Arc2D | None:
    """
    CATIA の Arc2D/Circle2D から Arc2D dataclass を生成。
    StartAngle/EndAngle が無い（真円）場合は対象外。
    """
    try:
        c = g.CenterPoint
        r = float(g.Radius)
        a0 = float(g.StartAngle)
        a1 = float(g.EndAngle)
        span = normalize_angle(a1 - a0)
        deg_span = math.degrees(span)
        return Arc2D(
            cx=float(c.X), cy=float(c.Y), r=r,
            a0=a0, a1=a1, deg_span=deg_span, obj=g,
            name=getattr(g, "Name", "")
        )
    except Exception:
        return None


# =============================================================================
# 探索ヘルパ（U の端点に接続する直線、直角接続の探索 など）
# =============================================================================

def find_line_at_point(lines: Dict[str, Line2D],
                       x: float, y: float, tol: float
                       ) -> Tuple[str | None, float, float, float]:
    """
    点 (x, y) に端点が一致する線分を検索。
    一致した場合、その線分キーと「反対側端点座標（=自由端）」、長さを返す。
    """
    for k, line in lines.items():
        if nearly(line.x1, x, tol) and nearly(line.y1, y, tol):
            return k, line.x2, line.y2, line.length
        if nearly(line.x2, x, tol) and nearly(line.y2, y, tol):
            return k, line.x1, line.y1, line.length
    return None, 0.0, 0.0, 0.0


def right_angle_neighbors_at(lines: Dict[str, Line2D],
                             lb: Line2D,
                             at_start: bool,
                             tol_xy: float,
                             tol_right: float) -> List[str]:
    """
    線 lb の start 端 or end 端に直角（90°±tol_right）で接続する線のキー集合を返す。
    直角判定はベクトル間角度を使用。
    """
    bx, by = (lb.x1, lb.y1) if at_start else (lb.x2, lb.y2)
    vx, vy = ((lb.x2 - lb.x1, lb.y2 - lb.y1)
              if at_start else (lb.x1 - lb.x2, lb.y1 - lb.y2))
    result: List[str] = []

    for k, line in lines.items():
        if line is lb:
            continue

        # lb の端点 (bx, by) を共有しているか（端点一致）
        touch = False
        ux = uy = 0.0
        if nearly(line.x1, bx, tol_xy) and nearly(line.y1, by, tol_xy):
            ux, uy = line.x2 - line.x1, line.y2 - line.y1
            touch = True
        elif nearly(line.x2, bx, tol_xy) and nearly(line.y2, by, tol_xy):
            ux, uy = line.x1 - line.x2, line.y1 - line.y2
            touch = True

        if not touch:
            continue

        # 直角判定（90°±tol_right）
        ang = angle_between(vx, vy, ux, uy)
        if abs(ang - 90.0) <= tol_right:
            result.append(k)

    return result


def free_end_opposite_to(l: Line2D, lb: Line2D, tol: float) -> Tuple[float, float]:
    """
    線 l が lb と接していると仮定し、l の「自由端」座標を返す。
    - l の片端が lb のいずれか端点に一致していれば、もう一端が自由端。
    """
    if (nearly(l.x1, lb.x1, tol) and nearly(l.y1, lb.y1, tol)) or \
       (nearly(l.x1, lb.x2, tol) and nearly(l.y1, lb.y2, tol)):
        return l.x2, l.y2
    return l.x1, l.y1


# =============================================================================
# スナップショット撮影（選択 → ReframeOnSelection → PNG保存）
# =============================================================================

def sanitize_filename(s: str) -> str:
    """Windows のファイル名に使えない文字をアンダースコアに置換。"""
    for ch in r':|\/*?"<> &':
        s = s.replace(ch, "_")
    return s


def ensure_folder(path: str) -> None:
    """フォルダが無ければ作成。"""
    Path(path).mkdir(parents=True, exist_ok=True)


def capture_pattern_snapshot(part_doc: Any,
                             objs: List[Any],
                             png_path: str) -> None:
    """
    指定オブジェクト群を選択 → ReframeOnSelection → 画像キャプチャ。
    - ReframeOnSelection が使えない環境では Reframe にフォールバック。
    """
    sel = part_doc.Selection
    part = part_doc.Part
    sel.Clear()

    # 幾何オブジェクト（COM）から参照を作成して選択セットに追加
    for obj in objs:
        if obj is None:
            continue
        try:
            ref = part.CreateReferenceFromObject(obj)
            if ref is not None:
                sel.Add(ref)
        except Exception:
            # 一部のオブジェクトで参照作成できない場合があるため握りつぶす
            pass

    viewer = part_doc.Application.ActiveWindow.ActiveViewer
    try:
        # 選択範囲に画面をフィット
        viewer.ReframeOnSelection()
    except Exception:
        # 古い環境などでメソッドが使えない場合
        viewer.Reframe()

    # 画面キャプチャ（PNG）
    part_doc.Application.ActiveWindow.CapturePictureFile(png_path, "png")
    sel.Clear()


# =============================================================================
# 主要ロジック：スケッチ解析（抽出→計測→スナップ→行レコード構築）
# =============================================================================

def analyze_sketch(part_doc: Any,
                   part: Any,
                   body: Any,
                   sk: Any,
                   anchors: List[Tuple[float, float]],
                   rows: List[List[Any]],
                   snaps: Dict[str, str]) -> None:
    """
    1スケッチを解析して、U→I の組み合わせを探し、行データとスナップを作成。
    - U: 半径1〜2mm & 180°±Tol の円弧 + 端点一致の2直線
    - I: 3直線の組で、中央線 Lb の両端が別直線と直角接続、主要点が U 円内
    - 採用済みパターン（U中心）との距離が MIN_SPACING 未満ならスキップ
    """
    geos = getattr(sk, "GeometricElements", None)
    if geos is None:
        return

    # スケッチ名が NameLike に合致するなら、要素名が合致しなくても優先的に通す
    use_u = len(NAME_LIKE_U) > 0
    use_i = len(NAME_LIKE_I) > 0
    sketch_has_u = (not use_u) or fnmatch.fnmatch(getattr(sk, "Name", ""), NAME_LIKE_U)
    sketch_has_i = (not use_i) or fnmatch.fnmatch(getattr(sk, "Name", ""), NAME_LIKE_I)

    # 要素収集（Line2D / Arc2D）
    lines: Dict[str, Line2D] = {}
    arcs: Dict[str, Arc2D] = {}

    for i in range(1, int(geos.Count) + 1):
        g = geos.Item(i)

        if is_line2d(g):
            line = fill_line2d(g)
            if line:
                lines[str(len(lines) + 1)] = line

        elif is_arc2d(g):
            arc = fill_arc2d(g)
            # U向けの数値条件（半径 & 角度）を満たす円弧のみ候補に
            if arc and (R_MIN <= arc.r <= R_MAX) and (abs(arc.deg_span - 180.0) <= TOL_ANG_DEG):
                # 名前フィルタ：スケッチ名で合致しない場合は、円弧名でもチェック
                if use_u and not sketch_has_u:
                    if not fnmatch.fnmatch(getattr(g, "Name", ""), NAME_LIKE_U):
                        continue
                arc.name = getattr(g, "Name", "")
                arcs[str(len(arcs) + 1)] = arc

    if not arcs:
        return

    # 後続の出力（行データ）で使う識別情報
    doc_name = getattr(part_doc, "Name", "")
    part_name = getattr(part, "Name", "")
    body_name = getattr(body, "Name", "")
    sk_name = getattr(sk, "Name", "")

    # ---- U 単位でループし、I を探す ----
    for ak, u in arcs.items():
        # Uの円弧端点座標を算出
        sx, sy, ex, ey = arc_endpoints(u.cx, u.cy, u.r, u.a0, u.a1)

        # 端点一致（TOL_XY）する直線を探し、各直線の「自由端（円弧でない側）」を取得
        k_l1, l1_ox, l1_oy, l1_len = find_line_at_point(lines, sx, sy, TOL_XY)
        k_l2, l2_ox, l2_oy, l2_len = find_line_at_point(lines, ex, ey, TOL_XY)
        if not (k_l1 and k_l2):
            # どちらか片方の端点で直線が見つからないなら U として不成立
            continue

        # U の開口幅：2直線の「自由端」間距離
        opening = dist2d(l1_ox, l1_oy, l2_ox, l2_oy)
        u_id = f"{sk_name}:U{ak}"

        # 多数・連続配置の間引き：既存アンカー（U中心）との距離が MIN_SPACING 未満なら除外
        if any(dist2d(ax, ay, u.cx, u.cy) < MIN_SPACING for ax, ay in anchors):
            continue

        found_i = False  # この U に対して I が見つかったかどうか

        # I 候補：中央線 Lb の両端で、別線と直角に接続（90°±TOL_RIGHT）
        for lk, lb in list(lines.items()):
            s_ng = right_angle_neighbors_at(lines, lb, at_start=True,
                                            tol_xy=TOL_XY, tol_right=TOL_RIGHT)
            e_ng = right_angle_neighbors_at(lines, lb, at_start=False,
                                            tol_xy=TOL_XY, tol_right=TOL_RIGHT)
            if not s_ng or not e_ng:
                continue  # 両端で直角接続する相手がいなければ I 不成立

            for sk1 in s_ng:
                for sk2 in e_ng:
                    if sk1 == sk2:
                        continue  # 同じ線を両端に使うのは不可
                    la = lines[sk1]
                    lc = lines[sk2]

                    # I の「自由端」2点（Lbに接していない側）を取得
                    a_fx, a_fy = free_end_opposite_to(la, lb, TOL_XY)
                    c_fx, c_fy = free_end_opposite_to(lc, lb, TOL_XY)

                    # I の主要点（Lb両端 + 自由端2点）が U 円の内側にあるか
                    inside = (
                        inside_circle(u.cx, u.cy, u.r - 1e-3, lb.x1, lb.y1) and
                        inside_circle(u.cx, u.cy, u.r - 1e-3, lb.x2, lb.y2) and
                        inside_circle(u.cx, u.cy, u.r - 1e-3, a_fx, a_fy) and
                        inside_circle(u.cx, u.cy, u.r - 1e-3, c_fx, c_fy)
                    )
                    if not inside:
                        continue

                    # U の2直線（端点一致で見つけた線）と I の3直線のセグメント最短距離
                    ul1 = lines[k_l1]
                    ul2 = lines[k_l2]

                    d_l1a = dist_seg_to_seg(ul1, la)
                    d_l1b = dist_seg_to_seg(ul1, lb)
                    d_l1c = dist_seg_to_seg(ul1, lc)
                    u2i_l1_min = min3(d_l1a, d_l1b, d_l1c)

                    d_l2a = dist_seg_to_seg(ul2, la)
                    d_l2b = dist_seg_to_seg(ul2, lb)
                    d_l2c = dist_seg_to_seg(ul2, lc)
                    u2i_l2_min = min3(d_l2a, d_l2b, d_l2c)

                    u2i_min = min(u2i_l1_min, u2i_l2_min)

                    # I の識別子（中央線キー + 両端キー）
                    right_ok = "OK"
                    i_id = f"{sk_name}:I{lk}-{sk1}-{sk2}"

                    # ---- スナップ（パターン単位）----
                    snap_label = f"{u_id} + {i_id}"
                    img_path = os.path.join(
                        SNAP_DIR, sanitize_filename(snap_label) + ".png"
                    )
                    capture_pattern_snapshot(
                        part_doc,
                        [u.obj, ul1.obj, ul2.obj, la.obj, lb.obj, lc.obj],
                        img_path,
                    )
                    snaps.setdefault(snap_label, img_path)

                    # ---- 行追加（Excel 左表に出す 1 レコード）----
                    rows.append([
                        doc_name, part_name, body_name, sk_name, "U+I",
                        rounder(u_id), round(u.r, 3), round(u.deg_span, 3),
                        round(l1_len, 3), round(l2_len, 3), round(opening, 3),
                        rounder(i_id), round(la.length, 3), round(lb.length, 3),
                        round(lc.length, 3), right_ok,
                        round(u2i_l1_min, 3), round(u2i_l2_min, 3), round(u2i_min, 3),
                        snap_label,
                    ])
                    found_i = True

                    # この U を採用したのでアンカー追加（以降のパターン間引きに使用）
                    anchors.append((u.cx, u.cy))

        # I が見つからなかった場合でも U 単独行を出力（スナップは U 要素のみ）
        if not found_i:
            snap_label_u = u_id
            img_path_u = os.path.join(
                SNAP_DIR, sanitize_filename(snap_label_u) + ".png"
            )
            capture_pattern_snapshot(
                part_doc,
                [u.obj, lines[k_l1].obj, lines[k_l2].obj],
                img_path_u,
            )
            snaps.setdefault(snap_label_u, img_path_u)

            rows.append([
                doc_name, part_name, body_name, sk_name, "U",
                rounder(u_id), round(u.r, 3), round(u.deg_span, 3),
                round(l1_len, 3), round(l2_len, 3), round(opening, 3),
                "", "", "", "", "",
                "", "", "",
                snap_label_u,
            ])
            anchors.append((u.cx, u.cy))


def rounder(s: Any) -> Any:
    """
    Excel に書くときに、識別子などの文字列が勝手に数値/日付解釈されるのを
    避けたい場合は、必要に応じて前置記号をつける等の処理をここに実装。
    今回はそのまま返す（説明用のフック）。
    """
    return s


# =============================================================================
# Part / Product 走査（複数 Part にも対応）
# =============================================================================

def process_part(part_doc: Any,
                 anchors: List[Tuple[float, float]],
                 rows: List[List[Any]],
                 snaps: Dict[str, str]) -> None:
    """
    1つの PartDocument を走査し、全 Body → Sketch を解析。
    """
    part = part_doc.Part
    bodies = getattr(part, "Bodies", None)
    if bodies is None:
        return

    for b in range(1, int(bodies.Count) + 1):
        body = bodies.Item(b)
        sketches = getattr(body, "Sketches", None)
        if sketches is None:
            continue

        for s in range(1, int(sketches.Count) + 1):
            sk = sketches.Item(s)
            analyze_sketch(part_doc, part, body, sk, anchors, rows, snaps)


def process_product(prod_doc: Any,
                    anchors: List[Tuple[float, float]],
                    rows: List[List[Any]],
                    snaps: Dict[str, str]) -> None:
    """
    ProductDocument を走査し、配下の参照 PartDocument を順次処理。
    """
    prod = prod_doc.Product
    count = int(prod.Products.Count)
    for i in range(1, count + 1):
        child = prod.Products.Item(i)

        # 参照先ドキュメントを取得（PartDocument のはず）
        ref_doc = child.ReferenceProduct.Parent
        # CATIA の API では Type で判別（"Part" or "Product" など）
        if getattr(ref_doc, "Type", "") == "Part":
            process_part(ref_doc, anchors, rows, snaps)


# =============================================================================
# Excel 出力（1シート左右並置：左=結果、右=スナップ）
# =============================================================================

def export_excel_one_sheet(rows: List[List[Any]],
                           snaps: Dict[str, str],
                           out_path: str) -> None:
    """
    Excel の 1 シートに、左側に計測結果（表）、右側にスナップ画像を並置。
    画像の高さは SNAPSHOT_HEIGHT に収まるよう縮小。
    """
    xl = win32.Dispatch("Excel.Application")
    xl.Visible = False
    wb = xl.Workbooks.Add()
    sh = wb.Sheets(1)
    sh.Name = "Report"

    # 左：ヘッダ行
    headers = [
        "Doc", "Part", "Body", "Sketch", "Pattern",
        "U_ID", "U_R(mm)", "U_Ang(deg)", "U_L1(mm)", "U_L2(mm)", "U_Opening(mm)",
        "I_ID", "I_La(mm)", "I_Lb(mm)", "I_Lc(mm)", "I_RightOK",
        "U2I_L1_Min(mm)", "U2I_L2_Min(mm)", "U2I_Min(mm)", "SnapLabel",
    ]
    for c, h in enumerate(headers, start=1):
        sh.Cells(1, c).Value = h

    # 左：データ行
    for r_idx, rec in enumerate(rows, start=2):
        for c_idx, val in enumerate(rec, start=1):
            sh.Cells(r_idx, c_idx).Value = val

    # 列幅自動調整
    sh.Columns("A:T").AutoFit()

    # 右：Snapshots（START_COL_SNAPSHOT 列から並べる）
    sc = START_COL_SNAPSHOT
    sh.Cells(1, sc).Value = "Pattern (SnapLabel)"
    sh.Cells(1, sc + 1).Value = "Image"
    sh.Columns(sc).ColumnWidth = 48
    sh.Columns(sc + 1).ColumnWidth = 48

    row = 2  # 画像開始行
    for label, path in snaps.items():
        sh.Cells(row, sc).Value = label
        try:
            pic = sh.Pictures().Insert(path)  # 画像挿入
            pic.Top = sh.Cells(row, sc + 1).Top
            pic.Left = sh.Cells(row, sc + 1).Left

            # 画像が大きすぎる場合は高さベースでスケール
            if pic.Height > SNAPSHOT_HEIGHT:
                scale = SNAPSHOT_HEIGHT / pic.Height
                pic.Height = SNAPSHOT_HEIGHT
                pic.Width = pic.Width * scale

            # 画像の高さに応じて次の挿入行をオフセット
            row += int(pic.Height / sh.Rows(1).Height) + 2

        except Exception:
            # 画像が見つからない・読み込めない等の場合はスキップ
            row += 1

    # ファイル保存（.xlsx）。フォルダが無ければ作成。
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.SaveAs(out_path, 51)  # 51 = xlOpenXMLWorkbook
    except Exception:
        # Excel のバージョンや権限等でエラーの場合のフォールバック
        wb.SaveAs(out_path)

    wb.Close(SaveChanges=False)
    xl.Quit()


# =============================================================================
# エントリポイント
# =============================================================================

def main() -> None:
    """
    実行手順：
      1) CATIA を起動して対象 Part / Product を開く
        （Windows で pip install pywin32。）
      2) `python u_i_report.py` を実行
      3) C:\temp\U_I_report.xlsx（1シートで左=Results, 右=Snapshots） と C:\temp\ui_snaps\*.png が生成

      補足:
        Excelの「SnapLabel」列は、右側の画像ラベルと一致（行⇔画像の対応確認に便利）。
        Excelが画像挿入で高さオートに追従しないケースがあり、そこは行のオフセット計算で概ね並ぶよう設定。必要なら SNAPSHOT_HEIGHT や開始列 START_COL_SNAPSHOT を調整。
        「名前フィルタ」を完全に無視したい場合は NAME_LIKE_U = ""/NAME_LIKE_I = "" と設定。
    """
    ensure_folder(SNAP_DIR)

    # CATIA に接続し、アクティブドキュメントを取得
    _, doc = connect()

    # 出力用コンテナ
    rows: List[List[Any]] = []                  # 左表の全行
    snaps: Dict[str, str] = {}                  # 右側のスナップ（label -> path）
    anchors: List[Tuple[float, float]] = []     # 採用済みパターンのU中心座標

    # ドキュメント種別で処理を分岐（Part or Product）
    doc_type = getattr(doc, "Type", "")
    if doc_type == "Part":
        process_part(doc, anchors, rows, snaps)
    elif doc_type == "Product":
        process_product(doc, anchors, rows, snaps)
    else:
        raise RuntimeError("Part または Product を開いてください。")

    export_excel_one_sheet(rows, snaps, OUT_PATH)
    print(f"完了: 行数={len(rows)}  出力={OUT_PATH}")


if __name__ == "__main__":
    main()
