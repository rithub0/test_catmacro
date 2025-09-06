"""
U/I 3D Extractor (Python + pywin32)
===================================
CATIAの3Dモデルから、指定形状（Uturn: 半円弧+直線2、Iturn: U内のコの字）を抽出して検査し、
HTMLレポート + PNGスナップで資料化するスクリプト。

★ポイント
- スケッチ不要。トポロジ(Topology)の Face / Edge を直接走査。
- Edgeの「端点/中点」「長さ」「半径」を SPAWorkbench.Measurable から取得。
  - 弧角 ≈ 長さ / 半径（deg）で半円(≈180°)判定。
- 面ごとのローカル2D座標系(平面の原点/2軸)に射影し、幾何判定（直角/内側性/距離など）を安定化。
- U/Iの各候補を検査：半径範囲・直角(90°±tol)・半円内半平面条件・U↔Iの最短距離(3D)。
- レポートは HTML（左：表、右：サムネ画像）。Excel不要。
- スナップ撮影は ReframeOnSelection → CapturePictureFile。負荷対策で間引き可。

前提
----
- Windows + CATIA V5 がインストールされ、対象 Part/Product を開いた状態で実行
- Python 3 + pywin32 (pip install pywin32)

使い方
------
1) CATIAで Part か Product を開く
2) 本スクリプトを `ui_3d_report.py` などで保存
3) `pip install pywin32`
4) `python ui_3d_report.py`
5) `C:\temp\ui3d_report\Report.html` と `snaps\*.png` が生成される

注意
----
- Macでは pywin32/COM が使えないため、Windows環境が必要（仮想/リモートでも可）
- CATIAのリリース差でMeasurableの挙動が多少異なることがあります（例外を握り、近似で補完）
"""

from __future__ import annotations

import math
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    import win32com.client as win32
except Exception as e:
    raise SystemExit(
        "pywin32 が必要です。先に `pip install pywin32` を実行してください。"
    ) from e


# ========================= ユーザー設定（要件に合わせて調整） =========================

R_MIN: float = 1.0            # U弧 半径下限 [mm]
R_MAX: float = 2.0            # U弧 半径上限 [mm]
TOL_DEG_ARC: float = 1.0      # 半円判定 180°±この値 [deg]
TOL_RIGHT_DEG: float = 3.0    # 直角判定 90°±この値 [deg]
TOL_JOIN: float = 0.02        # 端点一致（3D）許容 [mm]
MIN_SPACING3D: float = 100.0  # 採用Uパターンの3D中心どうし最小距離 [mm]

ENABLE_SNAPS: bool = True     # スナップ撮影 ON/OFF
SNAP_EVERY_N: int = 1         # 何件に1回撮るか（1=毎回、2=2件に1回…）
SNAP_MAX: int = 200           # 最大撮影枚数
SNAP_HEIGHT_PX: int = 240     # HTML上のサムネ高さ(px)

OUT_DIR: str = r"C:\temp\ui3d_report"  # 出力先
SNAP_DIRNAME: str = "snaps"            # スナップ保存サブフォルダ名

# ========================= 内部カウンタ（スナップ間引き用） =========================

_snap_count = 0
_pat_count = 0  # パターンのカウント（U+I or U単）


# ========================= 汎用ユーティリティ（ベクトル/幾何） =========================

def ensure_folder(path: str) -> None:
    """フォルダが無ければ作成。"""
    Path(path).mkdir(parents=True, exist_ok=True)


def arr3(x: float, y: float, z: float) -> Tuple[float, float, float]:
    return float(x), float(y), float(z)


def v_sub(a: Sequence[float], b: Sequence[float]) -> Tuple[float, float, float]:
    return a[0] - b[0], a[1] - b[1], a[2] - b[2]


def v_dot(a: Sequence[float], b: Sequence[float]) -> float:
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def v_len(a: Sequence[float]) -> float:
    return math.sqrt(v_dot(a, a))


def v_norm(a: Sequence[float]) -> Tuple[float, float, float]:
    n = v_len(a)
    return (a[0] / n, a[1] / n, a[2] / n) if n > 0 else (0.0, 0.0, 0.0)


def dist3(a: Sequence[float], b: Sequence[float]) -> float:
    return v_len(v_sub(a, b))


def proj2d(p: Sequence[float], o: Sequence[float], u: Sequence[float], v: Sequence[float]) -> Tuple[float, float]:
    """
    面の原点o・面内基底ベクトルu,vを使って3D点pをローカル2D座標へ射影する。
    x = (p - o)·u, y = (p - o)·v
    """
    w = v_sub(p, o)
    return v_dot(w, u), v_dot(w, v)


def dist2d(x1: float, y1: float, x2: float, y2: float) -> float:
    return math.hypot(x2 - x1, y2 - y1)


def angle2d(ax: float, ay: float, bx: float, by: float) -> float:
    """
    2Dベクトル(a)と(b)のなす角度（0〜180°）。
    直角判定や半円内チェック用。
    """
    da = math.hypot(ax, ay)
    db = math.hypot(bx, by)
    if da == 0 or db == 0:
        return 0.0
    d = (ax * bx + ay * by) / (da * db)
    d = max(-1.0, min(1.0, d))
    deg = math.degrees(math.atan2(math.sqrt(1 - d * d), d))
    return deg if deg >= 0 else deg + 180.0


def collinear(p0: Sequence[float], pm: Sequence[float], p1: Sequence[float],
              tol_mm: float = 0.01, tol_deg: float = 0.5) -> bool:
    """
    3点が一直線か？
    ∠(p0→pm, pm→p1) ≈ 180° かどうかで判定する。
    長さゼロ近傍は除外。
    """
    a = v_sub(pm, p0)
    b = v_sub(p1, pm)
    la, lb = v_len(a), v_len(b)
    if la < tol_mm or lb < tol_mm:
        return False
    d = (a[0] * b[0] + a[1] * b[1] + a[2] * b[2]) / (la * lb)
    d = max(-1.0, min(1.0, d))
    ang = math.degrees(math.atan2(math.sqrt(1 - d * d), d))
    if ang < 0:
        ang += 180.0
    return abs(ang - 180.0) <= tol_deg


def key_of_point(p: Sequence[float]) -> str:
    """
    端点一致のための量子化キー。
    TOL_JOIN で丸めることで、面取りや数値誤差による微ズレを許容。
    """
    q = TOL_JOIN if TOL_JOIN > 0 else 0.01
    return f"{round(p[0]/q):d}|{round(p[1]/q):d}|{round(p[2]/q):d}"


# ========================= データ構造 =========================

@dataclass
class EdgeData:
    """1本のエッジ（ARC or LINE）の計測結果。"""
    ref: Any                        # Edge の Reference
    type: str                       # "ARC" | "LINE" | "CURVE"
    p0: Tuple[float, float, float]  # start 3D
    pm: Tuple[float, float, float]  # mid 3D
    p1: Tuple[float, float, float]  # end 3D
    length: float                   # 長さ [mm]
    radius: Optional[float] = None  # 半径 [mm]（円弧のみ）
    angle_deg: Optional[float] = None  # 弧角 [deg]（円弧のみ）≈ len/r*180/pi
    center: Optional[Tuple[float, float, float]] = None  # 中心3D（取れない場合あり）


@dataclass
class UPattern:
    """Uパターン（半円弧+直線2）"""
    arc: EdgeData
    l1: EdgeData
    l2: EdgeData
    opening_2d: float                              # Uの開口距離（面2D上での自由端間距離）
    center3d: Tuple[float, float, float]          # U弧の中心（3D）※近似含む


# ========================= CATIA 接続/ラッパ =========================

def connect() -> Tuple[Any, Any]:
    """
    CATIA.Application に接続し、アクティブドキュメントを返す。
    """
    app = win32.Dispatch("CATIA.Application")
    doc = app.ActiveDocument
    if doc is None:
        raise RuntimeError("CATIAで Part / Product を開いた状態で実行してください。")
    return app, doc


def get_spa(doc: Any) -> Any:
    """SPAWorkbench を取得（測定API入り口）。"""
    return doc.GetWorkbench("SPAWorkbench")


# ========================= トポロジ走査：Face / Edge 取得 =========================

def search_faces(part_doc: Any) -> List[Any]:
    """
    ドキュメント内の全フェースを Selection.Search で取得。
    検索構文: "Topology.Face,all"
    """
    sel = part_doc.Selection
    sel.Clear()
    sel.Search("Topology.Face,all")
    return [sel.Item(i).Reference for i in range(1, sel.Count + 1)]


def edges_of_face(part_doc: Any, face_ref: Any) -> List[Any]:
    """
    あるフェースの境界エッジ集合を Selection.Search で取得。
    フェースを選択→ "Topology.Edge,sel" で境界に限定。
    """
    sel = part_doc.Selection
    sel.Clear()
    sel.Add(face_ref)
    sel.Search("Topology.Edge,sel")
    return [sel.Item(i).Reference for i in range(1, sel.Count + 1)]


def get_plane(measurable_face: Any) -> Tuple[Tuple[float, float, float],
                                             Tuple[float, float, float],
                                             Tuple[float, float, float]]:
    """
    Faceの平面情報（原点 + 面内2軸）を取得。
    measurable.GetPlane() は [Ox,Oy,Oz, Ux,Uy,Uz, Vx,Vy,Vz] を返す想定。
    """
    arr = [0.0] * 9
    measurable_face.GetPlane(arr)  # 例外になる面もある（例：曲面）→呼び出し側でtry
    origin = arr3(arr[0], arr[1], arr[2])
    u = v_norm(arr3(arr[3], arr[4], arr[5]))
    v = v_norm(arr3(arr[6], arr[7], arr[8]))
    return origin, u, v


def edge_measure(spa: Any, edge_ref: Any) -> Optional[EdgeData]:
    """
    1本のEdgeから、端点/中点、長さ、半径（あれば）、中心（あれば）を取得。
    半径>0なら円弧と見なし、弧角[deg] ≈ 長さ/半径*180/pi を付与。
    直線は、start-mid-end の一直線性で判断（近似）。
    """
    try:
        m = spa.GetMeasurable(edge_ref)
    except Exception:
        return None

    # 端点・中点
    p = [0.0] * 9
    try:
        m.GetPointsOnCurve(p)  # start(0..2), mid(3..5), end(6..8)
    except Exception:
        return None

    p0 = arr3(p[0], p[1], p[2])
    pm = arr3(p[3], p[4], p[5])
    p1 = arr3(p[6], p[7], p[8])

    # 長さ
    try:
        length = float(m.Length)
    except Exception:
        length = 0.0

    # 半径（取れれば円弧）
    radius = None
    angle_deg = None
    edge_type = "CURVE"
    center = None

    try:
        r = float(m.Radius)  # 円/円弧/円筒のとき取れることがある
        if r > 0:
            radius = r
            edge_type = "ARC"
            angle_deg = (length / r) * (180.0 / math.pi)
            # 中心（取得できる場合のみ）
            cc = [0.0] * 3
            try:
                m.GetCenter(cc)
                center = arr3(cc[0], cc[1], cc[2])
            except Exception:
                center = None
    except Exception:
        # 半径が取れない → 直線かその他
        if collinear(p0, pm, p1, tol_mm=0.01, tol_deg=0.5):
            edge_type = "LINE"
        else:
            edge_type = "CURVE"

    return EdgeData(
        ref=edge_ref,
        type=edge_type,
        p0=p0,
        pm=pm,
        p1=p1,
        length=length,
        radius=radius,
        angle_deg=angle_deg,
        center=center,
    )


# ========================= U/I 検出ロジック（面単位） =========================

def build_endmap(edges: List[EdgeData]) -> Dict[str, List[EdgeData]]:
    """
    端点量子化キー -> その端点を持つエッジ群 のインデックス。
    U弧端に接続する直線の探索を高速化。
    """
    mp: Dict[str, List[EdgeData]] = {}
    for e in edges:
        for pt in (e.p0, e.p1):
            k = key_of_point(pt)
            mp.setdefault(k, []).append(e)
    return mp


def lines_at_key(endmap: Dict[str, List[EdgeData]], key: str) -> List[EdgeData]:
    """端点キー key を共有する LINE エッジだけ抽出。"""
    return [e for e in endmap.get(key, []) if e.type == "LINE"]


def ensure_center3d(arc: EdgeData, origin: Tuple[float, float, float],
                    u: Tuple[float, float, float],
                    v: Tuple[float, float, float]) -> Tuple[float, float, float]:
    """
    円弧中心が取得できない場合、面2Dに射影して端点の二等分点から近似。
    """
    if arc.center is not None:
        return arc.center
    p0x, p0y = proj2d(arc.p0, origin, u, v)
    p1x, p1y = proj2d(arc.p1, origin, u, v)
    cx2 = 0.5 * (p0x + p1x)
    cy2 = 0.5 * (p0y + p1y)
    # 2D→3D復元： origin + cx2*u + cy2*v
    return (
        origin[0] + cx2 * u[0] + cy2 * v[0],
        origin[1] + cx2 * u[1] + cy2 * v[1],
        origin[2] + cx2 * u[2] + cy2 * v[2],
    )


def free_end2d(line: EdgeData, u_end3: Tuple[float, float, float],
               origin: Tuple[float, float, float],
               u: Tuple[float, float, float],
               v: Tuple[float, float, float]) -> Tuple[float, float]:
    """
    U弧の端点（3D）に接続している line の「自由端」（反対側端点）を面2Dで返す。
    """
    k_line_p0 = key_of_point(line.p0)
    k_u_end = key_of_point(u_end3)
    tgt3 = line.p1 if k_line_p0 == k_u_end else line.p0
    return proj2d(tgt3, origin, u, v)


def free_end2d_against(line: EdgeData, lb: EdgeData,
                       origin: Tuple[float, float, float],
                       u: Tuple[float, float, float],
                       v: Tuple[float, float, float]) -> Tuple[float, float]:
    """
    line が中央線 Lb と接している前提で、line の自由端を2Dで返す。
    """
    k_lb0 = key_of_point(lb.p0)
    k_lb1 = key_of_point(lb.p1)
    tgt3 = line.p1 if key_of_point(line.p0) in (k_lb0, k_lb1) else line.p0
    return proj2d(tgt3, origin, u, v)


def semicircle_bisector(arc: EdgeData,
                        origin: Tuple[float, float, float],
                        u: Tuple[float, float, float],
                        v: Tuple[float, float, float]
                        ) -> Tuple[float, float, float, float]:
    """
    半円の「内側」を決めるための半平面定義を返す。
    - 中心C(2D) = arc中心(3D)を2Dへ射影
    - 端点ベクトル C→P0, C→P1 の二等分ベクトル b を計算（正規化）
    内側条件: (P - C)·b >= 0 かつ |P-C| <= r
    """
    cx, cy = proj2d(ensure_center3d(arc, origin, u, v), origin, u, v)
    p0x, p0y = proj2d(arc.p0, origin, u, v)
    p1x, p1y = proj2d(arc.p1, origin, u, v)
    v0x, v0y = p0x - cx, p0y - cy
    v1x, v1y = p1x - cx, p1y - cy
    bx, by = v0x + v1x, v0y + v1y
    n = math.hypot(bx, by)
    if n == 0:
        bx, by = 1.0, 0.0
    else:
        bx, by = bx / n, by / n
    return cx, cy, bx, by


def in_semicircle(x: float, y: float, cx: float, cy: float,
                  bx: float, by: float, r: float) -> bool:
    """
    点(x,y)が半径rの円の内側、かつ二等分ベクトルによる半平面の内側にあるか。
    """
    dx, dy = x - cx, y - cy
    if dx * dx + dy * dy > (r + 1e-3) ** 2:
        return False
    return (dx * bx + dy * by) >= -1e-3


def right_neighbors(lines: List[EdgeData], lb: EdgeData, at_start: bool,
                    origin: Tuple[float, float, float],
                    u: Tuple[float, float, float],
                    v: Tuple[float, float, float],
                    tol_right: float = TOL_RIGHT_DEG) -> List[EdgeData]:
    """
    Lbの片端に直角(90°±tol)で接続する直線を列挙（面2Dで角度判定）。
    """
    b0x, b0y = proj2d(lb.p0, origin, u, v)
    b1x, b1y = proj2d(lb.p1, origin, u, v)
    bx = (b1x - b0x) if at_start else (b0x - b1x)
    by = (b1y - b0y) if at_start else (b0y - b1y)

    key_anchor = key_of_point(lb.p0 if at_start else lb.p1)
    res: List[EdgeData] = []
    for L in lines:
        if L is lb:
            continue
        # 端点共有？
        if key_of_point(L.p0) != key_anchor and key_of_point(L.p1) != key_anchor:
            continue
        p0x, p0y = proj2d(L.p0, origin, u, v)
        p1x, p1y = proj2d(L.p1, origin, u, v)
        ux, uy = (p1x - p0x, p1y - p0y) if key_of_point(L.p0) == key_anchor else (p0x - p1x, p0y - p1y)
        ang = angle2d(bx, by, ux, uy)
        if abs(ang - 90.0) <= tol_right:
            res.append(L)
    return res


def min_distance_between_refs(spa: Any, r1: Any, r2: Any) -> float:
    """
    Measurable.GetMinimumDistance(ref2) を利用して 3D最短距離を返す。
    取得できなければ巨大値。
    """
    try:
        m1 = spa.GetMeasurable(r1)
        d = float(m1.GetMinimumDistance(r2))
        return d
    except Exception:
        return 1e9


# ========================= メイン処理（Part / Product） =========================

def scan_part3d(part_doc: Any, rows: List[List[Any]]) -> None:
    """
    1 PartDocument を全フェース走査 → U/I検出 → rows にレコードを蓄積。
    rowsの1レコード：
      [Doc, Part, FaceId, Pattern, U_R, U_Ang, U_L1, U_L2, Opening,
       I_La, I_Lb, I_Lc, RightOK, U2I_Min, SnapRelPath]
    """
    global _snap_count, _pat_count

    part = part_doc.Part
    spa = get_spa(part_doc)

    face_refs = search_faces(part_doc)
    if not face_refs:
        return

    for f_idx, face_ref in enumerate(face_refs, start=1):
        # 面の平面が取れない（曲面など）はスキップ
        try:
            meas_face = spa.GetMeasurable(face_ref)
            origin, u, v = get_plane(meas_face)
        except Exception:
            continue

        # 面の境界エッジ群を収集
        edge_refs = edges_of_face(part_doc, face_ref)
        edges: List[EdgeData] = []
        for er in edge_refs:
            ed = edge_measure(spa, er)
            if ed:
                edges.append(ed)
        if not edges:
            continue

        # U候補：半径/弧角条件を満たす円弧 + 端点接続の直線2本
        arcs = [e for e in edges if e.type == "ARC" and e.radius and e.angle_deg is not None
                and (R_MIN <= e.radius <= R_MAX) and (abs(e.angle_deg - 180.0) <= TOL_DEG_ARC)]
        lines = [e for e in edges if e.type == "LINE"]
        if not arcs or not lines:
            continue

        endmap = build_endmap(edges)
        u_candidates: List[UPattern] = []

        for ua in arcs:
            k0 = key_of_point(ua.p0)
            k1 = key_of_point(ua.p1)
            cand_l1 = lines_at_key(endmap, k0)
            cand_l2 = lines_at_key(endmap, k1)
            if not cand_l1 or not cand_l2:
                continue

            # ここでは最初の組を採用（必要なら最適選択に拡張可）
            l1 = cand_l1[0]
            l2 = cand_l2[0]

            # U開口：面2Dで自由端どうしの距離
            l1_free_x, l1_free_y = free_end2d(l1, ua.p0, origin, u, v)
            l2_free_x, l2_free_y = free_end2d(l2, ua.p1, origin, u, v)
            opening = dist2d(l1_free_x, l1_free_y, l2_free_x, l2_free_y)

            c3 = ensure_center3d(ua, origin, u, v)
            u_candidates.append(UPattern(arc=ua, l1=l1, l2=l2, opening_2d=opening, center3d=c3))

        if not u_candidates:
            continue

        # U間の最小距離（3D中心）で間引き
        accepted_u: List[UPattern] = []
        anchors: List[Tuple[float, float, float]] = []
        for up in u_candidates:
            if all(dist3(up.center3d, a) >= MIN_SPACING3D for a in anchors):
                accepted_u.append(up)
                anchors.append(up.center3d)

        if not accepted_u:
            continue

        # I検出（面2Dで直角接続 + U半円内の半平面条件）
        for up in accepted_u:
            ua, ul1, ul2 = up.arc, up.l1, up.l2
            # 半円内半平面ベクトル
            cx2, cy2, bx, by = semicircle_bisector(ua, origin, u, v)
            r = ua.radius if ua.radius else 0.0

            found_i = False
            for lb in lines:
                # 中央線Lbの両端が半円内か？
                lb0x, lb0y = proj2d(lb.p0, origin, u, v)
                lb1x, lb1y = proj2d(lb.p1, origin, u, v)
                if not (in_semicircle(lb0x, lb0y, cx2, cy2, bx, by, r) and
                        in_semicircle(lb1x, lb1y, cx2, cy2, bx, by, r)):
                    continue

                # 直角接続する相手線（両端）
                s_ng = right_neighbors(lines, lb, at_start=True, origin=origin, u=u, v=v,
                                       tol_right=TOL_RIGHT_DEG)
                e_ng = right_neighbors(lines, lb, at_start=False, origin=origin, u=u, v=v,
                                       tol_right=TOL_RIGHT_DEG)
                if not s_ng or not e_ng:
                    continue

                la = s_ng[0]
                lc = e_ng[0]

                # 自由端も半円内か
                afx, afy = free_end2d_against(la, lb, origin, u, v)
                cfx, cfy = free_end2d_against(lc, lb, origin, u, v)
                if not (in_semicircle(afx, afy, cx2, cy2, bx, by, r) and
                        in_semicircle(cfx, cfy, cx2, cy2, bx, by, r)):
                    continue

                # U直線2本 ↔ I直線3本 の3D最短距離の最小
                spa = get_spa(part_doc)
                dmin = min(
                    min(
                        min_distance_between_refs(spa, up.l1.ref, la.ref),
                        min_distance_between_refs(spa, up.l1.ref, lb.ref),
                        min_distance_between_refs(spa, up.l1.ref, lc.ref),
                    ),
                    min(
                        min_distance_between_refs(spa, up.l2.ref, la.ref),
                        min_distance_between_refs(spa, up.l2.ref, lb.ref),
                        min_distance_between_refs(spa, up.l2.ref, lc.ref),
                    ),
                )

                # スナップ撮影（間引きフラグで制御）
                snap_rel = f"{SNAP_DIRNAME}/{sanitize(u_id(part_doc, f_idx, ua, up))}__{sanitize(i_id(part_doc, f_idx, lb, la, lc))}.png"
                snap_abs = str(Path(OUT_DIR, snap_rel))
                if should_snap():
                    capture_snapshot(part_doc, [ua.ref, ul1.ref, ul2.ref, la.ref, lb.ref, lc.ref], snap_abs)

                # レコード追加
                rows.append([
                    part_doc.Name, part.Name, f"Face#{f_idx}", "U+I",
                    round(ua.radius or 0.0, 3), round(ua.angle_deg or 0.0, 2),
                    round(ul1.length, 3), round(ul2.length, 3), round(up.opening_2d, 3),
                    round(la.length, 3), round(lb.length, 3), round(lc.length, 3), "OK",
                    round(dmin, 3), snap_rel
                ])
                found_i = True
                break  # このUでは最初のIのみ採用

            if not found_i:
                # U単独でレコード化（参考）
                snap_rel = f"{SNAP_DIRNAME}/{sanitize(u_id(part_doc, f_idx, ua, up))}.png"
                snap_abs = str(Path(OUT_DIR, snap_rel))
                if should_snap():
                    capture_snapshot(part_doc, [ua.ref, ul1.ref, ul2.ref], snap_abs)

                rows.append([
                    part_doc.Name, part.Name, f"Face#{f_idx}", "U",
                    round(ua.radius or 0.0, 3), round(ua.angle_deg or 0.0, 2),
                    round(ul1.length, 3), round(ul2.length, 3), round(up.opening_2d, 3),
                    "", "", "", "", "", snap_rel
                ])


def scan_product3d(prod_doc: Any, rows: List[List[Any]]) -> None:
    """ProductDocument 配下の Part を順次処理。"""
    prod = prod_doc.Product
    for i in range(1, prod.Products.Count + 1):
        child = prod.Products.Item(i)
        ref_doc = child.ReferenceProduct.Parent
        if getattr(ref_doc, "Type", "") == "Part":
            scan_part3d(ref_doc, rows)


# ========================= スナップ & レポート出力 =========================

def should_snap() -> bool:
    """
    スナップ撮影の間引き制御：
    - ENABLE_SNAPS が False → 撮らない
    - SNAP_MAX に達した → 撮らない
    - SNAP_EVERY_N の間引き → 例えばN=2なら偶数回のみ撮る
    """
    global _snap_count, _pat_count
    _pat_count += 1
    if not ENABLE_SNAPS:
        return False
    if _snap_count >= SNAP_MAX:
        return False
    if (_pat_count % SNAP_EVERY_N) != 0:
        return False
    return True


def capture_snapshot(part_doc: Any, refs: List[Any], png_path: str) -> None:
    """
    選択→ReframeOnSelection（失敗時はReframe）→ CapturePictureFile。
    画像保存フォルダは事前に作成しておく。
    """
    global _snap_count
    sel = part_doc.Selection
    sel.Clear()
    for r in refs:
        if r is not None:
            sel.Add(r)
    viewer = part_doc.Application.ActiveWindow.ActiveViewer
    try:
        viewer.ReframeOnSelection()
    except Exception:
        viewer.Reframe()
    ensure_folder(str(Path(png_path).parent))
    part_doc.Application.ActiveWindow.CapturePictureFile(png_path, "png")
    sel.Clear()
    _snap_count += 1


def sanitize(s: str) -> str:
    """ファイル名に使えない/困る文字を安全化。"""
    for ch in r':|\/*?"<>& ':
        s = s.replace(ch, "_")
    return s


def u_id(part_doc: Any, face_idx: int, arc: EdgeData, up: UPattern) -> str:
    """Uパターンの簡易ID（Face番号 + 半径）"""
    r = arc.radius or 0.0
    return f"U_F{face_idx}_R{r:.3f}"


def i_id(part_doc: Any, face_idx: int, lb: EdgeData, la: EdgeData, lc: EdgeData) -> str:
    """Iパターンの簡易ID。必要なら詳細に拡張可能。"""
    return f"I_F{face_idx}"


def build_html_report(rows: List[List[Any]], html_path: str) -> None:
    """
    HTMLレポートを生成。
    1件ごとに左：明細表、右：サムネイル画像の2カラムカードで並べる。
    """
    html = []
    html.append("<!DOCTYPE html><html><head><meta charset='utf-8'><title>U/I 3D Report</title>")
    html.append(
        "<style>"
        "body{font-family:Segoe UI,Arial,sans-serif;margin:16px}"
        ".row{display:grid;grid-template-columns:1fr 320px;gap:16px;margin:12px 0;padding:8px;"
        "border:1px solid #e5e5e5;border-radius:8px}"
        "table{border-collapse:collapse;width:100%}"
        "th,td{border:1px solid #ccc;padding:6px 8px;font-size:12px}"
        "th{background:#f5f5f5;position:sticky;top:0}"
        ".snap{border:1px solid #ddd;border-radius:6px;padding:6px;text-align:center}"
        "</style></head><body>"
    )
    html.append(f"<h2>U/I 3D Report</h2><p>Total: {len(rows)}</p>")

    def tr(k: str, v: Any) -> str:
        return f"<tr><th>{k}</th><td>{v}</td></tr>"

    for r in rows:
        # r = [Doc, Part, FaceId, Pattern, U_R, U_Ang, U_L1, U_L2, Opening,
        #      I_La, I_Lb, I_Lc, RightOK, U2I_Min, SnapRelPath]
        snap_rel = r[14] if len(r) > 14 else ""
        html.append("<div class='row'>")
        html.append("<table><tbody>")
        html.append(tr("Doc", r[0]) + tr("Part", r[1]) + tr("Face", r[2]) + tr("Pattern", r[3]))
        html.append(
            tr("U_R [mm]", r[4]) + tr("U_Ang [deg]", r[5]) +
            tr("U_L1 [mm]", r[6]) + tr("U_L2 [mm]", r[7]) + tr("U_Opening [mm]", r[8])
        )
        html.append(
            tr("I_La [mm]", r[9]) + tr("I_Lb [mm]", r[10]) +
            tr("I_Lc [mm]", r[11]) + tr("Right90°", r[12]) + tr("U↔I MinDist [mm]", r[13])
        )
        html.append("</tbody></table>")

        if snap_rel:
            img_path_web = snap_rel.replace("\\", "/")
            html.append(
                f"<div class='snap'><img src='{img_path_web}' "
                f"style='max-width:300px;height:{SNAP_HEIGHT_PX}px;object-fit:contain'><br>"
                f"<small>{img_path_web}</small></div>"
            )
        else:
            html.append("<div class='snap'><em>no snapshot</em></div>")
        html.append("</div>")  # .row

    html.append("</body></html>")

    out = Path(html_path)
    ensure_folder(str(out.parent))
    out.write_text("\n".join(html), encoding="utf-16")  # 日本語も安全に表示


# ========================= エントリーポイント =========================

def main() -> None:
    if not sys.platform.startswith("win"):
        print("⚠️ 本スクリプトは Windows + CATIA + pywin32 前提です。")
    ensure_folder(OUT_DIR)
    ensure_folder(str(Path(OUT_DIR, SNAP_DIRNAME)))

    app, doc = connect()
    rows: List[List[Any]] = []

    doc_type = getattr(doc, "Type", "")
    if doc_type == "Part":
        scan_part3d(doc, rows)
    elif doc_type == "Product":
        scan_product3d(doc, rows)
    else:
        raise RuntimeError("Part または Product を開いてください。")

    html_path = str(Path(OUT_DIR, "Report.html"))
    build_html_report(rows, html_path)
    print(f"完了: {len(rows)} 件 → {html_path}")


if __name__ == "__main__":
    main()
