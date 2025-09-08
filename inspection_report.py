# -*- coding: utf-8 -*-
"""
CATIA 検査自動化テンプレート（Python版 / COM）
- 3Dモデルから Publication 指定の形状を取得
- 可能なら断面交線を生成（失敗しても続行）
- A-B の最小距離を計測し、閾値で判定
- 図面上で「左：断面ラベル（簡易）／右：判定表」を左右併記
- PDF 出力

前提:
  - Windows + CATIA 環境
  - Python: pip install pywin32
  - CATIA を起動し、Part をアクティブにして実行
"""
from __future__ import annotations

import sys
from typing import Tuple, Optional

import win32com.client as win32
from win32com.client import constants


# ========================
# 設定値（必要に応じて編集）
# ========================
PUB_A_NAME = "U_TARGET"          # Publication名（A側）
PUB_B_NAME = "I_TARGET"          # Publication名（B側）
GAP_MIN_MM = 25.0                # 合格下限（最小離隔 ≧ この値）
SECTION_PLANE = "XY"             # "XY" / "YZ" / "ZX"
DRAW_SHEET_FORMAT = "A4"         # A0/A1/A2/A3/A4/A5
PDF_OUT_PATH = r"C:\temp\inspection_report.pdf"  # 出力先（存在する書き込み可ディレクトリ）


def get_catia():
    """起動中の CATIA.Application に接続（Dispatch）。"""
    try:
        return win32.Dispatch("CATIA.Application")
    except Exception as e:
        raise RuntimeError("CATIA に接続できません。CATIA を起動してください。") from e


def get_active_part_document(catia):
    """アクティブドキュメントが PartDocument かを確認して返す。"""
    if catia.Documents.Count == 0:
        raise RuntimeError("アクティブなドキュメントがありません。Part を開いてから実行してください。")
    doc = catia.ActiveDocument
    if str(doc.Type).lower() != "part":
        raise RuntimeError(f"PartDocument 上で実行してください。現在: {doc.Type}")
    return doc


def get_reference_from_publication(part, pub_name: str):
    """Publication 名から Reference を取得。見つからない場合は None。"""
    pubs = part.Publications
    try:
        pub = pubs.Item(pub_name)
    except Exception:
        return None
    return pub.Valuation  # Reference を返す


def get_origin_plane(part, plane_key: str):
    """原点要素から XY/YZ/ZX の平面オブジェクトを返す。"""
    origin = part.OriginElements
    key = plane_key.strip().upper()
    if key == "XY":
        return origin.PlaneXY
    if key == "YZ":
        return origin.PlaneYZ
    return origin.PlaneZX


def create_section_curve_if_possible(part, plane_key: str) -> bool:
    """
    断面交線をハイブリッドジオメトリセットに作成（可能な範囲）。
    - HybridShapeFactory.AddNewSection を試す → 失敗なら AddNewIntersection を試す。
    - いずれも環境差があるため、失敗しても False を返して続行可能。
    """
    try:
        base_plane = get_origin_plane(part, plane_key)
        hsf = part.HybridShapeFactory
        h_bodies = part.HybridBodies

        # ジオメトリ格納先のセット（なければ作成）
        set_name = "CS_Section"
        try:
            sec_set = h_bodies.Item(set_name)
        except Exception:
            sec_set = h_bodies.Add()
            sec_set.Name = set_name

        bodies = part.Bodies
        if bodies.Count == 0:
            return False
        tgt_body = bodies.Item(1)

        ref_plane = part.CreateReferenceFromObject(base_plane)
        ref_body = part.CreateReferenceFromObject(tgt_body)

        sec_feat = None
        # 1) AddNewSection を試す
        try:
            sec_feat = hsf.AddNewSection(ref_body, ref_plane)
        except Exception:
            sec_feat = None

        # 2) ダメなら AddNewIntersection を試す
        if sec_feat is None:
            try:
                # 引数順は実装により異なることがあるため、両順序を試す
                try:
                    sec_feat = hsf.AddNewIntersection(ref_body, ref_plane)
                except Exception:
                    sec_feat = hsf.AddNewIntersection(ref_plane, ref_body)
            except Exception:
                sec_feat = None

        if sec_feat is None:
            return False

        sec_set.AppendHybridShape(sec_feat)
        part.InWorkObject = sec_feat
        part.Update()
        return True

    except Exception:
        return False


def get_minimum_distance(part, ref1, ref2) -> Tuple[float, Tuple[float, float, float], Tuple[float, float, float]]:
    """
    A-B の最小距離と最近点座標を取得。
    pywin32 では GetMinimumDistance(ref) がタプルを返す実装が一般的:
      (distance, x1, y1, z1, x2, y2, z2)
    """
    spa = part.Parent.GetWorkbench("SPAWorkbench")
    meas1 = spa.GetMeasurable(ref1)

    # 距離と2点座標を一括取得
    distance, x1, y1, z1, x2, y2, z2 = meas1.GetMinimumDistance(ref2)
    p1 = (float(x1), float(y1), float(z1))
    p2 = (float(x2), float(y2), float(z2))
    return float(distance), p1, p2


def _paper_const(fmt: str):
    """紙サイズの定数マッピング（未定義は A4）。"""
    f = fmt.strip().upper()
    return {
        "A0": constants.catPaperA0,
        "A1": constants.catPaperA1,
        "A2": constants.catPaperA2,
        "A3": constants.catPaperA3,
        "A4": constants.catPaperA4,
        "A5": constants.catPaperA5,
    }.get(f, constants.catPaperA4)


def build_drawing_and_export_pdf(
    catia,
    part_doc,
    gap_mm: float,
    pass_fail: str,
    reason: str,
    section_ok: bool,
    pdf_path: str,
    sheet_format: str,
    section_plane: str,
    gap_min_mm: float,
) -> bool:
    """
    図面生成と PDF 出力。
    左ビュー: 断面ラベル（簡易）
    右ビュー: 判定表（4行×3列）
    """
    try:
        drw = catia.Documents.Add("Drawing")
        sheet = drw.Sheets.Item(1)
        sheet.PaperSize = _paper_const(sheet_format)

        # 左側ビュー
        views = sheet.Views
        v_left = views.Add("View_Left")
        v_left.x = 30.0     # 左余白
        v_left.y = 180.0    # 上からの位置（mm）
        v_left.Scale = 1.0

        txts = v_left.Texts
        s_label = (
            f"断面: {section_plane}（交線生成済）"
            if section_ok
            else f"断面: {section_plane}（交線生成不可／環境未対応）"
        )
        txts.Add(s_label, 10.0, 140.0)
        txts.Add(f"寸法（A-B最小離隔）: {gap_mm:.3f} mm", 10.0, 125.0)

        # 右側ビュー（判定表）
        v_right = views.Add("View_Right")
        v_right.x = 140.0
        v_right.y = 180.0
        v_right.Scale = 1.0

        tables = sheet.Tables
        # (X, Y, 行, 列, 列幅, 行高)
        tbl = tables.Add(135.0, 260.0, 4, 3, 35.0, 10.0)

        # ヘッダ
        tbl.SetCellString(1, 1, "検査項目")
        tbl.SetCellString(1, 2, "基準")
        tbl.SetCellString(1, 3, "結果")

        # 行1: 最小離隔
        tbl.SetCellString(2, 1, "最小離隔 A-B")
        tbl.SetCellString(2, 2, f"≧ {gap_min_mm:.3f} mm")
        tbl.SetCellString(2, 3, f"{gap_mm:.3f} mm")

        # 行2: 判定
        tbl.SetCellString(3, 1, "判定")
        tbl.SetCellString(3, 2, "—")
        tbl.SetCellString(3, 3, pass_fail)

        # 行3: 理由
        tbl.SetCellString(4, 1, "理由")
        tbl.SetCellString(4, 2, "—")
        tbl.SetCellString(4, 3, reason)

        # 備考
        v_right.Texts.Add(
            f"備考: Publication名 A={PUB_A_NAME} / B={PUB_B_NAME}",
            135.0,
            110.0,
        )

        # PDF 出力
        drw.ExportData(pdf_path, "pdf")
        return True
    except Exception:
        return False


def main():
    catia = get_catia()
    part_doc = get_active_part_document(catia)
    part = part_doc.Part

    # Publication から対象参照を取得
    ref_a = get_reference_from_publication(part, PUB_A_NAME)
    ref_b = get_reference_from_publication(part, PUB_B_NAME)
    if ref_a is None or ref_b is None:
        raise RuntimeError(
            f"Publication が見つかりません。A='{PUB_A_NAME}', B='{PUB_B_NAME}' を Part の Publications に登録してください。"
        )

    # 断面（可能なら）作成：失敗しても続行
    section_ok = create_section_curve_if_possible(part, SECTION_PLANE)
    part.Update()

    # 計測（最小距離）
    gap_mm, p1, p2 = get_minimum_distance(part, ref_a, ref_b)

    # 判定
    if gap_mm >= GAP_MIN_MM:
        pass_fail = "PASS"
        reason = "最小離隔が閾値以上"
    else:
        pass_fail = "FAIL"
        reason = "最小離隔が閾値未満"

    # 図面生成 + PDF 出力
    ok = build_drawing_and_export_pdf(
        catia=catia,
        part_doc=part_doc,
        gap_mm=gap_mm,
        pass_fail=pass_fail,
        reason=reason,
        section_ok=section_ok,
        pdf_path=PDF_OUT_PATH,
        sheet_format=DRAW_SHEET_FORMAT,
        section_plane=SECTION_PLANE,
        gap_min_mm=GAP_MIN_MM,
    )
    if not ok:
        raise RuntimeError("図面の作成または PDF 出力に失敗しました。保存先や権限を確認してください。")

    print(f"完了: {PDF_OUT_PATH}")
    print(f"最小離隔: {gap_mm:.3f} mm, 判定: {pass_fail}, 理由: {reason}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        sys.exit(1)
