' Language: VBScript (CATScript)
' 検査自動化テンプレート
' 1) Publication名で対象形状を抽出
' 2) 断面生成（可能なら実行、失敗しても続行）
' 3) 寸法計測（例：最小距離）＋判定
' 4) 図面に「断面ラベル（簡易）＋判定表」を左右併記
' 5) PDF出力

Option Explicit

'========================
' 設定値（必要に応じて編集）
'========================
Const PUB_A_NAME = "U_TARGET"   ' Publication名（例：U側ターゲット）
Const PUB_B_NAME = "I_TARGET"   ' Publication名（例：I側ターゲット）
Const GAP_MIN_MM = 25.0         ' 合格下限（例：最小離隔 ≧ 25mm）

' 断面の基準面（OriginElementsのXY/YZ/ZXいずれかを使用）
Const SECTION_PLANE = "XY"      ' "XY" / "YZ" / "ZX"
Const DRAW_SHEET_FORMAT = "A4"
Const PDF_OUT_PATH = "C:\temp\inspection_report.pdf"  ' 出力先（存在する書き込み可ディレクトリ）

'========================
' メイン
'========================
Sub CATMain()
    On Error Resume Next

    If CATIA Is Nothing Then
        MsgBox "CATIAが起動していません。", vbCritical : Exit Sub
    End If

    If CATIA.Documents.Count = 0 Then
        MsgBox "アクティブなドキュメントがありません。Partを開いてから実行してください。", vbCritical : Exit Sub
    End If

    Dim doc : Set doc = CATIA.ActiveDocument
    If LCase(doc.Type) <> "part" Then
        MsgBox "PartDocument上で実行してください。現在: " & doc.Type, vbCritical : Exit Sub
    End If

    Dim partDoc : Set partDoc = doc
    Dim part    : Set part    = partDoc.Part

    '---- Publicationから対象形状を取得
    Dim refA : Set refA = GetRefFromPublication(part, PUB_A_NAME)
    Dim refB : Set refB = GetRefFromPublication(part, PUB_B_NAME)
    If (refA Is Nothing) Or (refB Is Nothing) Then
        MsgBox "Publicationが見つかりません。" & vbCrLf & _
               "A=" & PUB_A_NAME & ", B=" & PUB_B_NAME & vbCrLf & _
               "PartのPublicationに登録してください。", vbCritical
        Exit Sub
    End If

    '---- 断面（可能なら）作成：失敗しても計測に進む
    Dim sectionOk : sectionOk = CreateSectionCurveIfPossible(part, SECTION_PLANE)
    part.Update

    '---- 計測（例：A-Bの最小距離）
    Dim gapMM, p1, p2
    gapMM = 0 : Set p1 = Nothing : Set p2 = Nothing
    gapMM = GetMinimumDistance(part, refA, refB, p1, p2)

    '---- 判定
    Dim passFail, reason
    If gapMM >= GAP_MIN_MM Then
        passFail = "PASS"
        reason   = "最小離隔が閾値以上"
    Else
        passFail = "FAIL"
        reason   = "最小離隔が閾値未満"
    End If

    '---- 図面作成（左：断面ラベル／右：判定表）→ PDF出力
    Dim pdfOk : pdfOk = BuildDrawingAndExportPDF(partDoc, gapMM, passFail, reason, sectionOk, PDF_OUT_PATH)
    If pdfOk Then
        MsgBox "完了: " & PDF_OUT_PATH, vbInformation
    Else
        MsgBox "図面出力に失敗しました。保存先や権限を確認してください。", vbExclamation
    End If

End Sub

'========================
' Publication → Reference
'========================
Function GetRefFromPublication(part, pubName)
    On Error Resume Next
    Dim pubs : Set pubs = part.Publications
    Dim pub  : Set pub  = pubs.Item(pubName)
    If pub Is Nothing Then
        Set GetRefFromPublication = Nothing
        Exit Function
    End If
    Set GetRefFromPublication = pub.Valuation
End Function

'========================
' 最小距離（SPAWorkbench）
'========================
Function GetMinimumDistance(part, ref1, ref2, p1, p2)
    On Error Resume Next
    Dim spa : Set spa = part.Parent.GetWorkbench("SPAWorkbench")
    Dim meas1 : Set meas1 = spa.GetMeasurable(ref1)
    Dim meas2 : Set meas2 = spa.GetMeasurable(ref2)

    Dim d, x1, y1, z1, x2, y2, z2
    d = 0 : x1=0: y1=0: z1=0: x2=0: y2=0: z2=0

    ' GetMinimumDistanceは2つのMeasurable/Reference間の最近点距離と最近点座標を返す
    d = meas1.GetMinimumDistance(ref2, x1, y1, z1, x2, y2, z2)

    ' 返却用ポイントは使用しない場合もあるが、将来の寸法配置に活用可
    Set p1 = Array(x1, y1, z1)
    Set p2 = Array(x2, y2, z2)

    GetMinimumDistance = d
End Function

'========================
' 断面生成（可能なら実行）
' GSDのHybridShapeFactoryで原点面とBodyの交線を作成する例
' ※環境によりAPI差異があるため、失敗しても処理続行
'========================
Function CreateSectionCurveIfPossible(part, planeKey)
    On Error Resume Next

    Dim origin : Set origin = part.OriginElements
    Dim basePlane
    If UCase(planeKey) = "XY" Then
        Set basePlane = origin.PlaneXY
    ElseIf UCase(planeKey) = "YZ" Then
        Set basePlane = origin.PlaneYZ
    Else
        Set basePlane = origin.PlaneZX
    End If

    Dim hsf  : Set hsf  = part.HybridShapeFactory
    Dim hBodies : Set hBodies = part.HybridBodies
    Dim tgtBody
    ' 交線を置くジオメトリセット（なければ作成）
    Dim setName : setName = "CS_Section"
    Dim secSet
    On Error Resume Next
    Set secSet = hBodies.Item(setName)
    If secSet Is Nothing Then
        Set secSet = hBodies.Add()
        secSet.Name = setName
    End If

    ' 代表Body（最初のSolid）を対象に交線を作る簡易例
    ' ※製品構成や複数Bodyの場合は適宜拡張
    Dim bodies : Set bodies = part.Bodies
    If bodies.Count = 0 Then
        CreateSectionCurveIfPossible = False
        Exit Function
    End If
    Set tgtBody = bodies.Item(1)

    Dim refPlane : Set refPlane = part.CreateReferenceFromObject(basePlane)
    Dim refBody  : Set refBody  = part.CreateReferenceFromObject(tgtBody)

    ' 平面×ソリッドの交線（HybridShapeSection／Intersection系）
    Dim secFeat
    On Error Resume Next

    ' 1) Try: AddNewSection（環境により利用不可な場合あり）
    Set secFeat = hsf.AddNewSection(refBody, refPlane)
    If Err.Number <> 0 Or secFeat Is Nothing Then
        Err.Clear
        ' 2) Fallback: AddNewIntersection（こちらも環境差あり）
        Set secFeat = hsf.AddNewIntersection(refBody, refPlane)
    End If

    If Err.Number <> 0 Or secFeat Is Nothing Then
        Err.Clear
        CreateSectionCurveIfPossible = False
        Exit Function
    End If

    secSet.AppendHybridShape secFeat
    part.InWorkObject = secFeat
    part.Update

    CreateSectionCurveIfPossible = True
End Function

'========================
' 図面作成＋表記入＋PDF出力
' 左側：断面ラベル（簡易）
' 右側：判定表
'========================
Function BuildDrawingAndExportPDF(partDoc, gapMM, passFail, reason, sectionOk, pdfPath)
    On Error Resume Next

    Dim drw : Set drw = CATIA.Documents.Add("Drawing")
    If drw Is Nothing Then
        BuildDrawingAndExportPDF = False
        Exit Function
    End If

    Dim sheet : Set sheet = drw.Sheets.Item(1)
    sheet.PaperSize = DRAW_SHEET_FORMAT

    ' ビュー（左）作成
    Dim views : Set views = sheet.Views
    Dim vLeft : Set vLeft = views.Add("View_Left")
    vLeft.x = 30  ' 左余白
    vLeft.y = 180 ' 上からの位置（mm単位）
    vLeft.Scale = 1

    ' 断面ラベル（ここに断面ビュー生成を後で置換可能）
    Dim txts : Set txts = vLeft.Texts
    Dim sLabel
    If sectionOk Then
        sLabel = "断面: " & SECTION_PLANE & "（交線生成済）"
    Else
        sLabel = "断面: " & SECTION_PLANE & "（交線生成不可／環境未対応）"
    End If
    Call txts.Add(sLabel, 10, 140)
    Call txts.Add("寸法（例：A-B最小離隔）: " & FormatNumber(gapMM, 3) & " mm", 10, 125)

    ' ビュー（右）に表を作成
    Dim vRight : Set vRight = views.Add("View_Right")
    vRight.x = 140
    vRight.y = 180
    vRight.Scale = 1

    Dim tables : Set tables = sheet.Tables
    Dim tbl : Set tbl = tables.Add(135, 260, 4, 3, 35, 10) ' (X, Y, 行, 列, 列幅, 行高)

    ' ヘッダ
    Call tbl.SetCellString(1, 1, "検査項目")
    Call tbl.SetCellString(1, 2, "基準")
    Call tbl.SetCellString(1, 3, "結果")

    ' 行1：最小離隔
    Call tbl.SetCellString(2, 1, "最小離隔 A-B")
    Call tbl.SetCellString(2, 2, "≧ " & GAP_MIN_MM & " mm")
    Call tbl.SetCellString(2, 3, FormatNumber(gapMM, 3) & " mm")

    ' 行2：判定
    Call tbl.SetCellString(3, 1, "判定")
    Call tbl.SetCellString(3, 2, "—")
    Call tbl.SetCellString(3, 3, passFail)

    ' 行3：理由
    Call tbl.SetCellString(4, 1, "理由")
    Call tbl.SetCellString(4, 2, "—")
    Call tbl.SetCellString(4, 3, reason)

    ' 備考
    Dim tRightTxts : Set tRightTxts = vRight.Texts
    Call tRightTxts.Add("備考: Publication名 A=" & PUB_A_NAME & " / B=" & PUB_B_NAME, 135, 110)

    ' PDF出力
    drw.ExportData pdfPath, "pdf"

    BuildDrawingAndExportPDF = (Err.Number = 0)
End Function
