' Language="VBSCRIPT"
' ============================================================
' CATScript: Uturn(180°arc + 2 lines) & Iturn(inside U; 3 lines with two right angles)
' Extract → Measure → Excel(.xlsx)
' - U arc radius: 1–2 mm
' - Min spacing between accepted patterns: 100 mm (anchor = U arc center)
' - Per-pattern snapshot: Select items → ReframeOnSelection → Capture
' - Single Excel sheet "Report": left=Results table, right=Snapshots
' ============================================================

Option Explicit

Sub CATMain()
  Dim app : Set app = CATIA
  If app Is Nothing Then MsgBox "CATIAが見つかりません": Exit Sub

  Dim doc : Set doc = app.ActiveDocument
  If doc Is Nothing Then MsgBox "ドキュメントを開いてください": Exit Sub

  ' ===== 設定 =====
  Dim NameLikeU : NameLikeU = "Uturn*"    ' フィルタ（空で無視）
  Dim NameLikeI : NameLikeI = "Iturn*"
  Dim Rmin      : Rmin      = 1#
  Dim Rmax      : Rmax      = 2#
  Dim TolAngDeg : TolAngDeg = 0.5         ' 180°±0.5°
  Dim TolXY     : TolXY     = 0.01        ' 端点一致 tol
  Dim TolRight  : TolRight  = 3#          ' 直角 ±3°
  Dim MinSpacing: MinSpacing= 100#        ' mm（U円弧中心で判定）
  Dim OutPath   : OutPath   = "U_I_report.xlsx"
  Dim SnapDir   : SnapDir   = "ui_snaps"  ' 画像保存先
  Dim SnapshotHeight : SnapshotHeight : SnapshotHeight = 240 ' Excel貼付時の高さ(px)
  ' ===============

  EnsureFolder SnapDir

  ' 計測行（左表）
  Dim rows : Set rows = CreateObject("Scripting.Dictionary")
  ' 列: [Doc,Part,Body,Sketch,Pattern, U_ID,U_R,U_Ang,U_L1,U_L2,U_Opening,
  '      I_ID,I_La,I_Lb,I_Lc,I_RightOK, U2I_L1_Min,U2I_L2_Min,U2I_Min, SnapLabel]

  ' スナップ（右側）：label -> imgPath
  Dim snaps : Set snaps = CreateObject("Scripting.Dictionary")

  ' パターン間隔チェック用アンカー（U arc center）
  Dim anchors : Set anchors = CreateObject("Scripting.Dictionary")

  Select Case TypeName(doc)
    Case "PartDocument"
      ProcessPart doc, NameLikeU, NameLikeI, Rmin, Rmax, TolAngDeg, TolXY, TolRight, MinSpacing, anchors, rows, snaps, SnapDir, SnapshotHeight
    Case "ProductDocument"
      ProcessProduct doc, NameLikeU, NameLikeI, Rmin, Rmax, TolAngDeg, TolXY, TolRight, MinSpacing, anchors, rows, snaps, SnapDir, SnapshotHeight
    Case Else
      MsgBox "Part または Product を開いてください": Exit Sub
  End Select

  ExportExcelOneSheet rows, snaps, OutPath, SnapshotHeight
  MsgBox "完了: 行数=" & rows.Count & vbCrLf & "出力: " & OutPath
End Sub

' ---------------- Product配下 ----------------
Sub ProcessProduct(prodDoc, NameLikeU, NameLikeI, Rmin, Rmax, TolAngDeg, TolXY, TolRight, MinSpacing, anchors, rows, snaps, SnapDir, SnapshotHeight)
  Dim prod : Set prod = prodDoc.Product
  If prod Is Nothing Then Exit Sub

  Dim i
  For i = 1 To prod.Products.Count
    Dim child : Set child = prod.Products.Item(i)
    If Not child Is Nothing Then
      Dim refDoc : Set refDoc = child.ReferenceProduct.Parent
      If Not refDoc Is Nothing Then
        If TypeName(refDoc) = "PartDocument" Then
          ProcessPart refDoc, NameLikeU, NameLikeI, Rmin, Rmax, TolAngDeg, TolXY, TolRight, MinSpacing, anchors, rows, snaps, SnapDir, SnapshotHeight
        End If
      End If
    End If
  Next
End Sub

' ---------------- Part内 ----------------
Sub ProcessPart(partDoc, NameLikeU, NameLikeI, Rmin, Rmax, TolAngDeg, TolXY, TolRight, MinSpacing, anchors, rows, snaps, SnapDir, SnapshotHeight)
  Dim part : Set part = partDoc.Part
  If part Is Nothing Then Exit Sub

  Dim bodies : Set bodies = part.Bodies
  Dim b
  For b = 1 To bodies.Count
    Dim body : Set body = bodies.Item(b)
    Dim sketches
    On Error Resume Next
    Set sketches = body.Sketches
    On Error GoTo 0
    If Not sketches Is Nothing Then
      Dim s
      For s = 1 To sketches.Count
        Dim sk : Set sk = sketches.Item(s)
        AnalyzeSketch partDoc, part, body, sk, NameLikeU, NameLikeI, Rmin, Rmax, TolAngDeg, TolXY, TolRight, MinSpacing, anchors, rows, snaps, SnapDir, SnapshotHeight
      Next
    End If
  Next
End Sub

' ---------------- スケッチ解析（抽出・計測・間隔・スナップ） ----------------
Sub AnalyzeSketch(partDoc, part, body, sk, NameLikeU, NameLikeI, Rmin, Rmax, TolAngDeg, TolXY, TolRight, MinSpacing, anchors, rows, snaps, SnapDir, SnapshotHeight)
  Dim geos : Set geos = sk.GeometricElements
  If geos Is Nothing Then Exit Sub

  Dim useU : useU = (Len(NameLikeU) > 0)
  Dim useI : useI = (Len(NameLikeI) > 0)
  Dim sketchHasU : sketchHasU = (Not useU) Or (SafeLike(sk.Name, NameLikeU))
  Dim sketchHasI : sketchHasI = (Not useI) Or (SafeLike(sk.Name, NameLikeI))

  Dim lines : Set lines = CreateObject("Scripting.Dictionary") ' k -> dict{..., obj}
  Dim arcs  : Set arcs  = CreateObject("Scripting.Dictionary") ' k -> dict{..., obj}

  Dim i
  For i = 1 To geos.Count
    Dim g : Set g = geos.Item(i)
    Dim tn : tn = SafeTypeName(g)

    If tn = "Line2D" Then
      Dim L : Set L = CreateObject("Scripting.Dictionary")
      FillLine2D g, L
      If GetBool(L,"ok") Then
        L("name") = g.Name
        L("obj")  = g
        lines.Add CStr(lines.Count+1), L
      End If

    ElseIf tn = "Arc2D" Or tn = "Circle2D" Then
      Dim A : Set A = CreateObject("Scripting.Dictionary")
      FillArc2D g, A
      If GetBool(A,"ok") Then
        If (A("r") >= Rmin And A("r") <= Rmax) And (Abs(A("degSpan") - 180#) <= TolAngDeg) Then
          A("name") = g.Name
          A("obj")  = g
          arcs.Add CStr(arcs.Count+1), A
        End If
      End If
    End If
  Next

  If arcs.Count = 0 Then Exit Sub

  Dim ak
  For Each ak In arcs.Keys
    Dim U : Set U = arcs(ak)

    ' 名前条件（U）
    Dim namePassU : namePassU
    namePassU = sketchHasU
    If (Not namePassU) And (Len(NameLikeU) > 0) Then namePassU = SafeLike(U("name"), NameLikeU)
    If Not (useU And Not namePassU) Then

      ' U端点と直線
      Dim sx,sy,ex,ey : ArcEndpoints U("cx"),U("cy"),U("r"),U("a0"),U("a1"), sx,sy, ex,ey

      Dim kL1,kL2, L1ox,L1oy,L2ox,L2oy, L1len,L2len, L1name,L2name
      Dim f1 : f1 = FindLineAtPointKey(lines, sx, sy, TolXY, kL1, L1ox, L1oy, L1len, L1name)
      Dim f2 : f2 = FindLineAtPointKey(lines, ex, ey, TolXY, kL2, L2ox, L2oy, L2len, L2name)

      If f1 And f2 Then
        Dim opening : opening = Dist2D(L1ox,L1oy, L2ox,L2oy)
        Dim u_id : u_id = sk.Name & ":U" & ak

        ' ====== 間隔チェック（U arc center） ======
        If IsFarEnough(anchors, U("cx"), U("cy"), MinSpacing) Then

          ' ====== Iturn検出 ======
          Dim foundI : foundI = False
          Dim lk
          For Each lk In lines.Keys
            Dim Lb : Set Lb = lines(lk)
            Dim sNg : Set sNg = RightAngleNeighborsAt(lines, Lb, True,  TolXY, TolRight)
            Dim eNg : Set eNg = RightAngleNeighborsAt(lines, Lb, False, TolXY, TolRight)
            If sNg.Count > 0 And eNg.Count > 0 Then

              Dim sk1, sk2
              For Each sk1 In sNg.Keys
                For Each sk2 In eNg.Keys
                  If sk1 <> sk2 Then
                    Dim La : Set La = lines(sk1)
                    Dim Lc : Set Lc = lines(sk2)

                    ' I主要点がU円内
                    Dim aFx,aFy, cFx,cFy
                    FreeEndOppositeTo La, Lb, TolXY, aFx, aFy
                    FreeEndOppositeTo Lc, Lb, TolXY, cFx, cFy
                    Dim inside : inside = _
                        InsideCircle(U("cx"),U("cy"),U("r")-0.001, Lb("x1"),Lb("y1")) And _
                        InsideCircle(U("cx"),U("cy"),U("r")-0.001, Lb("x2"),Lb("y2")) And _
                        InsideCircle(U("cx"),U("cy"),U("r")-0.001, aFx,aFy) And _
                        InsideCircle(U("cx"),U("cy"),U("r")-0.001, cFx,cFy)

                    If inside Then
                      ' U直線 vs I直線 のセグメント最短距離
                      Dim UL1 : Set UL1 = lines(kL1)
                      Dim UL2 : Set UL2 = lines(kL2)

                      Dim dL1a : dL1a = DistSegToSeg(UL1, La)
                      Dim dL1b : dL1b = DistSegToSeg(UL1, Lb)
                      Dim dL1c : dL1c = DistSegToSeg(UL1, Lc)
                      Dim U2I_L1_Min : U2I_L1_Min = Min3(dL1a, dL1b, dL1c)

                      Dim dL2a : dL2a = DistSegToSeg(UL2, La)
                      Dim dL2b : dL2b = DistSegToSeg(UL2, Lb)
                      Dim dL2c : dL2c = DistSegToSeg(UL2, Lc)
                      Dim U2I_L2_Min : U2I_L2_Min = Min3(dL2a, dL2b, dL2c)

                      Dim U2I_Min : U2I_Min = IIf(U2I_L1_Min < U2I_L2_Min, U2I_L1_Min, U2I_L2_Min)

                      Dim rightOK : rightOK = "OK"
                      Dim i_id : i_id = sk.Name & ":I" & lk & "-" & sk1 & "-" & sk2

                      ' === スナップショット（パターン単位・選択→ReframeOnSelection） ===
                      Dim snapLabel : snapLabel = u_id & " + " & i_id
                      Dim imgPath   : imgPath   = SnapDir & "\" & SanitizeFileName(snapLabel) & ".png"
                      CapturePatternSnapshot partDoc, Array(U("obj"), UL1("obj"), UL2("obj"), La("obj"), Lb("obj"), Lc("obj")), imgPath
                      If Not snaps.Exists(snapLabel) Then snaps.Add snapLabel, imgPath

                      ' === 行追加 ===
                      Dim rec
                      rec = Array( _
                        partDoc.Name, part.Name, body.Name, sk.Name, "U+I", _
                        u_id, Round(U("r"),3), Round(U("degSpan"),3), Round(L1len,3), Round(L2len,3), Round(opening,3), _
                        i_id, Round(La("len"),3), Round(Lb("len"),3), Round(Lc("len"),3), rightOK, _
                        Round(U2I_L1_Min,3), Round(U2I_L2_Min,3), Round(U2I_Min,3), snapLabel _
                      )
                      rows.Add CStr(rows.Count+1), rec
                      foundI = True
                      AddAnchor anchors, U("cx"), U("cy")
                    End If
                  End If
                Next
              Next

            End If
          Next

          If Not foundI Then
            ' U単独：スナップはU構成要素のみ
            Dim snapLabelU : snapLabelU = u_id
            Dim imgPathU   : imgPathU   = SnapDir & "\" & SanitizeFileName(snapLabelU) & ".png"
            CapturePatternSnapshot partDoc, Array(U("obj"), lines(kL1)("obj"), lines(kL2)("obj")), imgPathU
            If Not snaps.Exists(snapLabelU) Then snaps.Add snapLabelU, imgPathU

            Dim recU
            recU = Array( _
              partDoc.Name, part.Name, body.Name, sk.Name, "U", _
              u_id, Round(U("r"),3), Round(U("degSpan"),3), Round(L1len,3), Round(L2len,3), Round(opening,3), _
              "", "", "", "", "", "", "", "", snapLabelU _
            )
            rows.Add CStr(rows.Count+1), recU
            AddAnchor anchors, U("cx"), U("cy")
          End If

        End If ' spacing

      End If ' f1 & f2

    End If ' name filter
  Next
End Sub

' ---------------- 幾何 & 補助 ----------------
Function GetBool(d, k)
  On Error Resume Next
  GetBool = CBool(d(k))
  If Err.Number <> 0 Then GetBool = False
  On Error GoTo 0
End Function

Function SafeTypeName(o)
  On Error Resume Next
  SafeTypeName = TypeName(o)
  If Err.Number <> 0 Then SafeTypeName = ""
  On Error GoTo 0
End Function

Function SafeLike(s, pat)
  On Error Resume Next
  SafeLike = (s Like pat)
  If Err.Number <> 0 Then SafeLike = False
  On Error GoTo 0
End Function

Sub FillLine2D(line2d, d)
  On Error Resume Next
  Dim p1,p2 : Set p1 = line2d.StartPoint : Set p2 = line2d.EndPoint
  If (p1 Is Nothing) Or (p2 Is Nothing) Then d("ok")=False: Exit Sub
  d("x1")=CDbl(p1.X): d("y1")=CDbl(p1.Y)
  d("x2")=CDbl(p2.X): d("y2")=CDbl(p2.Y)
  d("len")=Dist2D(d("x1"),d("y1"),d("x2"),d("y2"))
  d("ok")=True
  On Error GoTo 0
End Sub

Sub FillArc2D(a, d)
  On Error Resume Next
  Dim c : Set c = a.CenterPoint
  Dim r : r = a.Radius
  Dim a0 : a0 = a.StartAngle
  If Err.Number<>0 Then d("ok")=False: Err.Clear: Exit Sub
  Dim a1 : a1 = a.EndAngle
  If Err.Number<>0 Then d("ok")=False: Err.Clear: Exit Sub
  Dim span : span = NormalizeAngle(a1 - a0)
  d("cx")=CDbl(c.X): d("cy")=CDbl(c.Y)
  d("r") =CDbl(r)
  d("a0")=CDbl(a0): d("a1")=CDbl(a1)
  d("degSpan")=CDbl(span*180#/3.141592653589793#)
  d("ok")=True
  On Error GoTo 0
End Sub

Function NormalizeAngle(ang)
  Const PI=3.141592653589793#
  NormalizeAngle=ang
  Do While NormalizeAngle> PI: NormalizeAngle=NormalizeAngle-2*PI: Loop
  Do While NormalizeAngle<-PI: NormalizeAngle=NormalizeAngle+2*PI: Loop
  NormalizeAngle=Abs(NormalizeAngle)
End Function

Sub ArcEndpoints(cx,cy,r,a0,a1,ByRef sx,ByRef sy,ByRef ex,ByRef ey)
  sx = cx + r * Cos(a0)
  sy = cy + r * Sin(a0)
  ex = cx + r * Cos(a1)
  ey = cy + r * Sin(a1)
End Sub

Function Dist2D(x1,y1,x2,y2)
  Dist2D=Sqr((x2-x1)*(x2-x1)+(y2-y1)*(y2-y1))
End Function

Function Nearly(a,b,tol) : Nearly=(Abs(a-b)<=tol) : End Function

Function InsideCircle(cx,cy,r,x,y)
  InsideCircle=(Dist2D(cx,cy,x,y) <= r)
End Function

Function FindLineAtPointKey(lines, x, y, tol, ByRef keyOut, ByRef ox, ByRef oy, ByRef llen, ByRef lname)
  Dim k
  For Each k In lines.Keys
    Dim L : Set L = lines(k)
    If (Nearly(L("x1"),x,tol) And Nearly(L("y1"),y,tol)) Then
      keyOut = k: ox=L("x2"): oy=L("y2"): llen=L("len"): lname=""
      FindLineAtPointKey = True: Exit Function
    End If
    If (Nearly(L("x2"),x,tol) And Nearly(L("y2"),y,tol)) Then
      keyOut = k: ox=L("x1"): oy=L("y1"): llen=L("len"): lname=""
      FindLineAtPointKey = True: Exit Function
    End If
  Next
  FindLineAtPointKey = False
End Function

Function RightAngleNeighborsAt(lines, Lb, atStart, tolXY, tolRight)
  Dim res : Set res = CreateObject("Scripting.Dictionary")
  Dim bx,by, vx,vy
  If atStart Then
    bx=Lb("x1"): by=Lb("y1"): vx=Lb("x2")-Lb("x1"): vy=Lb("y2")-Lb("y1")
  Else
    bx=Lb("x2"): by=Lb("y2"): vx=Lb("x1")-Lb("x2"): vy=Lb("y1")-Lb("y2")
  End If
  Dim k
  For Each k In lines.Keys
    Dim L : Set L = lines(k)
    If Not (L Is Lb) Then
      Dim touch : touch=False
      Dim ux,uy
      If Nearly(L("x1"),bx,tolXY) And Nearly(L("y1"),by,tolXY) Then
        ux=L("x2")-L("x1"): uy=L("y2")-L("y1"): touch=True
      ElseIf Nearly(L("x2"),bx,tolXY) And Nearly(L("y2"),by,tolXY) Then
        ux=L("x1")-L("x2"): uy=L("y1")-L("y2"): touch=True
      End If
      If touch Then
        Dim ang : ang = AngleBetween(vx,vy, ux,uy)
        If Abs(ang-90#) <= tolRight Then res.Add k, True
      End If
    End If
  Next
  Set RightAngleNeighborsAt = res
End Function

Function AngleBetween(ax,ay,bx,by)
  Dim da : da=Sqr(ax*ax+ay*ay) : If da=0 Then AngleBetween=0: Exit Function
  Dim db : db=Sqr(bx*bx+by*by) : If db=0 Then AngleBetween=0: Exit Function
  Dim d  : d =(ax*bx+ay*by)/(da*db)
  If d>1 Then d=1
  If d<-1 Then d=-1
  AngleBetween = Atn(Sqr(1-d*d)/d)*180#/3.141592653589793#
  If AngleBetween<0 Then AngleBetween=AngleBetween+180#
End Function

Sub FreeEndOppositeTo(L, Lb, tol, ByRef fx, ByRef fy)
  If (Nearly(L("x1"),Lb("x1"),tol) And Nearly(L("y1"),Lb("y1"),tol)) _
     Or (Nearly(L("x1"),Lb("x2"),tol) And Nearly(L("y1"),Lb("y2"),tol)) Then
    fx=L("x2"): fy=L("y2")
  Else
    fx=L("x1"): fy=L("y1")
  End If
End Sub

Function DistSegToSeg(LA, LB)
  Dim d1 : d1 = DistPtToSeg(LA("x1"),LA("y1"), LB("x1"),LB("y1"), LB("x2"),LB("y2"))
  Dim d2 : d2 = DistPtToSeg(LA("x2"),LA("y2"), LB("x1"),LB("y1"), LB("x2"),LB("y2"))
  Dim d3 : d3 = DistPtToSeg(LB("x1"),LB("y1"), LA("x1"),LA("y1"), LA("x2"),LA("y2"))
  Dim d4 : d4 = DistPtToSeg(LB("x2"),LB("y2"), LA("x1"),LA("y1"), LA("x2"),LA("y2"))
  DistSegToSeg = Min4(d1,d2,d3,d4)
End Function

Function DistPtToSeg(px,py, x1,y1,x2,y2)
  Dim vx,vy : vx = x2 - x1 : vy = y2 - y1
  Dim wx,wy : wx = px - x1 : wy = py - y1
  Dim vv : vv = vx*vx + vy*vy
  If vv = 0 Then
    DistPtToSeg = Sqr((px-x1)*(px-x1) + (py-y1)*(py-y1))
    Exit Function
  End If
  Dim t : t = (wx*vx + wy*vy) / vv
  If t < 0 Then t = 0
  If t > 1 Then t = 1
  Dim cx,cy : cx = x1 + t*vx : cy = y1 + t*vy
  DistPtToSeg = Sqr((px-cx)*(px-cx) + (py-cy)*(py-cy))
End Function

Function Min3(a,b,c)
  Dim m : m=a : If b<m Then m=b : If c<m Then m=c
  Min3 = m
End Function

Function Min4(a,b,c,d)
  Dim m : m=a : If b<m Then m=b : If c<m Then m=c : If d<m Then m=d
  Min4 = m
End Function

' ---- パターン間隔管理 ----
Function IsFarEnough(anchors, cx, cy, minDist)
  Dim k
  For Each k In anchors.Keys
    Dim s : s = anchors(k)
    Dim p : p = Split(s, "|")
    Dim ax : ax = CDbl(p(0))
    Dim ay : ay = CDbl(p(1))
    If Dist2D(ax,ay,cx,cy) < minDist Then
      IsFarEnough = False : Exit Function
    End If
  Next
  IsFarEnough = True
End Function

Sub AddAnchor(anchors, cx, cy)
  anchors.Add CStr(anchors.Count+1), CStr(cx) & "|" & CStr(cy)
End Sub

' ---- パターン単位スナップショット ----
Sub CapturePatternSnapshot(partDoc, objs, pngPath)
  On Error Resume Next
  Dim sel : Set sel = partDoc.Selection
  Dim part : Set part = partDoc.Part
  sel.Clear

  Dim i
  For i = LBound(objs) To UBound(objs)
    If Not (objs(i) Is Nothing) Then
      Dim r : Set r = part.CreateReferenceFromObject(objs(i))
      If Not r Is Nothing Then sel.Add r
    End If
  Next

  Dim v : Set v = CATIA.ActiveWindow.ActiveViewer
  v.ReframeOnSelection
  If Err.Number <> 0 Then
    Err.Clear
    v.Reframe
  End If

  CATIA.ActiveWindow.CapturePictureFile pngPath, "png"
  sel.Clear
  On Error GoTo 0
End Sub

Function SanitizeFileName(s)
  Dim t : t = s
  t = Replace(t, ":", "_")
  t = Replace(t, "|", "_")
  t = Replace(t, "\", "_")
  t = Replace(t, "/", "_")
  t = Replace(t, "*", "_")
  t = Replace(t, "?", "_")
  t = Replace(t, """", "_")
  t = Replace(t, "<", "_")
  t = Replace(t, ">", "_")
  t = Replace(t, "&", "_")
  SanitizeFileName = t
End Function

Sub EnsureFolder(path)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

' ---------------- Excel出力（1シート左右並置） ----------------
Sub ExportExcelOneSheet(rows, snaps, outPath, SnapshotHeight)
  Dim xl : Set xl = CreateObject("Excel.Application")
  xl.Visible = False
  Dim wb : Set wb = xl.Workbooks.Add()
  Dim sh : Set sh = wb.Sheets(1)
  sh.Name = "Report"

  ' ---- 左：Results ----
  Dim headers
  headers = Array( _
    "Doc","Part","Body","Sketch","Pattern", _
    "U_ID","U_R(mm)","U_Ang(deg)","U_L1(mm)","U_L2(mm)","U_Opening(mm)", _
    "I_ID","I_La(mm)","I_Lb(mm)","I_Lc(mm)","I_RightOK", _
    "U2I_L1_Min(mm)","U2I_L2_Min(mm)","U2I_Min(mm)","SnapLabel" _
  )

  Dim c
  For c = 0 To UBound(headers)
    sh.Cells(1, c+1).Value = headers(c)
  Next

  Dim r, rec, rowIdx : rowIdx = 2
  For Each r In rows.Keys
    rec = rows(r)
    For c = 0 To UBound(headers)
      If c <= UBound(rec) Then sh.Cells(rowIdx, c+1).Value = rec(c)
    Next
    rowIdx = rowIdx + 1
  Next

  sh.Columns("A:T").AutoFit

  ' ---- 右：Snapshots（同一シート右側に配置） ----
  Dim startCol : startCol = 22  ' 列V以降に配置（A=1, T=20, U=21, V=22）
  sh.Cells(1, startCol).Value     = "Pattern (SnapLabel)"
  sh.Cells(1, startCol + 1).Value = "Image"
  sh.Columns(startCol).ColumnWidth = 48
  sh.Columns(startCol + 1).ColumnWidth = 48

  Dim idx : idx = 2
  Dim k
  For Each k In snaps.Keys
    sh.Cells(idx, startCol).Value = k
    Dim path : path = snaps(k)
    On Error Resume Next
    Dim pic : Set pic = sh.Pictures().Insert(path)
    If Err.Number = 0 Then
      pic.Top  = sh.Cells(idx, startCol + 1).Top
      pic.Left = sh.Cells(idx, startCol + 1).Left
      If pic.Height > SnapshotHeight Then
        Dim scale : scale = SnapshotHeight / pic.Height
        pic.Height = SnapshotHeight
        pic.Width  = pic.Width * scale
      End If
      idx = idx + Int((pic.Height / sh.Rows(1).Height)) + 2
    Else
      Err.Clear
      idx = idx + 1
    End If
    On Error GoTo 0
  Next

  ' 保存
  On Error Resume Next
  wb.SaveAs outPath, 51   ' xlOpenXMLWorkbook
  If Err.Number <> 0 Then
    Err.Clear
    wb.SaveAs outPath
  End If
  On Error GoTo 0

  wb.Close False
  xl.Quit
End Sub
