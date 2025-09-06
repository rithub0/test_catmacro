' Language="VBSCRIPT"
' ============================================================
' 3D U/I Extractor → Check → HTML report (+ PNG snapshots)
'   - Input: Part/Product 3D モデル
'   - Find: Uturn = 半円弧(≈180°) + 端点接続の直線×2（同一フェース上）
'           Iturn = U内部のコの字（直角×2の直線×3、同一フェース上）
'   - Check: 半径・直角・内側性（半円内の半平面条件）、U-I距離（最短）
'   - Output: /Report.html（表+サムネ）、/snaps/*.png
' ============================================================

Option Explicit

' ==================== 設定（要件に合わせて調整） ====================
Const R_MIN          = 1.0      ' U弧 半径最小[mm]
Const R_MAX          = 2.0      ' U弧 半径最大[mm]
Const TOL_DEG_ARC    = 1.0      ' 半円判定 180°±[deg]
Const TOL_RIGHT_DEG  = 3.0      ' 直角判定 90°±[deg]
Const TOL_JOIN       = 0.02     ' 端点一致 tol [mm]
Const MIN_SPACING3D  = 100.0    ' 採用パターン間の最小距離(3D, U中心同士) [mm]

Const ENABLE_SNAPS   = True     ' スナップ撮影ON/OFF
Const SNAP_EVERY_N   = 1        ' 何件に1回撮る（1=毎回）
Const SNAP_MAX       = 200      ' 最大撮影枚数
Const SNAP_HEIGHT_PX = 240      ' HTMLサムネ高さ
Const OUT_DIR        = "C:\temp\ui3d_report" ' 出力フォルダ

' ==================== 内部状態 ====================
Dim gSnapCount : gSnapCount = 0
Dim gPatCount  : gPatCount  = 0

' ============================================================
Sub CATMain()

  Dim app : Set app = CATIA
  If app Is Nothing Then MsgBox "CATIAが見つかりません": Exit Sub

  Dim doc : Set doc = app.ActiveDocument
  If doc Is Nothing Then MsgBox "ドキュメントを開いてください": Exit Sub

  EnsureFolder OUT_DIR
  EnsureFolder OUT_DIR & "\snaps"

  Dim rows : Set rows = CreateObject("System.Collections.ArrayList")
  ' rows: Array( Doc, Part, FaceId, Pattern, U_R, U_Ang, U_L1, U_L2, Opening,  _
  '              I_La, I_Lb, I_Lc, RightOK, U2I_Min, SnapRelPath )

  Select Case TypeName(doc)
    Case "PartDocument"
      ScanPart3D doc, rows
    Case "ProductDocument"
      ScanProduct3D doc, rows
    Case Else
      MsgBox "Part または Product を開いてください": Exit Sub
  End Select

  BuildHtmlReport rows, OUT_DIR & "\Report.html"
  MsgBox "完了: " & rows.Count & " 件" & vbCrLf & OUT_DIR & "\Report.html"

End Sub

' ============================================================
' Product配下パートを順次処理
Sub ScanProduct3D(prodDoc, rows)
  Dim prod : Set prod = prodDoc.Product
  If prod Is Nothing Then Exit Sub
  Dim i
  For i = 1 To prod.Products.Count
    Dim child : Set child = prod.Products.Item(i)
    If Not child Is Nothing Then
      Dim refDoc : Set refDoc = child.ReferenceProduct.Parent
      If Not refDoc Is Nothing Then
        If TypeName(refDoc) = "PartDocument" Then
          ScanPart3D refDoc, rows
        End If
      End If
    End If
  Next
End Sub

' ============================================================
' 1 Part を全フェース走査 → エッジ分類 → U/I検出 → 記録
Sub ScanPart3D(partDoc, rows)

  Dim part : Set part = partDoc.Part
  Dim sel  : Set sel  = partDoc.Selection
  Dim spa  : Set spa  = partDoc.GetWorkbench("SPAWorkbench")

  sel.Clear
  sel.Search "Topology.Face,all"   ' 全フェース収集（トップロジ）  :contentReference[oaicite:5]{index=5}
  If sel.Count = 0 Then Exit Sub

  Dim f
  For f = 1 To sel.Count

    ' ---- フェース基底情報（原点・面内2軸） ----
    Dim faceRef : Set faceRef = sel.Item(f).Reference
    Dim measF   : Set measF   = spa.GetMeasurable(faceRef)
    Dim pl()    : ReDim pl(8)
    On Error Resume Next
    measF.GetPlane pl  ' origin(0..2), dir1(3..5), dir2(6..8)  :contentReference[oaicite:6]{index=6}
    If Err.Number <> 0 Then Err.Clear: GoTo nextFace
    On Error GoTo 0

    Dim fOrigin(2), uDir(2), vDir(2)
    fOrigin(0)=pl(0): fOrigin(1)=pl(1): fOrigin(2)=pl(2)
    uDir(0)=pl(3): uDir(1)=pl(4): uDir(2)=pl(5)
    vDir(0)=pl(6): vDir(1)=pl(7): vDir(2)=pl(8)
    NormVec uDir: NormVec vDir

    ' ---- このフェースのエッジを列挙 ----
    Dim edges : Set edges = CreateObject("System.Collections.ArrayList")
    Dim eSel  : Set eSel  = partDoc.Selection
    eSel.Clear
    eSel.Add faceRef
    eSel.Search "Topology.Edge,sel" ' face の境界エッジ群  :contentReference[oaicite:7]{index=7}

    Dim e
    For e = 1 To eSel.Count
      Dim eref : Set eref = eSel.Item(e).Reference
      Dim em   : Set em   = spa.GetMeasurable(eref)

      Dim okEdge : okEdge = False
      Dim ed
      Set ed = CreateObject("Scripting.Dictionary")
      ed("ref") = eref

      ' 端点・中点（3D）
      Dim p() : ReDim p(8)
      On Error Resume Next
      em.GetPointsOnCurve p ' start(0..2), mid(3..5), end(6..8)  :contentReference[oaicite:8]{index=8}
      If Err.Number = 0 Then
        ed("p0") = Arr3(p(0),p(1),p(2))
        ed("pm") = Arr3(p(3),p(4),p(5))
        ed("p1") = Arr3(p(6),p(7),p(8))
        okEdge = True
      End If
      Err.Clear
      On Error GoTo 0
      If Not okEdge Then GoTo nextEdge

      ' 長さ
      Dim elen : elen = 0#
      On Error Resume Next
      elen = em.Length   ' ㎜  :contentReference[oaicite:9]{index=9}
      On Error GoTo 0
      ed("len") = CDbl(elen)

      ' 半径（円弧なら成功）
      Dim rad : rad = -1#
      On Error Resume Next
      rad = em.Radius    ' 円/球/円筒 or 円弧  :contentReference[oaicite:10]{index=10}
      On Error GoTo 0

      If rad > 0 Then
        ' 円弧エッジと仮定：弧角 ≒ 長さ/半径
        Dim angDeg : angDeg = (ed("len")/rad) * 180# / 3.141592653589793#
        ed("type") = "ARC"
        ed("r")    = rad
        ed("ang")  = angDeg
        ' 中心（取得できる場合のみ）
        On Error Resume Next
        Dim cc() : ReDim cc(2)
        em.GetCenter cc  ' CIRCLE/Sphere系  :contentReference[oaicite:11]{index=11}
        If Err.Number = 0 Then ed("c") = Arr3(cc(0),cc(1),cc(2))
        Err.Clear: On Error GoTo 0
      Else
        ' 直線性：start-mid-end が十分一直線なら LINE とみなす
        Dim isLine : isLine = Collinear(ed("p0"), ed("pm"), ed("p1"), 0.01, 0.5) ' [mm], [deg]
        If isLine Then
          ed("type") = "LINE"
        Else
          ed("type") = "CURVE" ' スプライン等、今回は不採用
        End If
      End If

      edges.Add ed
nextEdge:
    Next

    If edges.Count = 0 Then GoTo nextFace

    ' ---- U候補（半径/弧角チェック）抽出 ----
    Dim arcs : Set arcs = CreateObject("System.Collections.ArrayList")
    Dim lines: Set lines= CreateObject("System.Collections.ArrayList")
    Dim i
    For i = 0 To edges.Count-1
      Dim ei : Set ei = edges(i)
      If ei.Exists("type") Then
        If ei("type") = "ARC" Then
          If ei.Exists("r") And ei.Exists("ang") Then
            If ei("r")>=R_MIN And ei("r")<=R_MAX And Abs(ei("ang")-180#)<=TOL_DEG_ARC Then
              arcs.Add ei
            End If
          End If
        ElseIf ei("type") = "LINE" Then
          lines.Add ei
        End If
      End If
    Next
    If arcs.Count = 0 Then GoTo nextFace

    ' ---- 端点インデックス（端点一致を高速判定） ----
    Dim endMap : Set endMap = CreateObject("Scripting.Dictionary")
    For i = 0 To edges.Count-1
      Dim e0 : Set e0 = edges(i)
      If Not e0.Exists("p0") Then Continue For
      AddEndKey endMap, e0, "p0"
      AddEndKey endMap, e0, "p1"
    Next

    ' ---- U検出（半円弧 + その両端と接続する直線×2） ----
    Dim uList : Set uList = CreateObject("System.Collections.ArrayList")
    For i = 0 To arcs.Count-1
      Dim ua : Set ua = arcs(i)
      Dim u_p0 : u_p0 = KeyOfPoint(ua("p0"))
      Dim u_p1 : u_p1 = KeyOfPoint(ua("p1"))
      Dim candL1 : Set candL1 = LinesAtKey(lines, endMap, u_p0)
      Dim candL2 : Set candL2 = LinesAtKey(lines, endMap, u_p1)
      If candL1.Count=0 Or candL2.Count=0 Then Continue For

      ' 最初の組を採用（必要なら最短/最長などで最良組選定に拡張）
      Dim UL1 : Set UL1 = candL1(0)
      Dim UL2 : Set UL2 = candL2(0)

      ' U開口：UL1自由端とUL2自由端の距離（フェース2Dで測る）
      Dim c2() : c2 = EnsureCenter(ua, fOrigin, uDir, vDir)
      Dim l1_free() : l1_free = FreeEnd2D(UL1, ua("p0"), uDir, vDir, fOrigin)
      Dim l2_free() : l2_free = FreeEnd2D(UL2, ua("p1"), uDir, vDir, fOrigin)
      Dim opening : opening = Dist2D(l1_free(0),l1_free(1), l2_free(0),l2_free(1))

      ' Uパターン 1件
      Dim urec : Set urec = CreateObject("Scripting.Dictionary")
      urec("arc") = ua
      urec("L1")  = UL1
      urec("L2")  = UL2
      urec("opening") = opening
      urec("center")  = EnsureCenter(ua) ' 3D中心（無ければ近似後で補う）
      uList.Add urec
    Next

    If uList.Count = 0 Then GoTo nextFace

    ' ---- U/I間隔（3D中心）で間引き ----
    Dim acceptedU : Set acceptedU = CreateObject("System.Collections.ArrayList")
    Dim anchors : Set anchors = CreateObject("System.Collections.ArrayList") ' 3D点配列
    For i = 0 To uList.Count-1
      Dim ux : Set ux = uList(i)
      Dim uc() : uc = ux("center")
      If IsFarFromAnchors(anchors, uc, MIN_SPACING3D) Then
        acceptedU.Add ux
        anchors.Add uc
      End If
    Next
    If acceptedU.Count = 0 Then GoTo nextFace

    ' ---- I検出（U内のコの字：中央線Lbの両端に直角で接続） ----
    Dim j
    For j = 0 To acceptedU.Count-1
      Dim U : Set U = acceptedU(j)
      Dim UA : Set UA = U("arc")
      Dim UL1: Set UL1 = U("L1")
      Dim UL2: Set UL2 = U("L2")

      ' 半円内部の半平面条件を定義（フェース2D基準）
      Dim cx2,cy2, bUx,bUy
      Dim sc() : sc = SemiPlaneForArc(UA, fOrigin, uDir, vDir, cx2, cy2, bUx, bUy)

      ' I候補探索：フェース内の直線から中央線候補Lbを選ぶ
      Dim k
      Dim foundI : foundI = False
      For k = 0 To lines.Count-1
        Dim Lb : Set Lb = lines(k)
        ' Lbの両端2D
        Dim Lb0() : Lb0 = Proj2D(Lb("p0"), fOrigin, uDir, vDir)
        Dim Lb1() : Lb1 = Proj2D(Lb("p1"), fOrigin, uDir, vDir)
        ' 端点がU半円内か（円内 + ビセクタ半平面）
        If Not (InSemiCircle(Lb0(0),Lb0(1), cx2,cy2, bUx,bUy, UA("r"))) Then Continue For
        If Not (InSemiCircle(Lb1(0),Lb1(1), cx2,cy2, bUx,bUy, UA("r"))) Then Continue For

        ' Lbの両端それぞれに「直角接続する別線」を探す
        Dim sNg : Set sNg = RightNeighbors(lines, Lb, True,  fOrigin, uDir, vDir, TOL_JOIN, TOL_RIGHT_DEG)
        Dim eNg : Set eNg = RightNeighbors(lines, Lb, False, fOrigin, uDir, vDir, TOL_JOIN, TOL_RIGHT_DEG)
        If sNg.Count=0 Or eNg.Count=0 Then Continue For

        ' 組合せから1組採用
        Dim La : Set La = sNg(0)
        Dim Lc : Set Lc = eNg(0)

        ' 自由端2点もU半円内か確認
        Dim aFree() : aFree = FreeEnd2DAgainst(L a, Lb, fOrigin, uDir, vDir, TOL_JOIN)
        Dim cFree() : cFree = FreeEnd2DAgainst(L c, Lb, fOrigin, uDir, vDir, TOL_JOIN)
        If Not (InSemiCircle(aFree(0),aFree(1), cx2,cy2, bUx,bUy, UA("r"))) Then Continue For
        If Not (InSemiCircle(cFree(0),cFree(1), cx2,cy2, bUx,bUy, UA("r"))) Then Continue For

        ' U↔I 距離（最小）を計測（3D最短距離）
        Dim dMin : dMin = MinUtoILines(part, UL1, UL2, La, Lb, Lc)

        ' ---- スナップ（パターン単位） ----
        Dim snapRel, snapAbs
        snapRel = "snaps\" & Sanitize(UId(partDoc, part, f, UA)) & "__" & Sanitize(IId(partDoc, part, f, Lb, La, Lc)) & ".png"
        snapAbs = OUT_DIR & "\" & snapRel
        If ShouldSnap() Then
          Capture(partDoc, Array(UA("ref"), UL1("ref"), UL2("ref"), La("ref"), Lb("ref"), Lc("ref")), snapAbs)
        End If

        ' ---- 行追加 ----
        rows.Add Array( _
          partDoc.Name, part.Name, "Face#" & CStr(f), "U+I", _
          Round(UA("r"),3), Round(UA("ang"),2), Round(UL1("len"),3), Round(UL2("len"),3), Round(U("opening"),3), _
          Round(La("len"),3), Round(Lb("len"),3), Round(Lc("len"),3), "OK", _
          Round(dMin,3), snapRel _
        )

        foundI = True
        Exit For
      Next

      ' IがなければU単独で出力（参考用）
      If Not foundI Then
        Dim snapRelU, snapAbsU
        snapRelU = "snaps\" & Sanitize(UId(partDoc, part, f, UA)) & ".png"
        snapAbsU = OUT_DIR & "\" & snapRelU
        If ShouldSnap() Then
          Capture(partDoc, Array(UA("ref"), UL1("ref"), UL2("ref")), snapAbsU)
        End If

        rows.Add Array( _
          partDoc.Name, part.Name, "Face#" & CStr(f), "U", _
          Round(UA("r"),3), Round(UA("ang"),2), Round(UL1("len"),3), Round(UL2("len"),3), Round(U("opening"),3), _
          "", "", "", "", "", snapRelU _
        )
      End If

    Next ' acceptedU

nextFace:
  Next ' face

End Sub

' ===================== 幾何ヘルパ =========================
Function Arr3(x,y,z) : Arr3 = Array(CDbl(x),CDbl(y),CDbl(z)) : End Function

Sub NormVec(v)
  Dim n : n = Sqr(v(0)*v(0)+v(1)*v(1)+v(2)*v(2))
  If n>0 Then v(0)=v(0)/n: v(1)=v(1)/n: v(2)=v(2)/n
End Sub

Function Dot(a,b) : Dot = a(0)*b(0)+a(1)*b(1)+a(2)*b(2) : End Function

Function Sub3(a,b) : Sub3 = Array(a(0)-b(0), a(1)-b(1), a(2)-b(2)) : End Function

Function Dist3(a,b)
  Dist3 = Sqr((a(0)-b(0))^2 + (a(1)-b(1))^2 + (a(2)-b(2))^2)
End Function

Function Proj2D(p, o, u, v)
  Dim w : w = Sub3(p,o)
  Proj2D = Array(Dot(w,u), Dot(w,v))
End Function

Function Dist2D(x1,y1,x2,y2)
  Dist2D = Sqr((x2-x1)^2 + (y2-y1)^2)
End Function

Function Angle2D(ax,ay,bx,by)
  Dim da : da = Sqr(ax*ax+ay*ay) : If da=0 Then Angle2D=0: Exit Function
  Dim db : db = Sqr(bx*bx+by*by) : If db=0 Then Angle2D=0: Exit Function
  Dim d : d = (ax*bx+ay*by)/(da*db)
  If d>1 Then d=1
  If d<-1 Then d=-1
  Angle2D = Atn(Sqr(1-d*d)/d)*180#/3.141592653589793#
  If Angle2D<0 Then Angle2D=Angle2D+180#
End Function

Function Collinear(p0,pm,p1,tol_mm,tol_deg)
  ' 3点が一直線か：∠(p0->pm, pm->p1) ≈ 180°
  Dim a() : a = Sub3(pm,p0)
  Dim b() : b = Sub3(p1,pm)
  Dim ax : ax = a(0): Dim ay : ay = a(1): Dim az : az = a(2)
  Dim bx : bx = b(0): Dim by : by = b(1): Dim bz : bz = b(2)
  Dim da : da = Sqr(ax*ax+ay*ay+az*az)
  Dim db : db = Sqr(bx*bx+by*by+bz*bz)
  If da<tol_mm Or db<tol_mm Then Collinear=False: Exit Function
  Dim d : d = (ax*bx+ay*by+az*bz)/(da*db)
  If d>1 Then d=1
  If d<-1 Then d=-1
  Dim ang : ang = Atn(Sqr(1-d*d)/d)*180#/3.141592653589793#
  If ang<0 Then ang=ang+180#
  Collinear = (Abs(ang-180#) <= tol_deg)
End Function

Function KeyOfPoint(p)
  KeyOfPoint = CStr(Round(p(0)/TOL_JOIN,0)) & "|" & CStr(Round(p(1)/TOL_JOIN,0)) & "|" & CStr(Round(p(2)/TOL_JOIN,0))
End Function

Sub AddEndKey(endMap, ed, keyName)
  Dim k : k = KeyOfPoint(ed(keyName))
  If Not endMap.Exists(k) Then
    Dim lst : Set lst = CreateObject("System.Collections.ArrayList")
    endMap.Add k, lst
  End If
  endMap(k).Add ed
End Sub

Function LinesAtKey(lines, endMap, key)
  Dim lst : Set lst = CreateObject("System.Collections.ArrayList")
  If endMap.Exists(key) Then
    Dim i
    For i = 0 To endMap(key).Count-1
      Dim e : Set e = endMap(key)(i)
      If e("type")="LINE" Then lst.Add e
    Next
  End If
  Set LinesAtKey = lst
End Function

Function EnsureCenter(arc, Optional o, Optional u, Optional v)
  ' 3D中心が無い場合は端点中点から近似（円の垂直二等分）
  If arc.Exists("c") Then
    EnsureCenter = arc("c"): Exit Function
  End If
  ' 近似：フェース2Dに落として二等分線交点で推定
  Dim p0() : p0 = Proj2D(arc("p0"), o, u, v)
  Dim pm() : pm = Proj2D(arc("pm"), o, u, v)
  Dim p1() : p1 = Proj2D(arc("p1"), o, u, v)
  Dim cx2 : cx2 = (p0(0)+p1(0))/2
  Dim cy2 : cy2 = (p0(1)+p1(1))/2
  EnsureCenter = Arr3(cx2*u(0)+cy2*v(0)+o(0), cx2*u(1)+cy2*v(1)+o(1), cx2*u(2)+cy2*v(2)+o(2))
End Function

Function FreeEnd2D(L, u_end3, u, v, o)
  ' U端点に接しているLの自由端を2Dで返す
  Dim k0 : k0 = KeyOfPoint(L("p0"))
  Dim ku : ku = KeyOfPoint(u_end3)
  Dim pFree()
  If k0 = ku Then
    pFree = Proj2D(L("p1"), o, u, v)
  Else
    pFree = Proj2D(L("p0"), o, u, v)
  End If
  FreeEnd2D = pFree
End Function

Function FreeEnd2DAgainst(L, Lb, o, u, v, tolJoin)
  Dim kB0 : kB0 = KeyOfPoint(Lb("p0"))
  Dim kB1 : kB1 = KeyOfPoint(Lb("p1"))
  Dim kL0 : kL0 = KeyOfPoint(L("p0"))
  Dim kL1 : kL1 = KeyOfPoint(L("p1"))
  Dim tgt3
  If kL0=kB0 Or kL0=kB1 Then
    tgt3 = L("p1")
  Else
    tgt3 = L("p0")
  End If
  FreeEnd2DAgainst = Proj2D(tgt3, o, u, v)
End Function

Function SemiPlaneForArc(arc, o, u, v, ByRef cx2, ByRef cy2, ByRef bx, ByRef by)
  ' 半円の「内側半平面」を定義：端点方向ベクトルの二等分ベクトルで半平面を切る
  Dim p0() : p0 = Proj2D(arc("p0"), o, u, v)
  Dim p1() : p1 = Proj2D(arc("p1"), o, u, v)
  Dim c3() : c3 = EnsureCenter(arc, o, u, v)
  Dim c2() : c2 = Proj2D(c3, o, u, v)
  cx2 = c2(0): cy2 = c2(1)
  Dim v0x : v0x = p0(0)-cx2: Dim v0y : v0y = p0(1)-cy2
  Dim v1x : v1x = p1(0)-cx2: Dim v1y : v1y = p1(1)-cy2
  Dim bx0 : bx0 = v0x + v1x
  Dim by0 : by0 = v0y + v1y
  Dim nn : nn = Sqr(bx0*bx0+by0*by0)
  If nn=0 Then bx0=1: by0=0 Else bx0=bx0/nn: by0=by0/nn
  bx = bx0: by = by0
End Function

Function InSemiCircle(x,y, cx,cy, bx,by, r)
  Dim dx : dx = x-cx
  Dim dy : dy = y-cy
  If (dx*dx+dy*dy) > (r+0.001)*(r+0.001) Then
    InSemiCircle = False: Exit Function
  End If
  ' 半平面条件： (p-c)·b >= 0
  InSemiCircle = (dx*bx + dy*by >= -0.001)
End Function

Function RightNeighbors(lines, Lb, atStart, o, u, v, tolJoin, tolRight)
  Dim res : Set res = CreateObject("System.Collections.ArrayList")
  Dim bx,by, ux,uy
  Dim B0(), B1()
  B0 = Proj2D(Lb("p0"), o, u, v)
  B1 = Proj2D(Lb("p1"), o, u, v)
  If atStart Then
    bx = B1(0)-B0(0): by = B1(1)-B0(1)
  Else
    bx = B0(0)-B1(0): by = B0(1)-B1(1)
  End If

  Dim k
  For k = 0 To lines.Count-1
    Dim L : Set L = lines(k)
    If (L Is Lb) Then Continue For

    Dim share As Boolean
    share = (KeyOfPoint(L("p0"))=KeyOfPoint(IIf(atStart,Lb("p0"),Lb("p1"))) Or _
             KeyOfPoint(L("p1"))=KeyOfPoint(IIf(atStart,Lb("p0"),Lb("p1"))))
    If Not share Then Continue For

    Dim P0(), P1()
    P0 = Proj2D(L("p0"), o, u, v)
    P1 = Proj2D(L("p1"), o, u, v)
    If atStart Then
      ux = P1(0)-P0(0): uy = P1(1)-P0(1)
    Else
      ux = P0(0)-P1(0): uy = P0(1)-P1(1)
    End If

    Dim ang : ang = Angle2D(bx,by, ux,uy)
    If Abs(ang-90#) <= tolRight Then res.Add L
  Next
  Set RightNeighbors = res
End Function

Function MinUtoILines(part, UL1, UL2, La, Lb, Lc)
  ' 3D最短距離（Measurable.GetMinimumDistance）でU直線2本とI直線3本の最小を返す
  Dim spa : Set spa = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
  Dim dMin : dMin = 1E+9
  dMin = Min3( _
    Min3( MinDist(spa, UL1("ref"), La("ref")), MinDist(spa, UL1("ref"), Lb("ref")), MinDist(spa, UL1("ref"), Lc("ref")) ), _
    Min3( MinDist(spa, UL2("ref"), La("ref")), MinDist(spa, UL2("ref"), Lb("ref")), MinDist(spa, UL2("ref"), Lc("ref")) ), _
    dMin _
  )
  MinUtoILines = dMin
End Function

Function MinDist(spa, r1, r2)
  On Error Resume Next
  Dim m1 : Set m1 = spa.GetMeasurable(r1)
  MinDist = m1.GetMinimumDistance(r2)  ' 3D最短距離  :contentReference[oaicite:12]{index=12}
  If Err.Number<>0 Then MinDist=1E+9: Err.Clear
  On Error GoTo 0
End Function

Function Min3(a,b,c)
  Dim m : m=a : If b<m Then m=b : If c<m Then m=c
  Min3 = m
End Function

Function IsFarFromAnchors(anchors, p, minD)
  Dim i
  For i = 0 To anchors.Count-1
    If Dist3(anchors(i), p) < minD Then IsFarFromAnchors=False: Exit Function
  Next
  IsFarFromAnchors = True
End Function

' ===================== スナップ & 出力 =========================
Function ShouldSnap()
  gPatCount = gPatCount + 1
  If Not ENABLE_SNAPS Then ShouldSnap=False: Exit Function
  If gSnapCount >= SNAP_MAX Then ShouldSnap=False: Exit Function
  If (gPatCount Mod SNAP_EVERY_N) <> 0 Then ShouldSnap=False: Exit Function
  ShouldSnap=True
End Function

Sub Capture(partDoc, refs, pngPath)
  On Error Resume Next
  Dim sel : Set sel = partDoc.Selection
  sel.Clear
  Dim i
  For i = LBound(refs) To UBound(refs)
    If Not (refs(i) Is Nothing) Then sel.Add refs(i)
  Next
  Dim v : Set v = CATIA.ActiveWindow.ActiveViewer
  v.ReframeOnSelection
  If Err.Number <> 0 Then Err.Clear: v.Reframe
  CATIA.ActiveWindow.CapturePictureFile pngPath, "png"
  sel.Clear
  gSnapCount = gSnapCount + 1
  On Error GoTo 0
End Sub

Function Sanitize(s)
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
  Sanitize = t
End Function

Function UId(doc, part, fIdx, arc)
  Dim r : r = "U_F" & CStr(fIdx) & "_R" & CStr(Round(arc("r"),3))
  UId = r
End Function

Function IId(doc, part, fIdx, Lb, La, Lc)
  IId = "I_F" & CStr(fIdx)
End Function

Sub EnsureFolder(path)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Sub BuildHtmlReport(rows, htmlPath)
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  Dim ts  : Set ts  = fso.CreateTextFile(htmlPath, True, True)
  ts.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'><title>U/I 3D Report</title>"
  ts.WriteLine "<style>body{font-family:Segoe UI,Arial,sans-serif;margin:16px} table{border-collapse:collapse;width:100%} th,td{border:1px solid #ccc;padding:6px 8px;font-size:12px} th{background:#f5f5f5;position:sticky;top:0} .row{display:grid;grid-template-columns:1fr 320px;gap:16px;margin:12px 0;padding:8px;border:1px solid #e5e5e5;border-radius:8px} .snap{border:1px solid #ddd;border-radius:6px;padding:6px;text-align:center}</style>"
  ts.WriteLine "</head><body><h2>U/I 3D Report</h2>"
  ts.WriteLine "<p>Total: " & CStr(rows.Count) & "</p>"
  Dim i
  For i = 0 To rows.Count-1
    Dim r : r = rows(i)
    ts.WriteLine "<div class='row'>"
    ts.WriteLine "<table><tbody>"
    ts.WriteLine TR2("Doc", r(0)) & TR2("Part", r(1)) & TR2("Face", r(2)) & TR2("Pattern", r(3))
    ts.WriteLine TR2("U_R[mm]", r(4)) & TR2("U_Ang[deg]", r(5)) & TR2("U_L1[mm]", r(6)) & TR2("U_L2[mm]", r(7)) & TR2("U_Opening[mm]", r(8))
    ts.WriteLine TR2("I_La[mm]", r(9)) & TR2("I_Lb[mm]", r(10)) & TR2("I_Lc[mm]", r(11)) & TR2("Right90°", r(12)) & TR2("U↔I MinDist[mm]", r(13))
    ts.WriteLine "</tbody></table>"
    Dim img : img = r(14)
    If Len(img)>0 Then
      ts.WriteLine "<div class='snap'><img src='" & Replace(img,"\","/") & "' style='max-width:300px;height:" & SNAP_HEIGHT_PX & "px;object-fit:contain'><br><small>" & img & "</small></div>"
    Else
      ts.WriteLine "<div class='snap'><em>no snapshot</em></div>"
    End If
    ts.WriteLine "</div>"
  Next
  ts.WriteLine "</body></html>"
  ts.Close
End Sub

Function TR2(k,v)
  TR2 = "<tr><th>" & k & "</th><td>" & v & "</td></tr>"
End Function
