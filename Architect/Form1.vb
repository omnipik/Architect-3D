Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Math
Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class Form1
#Region "object structures"
    <Serializable()> Public Structure XYZ
        Public Property X As Double
        Public Property Y As Double
        Public Property Z As Double
        Public Property W As Double
        Public Property G As PointF
    End Structure

    <Serializable()> Public Structure MDATA
        Public Property Mesh As XYZ()
        Public Property Mesh2D As PointF()
        Public Property Origin As XYZ
        Public Property Center As XYZ
        Public Property Distance As Double
        Public Property LineColor As Color
        Public Property FillColor As Color
        Public Property Line As Boolean
        Public Property Fill As Boolean
        Public Property LabelIO As Boolean
        Public Property Label As String
        Public Property Font As Font
        Public Property Curve As Boolean
        Public Property Radian As Double
        Public Property SpinIO As Boolean
        Public Property SpinAngleX As Double
        Public Property SpinAngleY As Double
        Public Property SpinAngleZ As Double
        Public Property OrbitIO As Boolean
        Public Property OrbitAngleX As Double
        Public Property OrbitAngleY As Double
        Public Property OrbitAngleZ As Double
        Public Property TranslateIO As Boolean
        Public Property TranslateX As Double
        Public Property TranslateY As Double
        Public Property TranslateZ As Double
    End Structure

    <Serializable()> Public Structure MODEL
        Public Property M As MDATA()
        Public Property Camera As XYZ
    End Structure
    <Serializable()> Public Structure ZONE
        Public Property Z As MDATA()()
        Public Property Camera As XYZ
    End Structure

    <Serializable()> Public Structure MTX
        Public Property c1r1 As Double
        Public Property c1r2 As Double
        Public Property c1r3 As Double
        Public Property c1r4 As Double
        Public Property c2r1 As Double
        Public Property c2r2 As Double
        Public Property c2r3 As Double
        Public Property c2r4 As Double
        Public Property c3r1 As Double
        Public Property c3r2 As Double
        Public Property c3r3 As Double
        Public Property c3r4 As Double
        Public Property c4r1 As Double
        Public Property c4r2 As Double
        Public Property c4r3 As Double
        Public Property c4r4 As Double
    End Structure
#End Region
#Region "Declarations"
    Dim fileMicroData As String = "c:\architect\data\micro"
    Dim fileMacroData As String = "c:\architect\data\macro"
    Dim fileImageData As String = "c:\architect\data\Image"
    Dim gImage, mImage As Bitmap
    Dim gr, gm As Graphics
    Dim M_DATA() As String
    Dim modC As Integer = 0
    Dim bViewUp As Boolean = False
    Dim bViewDown As Boolean = False
    Dim bViewLeft As Boolean = False
    Dim bViewRight As Boolean = False
    Dim bZoomIn As Boolean = False
    Dim bZoomOut As Boolean = False
    Dim bShiftR As Boolean = False
    Dim bShiftL As Boolean = False
    Dim bShiftD As Boolean = False
    Dim bShiftU As Boolean = False
    Dim spTop, spBottom, spLeft, spRight, spNear, spFar As Double
    Dim VertexCount As Integer = 0
    Dim proTime As Long = 0
    Dim valueAS, valueAT, valueMacArray, valueMicArray As Integer
    Dim valueVPA, valueVPR, valueVPG, valueVPB, valueVGS, valueVGD As Integer
    Dim valueVPFA, valueVPFR, valueVPFG, valueVPFB As Integer
    Dim valueVOX, valueVOY, valueVOZ, valueVORX, valueVORY, valueVORZ, valueVTX, valueVTY, valueVTZ, valueVSX, valueVSY, valueVSZ, valueVMX, valueVMY, valueVMZ, valueVAX, valueVAY, valueVAZ As Double
    Dim valueVRD, valueRule, ValueMark, valueVPX, valueVPY, valueVPZ As Double
    Dim valueMicro, valueMacro, valueVCZ As Double
    Dim Horidist As Double
    Dim cB1, cB2, cB3, cB4, cR1, cR2, cR3, cR4, cG1, cG2, cG3, cG4, bcC, fcC As Color
    Dim camLX, camRX, camTY, camBY, camDist As Double
    Dim gPoint, wPoint, wEye, mEye, wFar, wOrigin, wHorizon As XYZ
    Dim rad As Double = (2 * PI / 360) * 4
    Dim ScrX, ScrY, ScrMX, ScrMY, AR As Double
    Dim _Rad15 As Double = (PI / 180) * 15
    Dim _Rad30 As Double = (PI / 180) * 30
    Dim _Rad45 As Double = (PI / 180) * 45
    Dim _Rad90 As Double = (PI / 180) * 90
    Dim _Rad180 As Double = (PI / 180) * 180
    Dim GridDataXZ()(), GridDataXY()(), GridDataYZ()() As XYZ
    Dim mView As MTX = MI()
    Dim modView As MTX = MI()
    Dim MicroData() As MDATA = Nothing
    Dim ModelData() As MDATA = Nothing
    Dim MacroData()() As MDATA = Nothing
    Dim PaintQ() As MDATA = Nothing
    Dim uCount As Integer = 0
    Dim clock As Stopwatch = New Stopwatch
#End Region
#Region "Data Intialization"
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If IO.Directory.Exists(fileMicroData) = False Then IO.Directory.CreateDirectory(fileMicroData)
        If IO.Directory.Exists(fileMacroData) = False Then IO.Directory.CreateDirectory(fileMacroData)
        If IO.Directory.Exists(fileImageData) = False Then IO.Directory.CreateDirectory(fileImageData)
        Int_Data()
        Int_Display()
        Int_Model()
    End Sub
    Private Sub Int_Data()
        cR1 = cColor(200, 255, 0, 0) : cR2 = cColor(200, 150, 0, 0)
        cR3 = cColor(200, 50, 0, 0) : cR4 = cColor(200, 25, 0, 0)
        cG1 = cColor(200, 0, 255, 0) : cG2 = cColor(200, 0, 150, 0)
        cG3 = cColor(200, 0, 50, 0) : cG4 = cColor(200, 0, 25, 0)
        cB1 = cColor(255, 0, 0, 255) : cB2 = cColor(255, 0, 0, 150)
        cB3 = cColor(255, 0, 0, 100) : cB4 = cColor(255, 0, 0, 25)
        bcC = cColor(255, 0, 0, 25) : fcC = cColor(255, 0, 0, 255)
        '--------------------------------------------------------'
        gImage = New Bitmap(MainDisplay.Width, MainDisplay.Height)
        gr = Graphics.FromImage(gImage)
        mImage = New Bitmap(ModelDisplay.Width, ModelDisplay.Height)
        gm = Graphics.FromImage(mImage)
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        camDist = -5000
        ScrX = MainDisplay.Width
        ScrY = MainDisplay.Height
        ScrMX = ModelDisplay.Width
        ScrMY = ModelDisplay.Height + 100
        wEye.X = 0 : wEye.Y = 0 : wEye.Z = camDist : wEye.W = 1
        mEye.X = 0 : mEye.Y = 0 : mEye.Z = camDist : mEye.W = 1
        wOrigin.X = 0 : wOrigin.Y = 0 : wOrigin.Z = 0 : wOrigin.W = 1
        '--------------------------------------------------------'
        mView = MI()
        modView = MI()
        mView = mXm(ROX(_Rad15), mView)
        wEye = norm(mXv((ROX(-_Rad15)), wEye))
        mView = mXm(mView, ROY(-_Rad45))
        wEye = norm(mXv(ROY(_Rad45), wEye))
        modView = mXm(ROX(_Rad15), modView)
        mEye = norm(mXv((ROX(-_Rad15)), mEye))
        '--------------------------------------------------------'
        valueAS = 0 : valueAT = 0 : valueVPA = 255 : valueVPR = 255 : valueVPG = 255 : valueVPB = 255
        valueVAX = 0 : valueVAY = 0 : valueVAZ = 0 : valueVOX = 0 : valueVOY = 0 : valueVOZ = 0
        valueVORX = 0 : valueVORY = 0 : valueVORZ = 0 : valueVMX = -10 : valueVMY = 0 : valueVMZ = 0
        valueVSX = 0 : valueVSY = 0 : valueVSZ = 0 : valueVRD = 10
        valueRule = 0 : ValueMark = 10 : valueMicro = 0.01 : valueMacro = 10
        valueVGS = 200 : valueVGD = 20 : valueVPX = 10 : valueVPY = 0 : valueVPZ = 0
        valueVPFA = 255 : valueVPFR = 150 : valueVPFG = 150 : valueVPFB = 150
        valueMacArray = 0 : valueMicArray = 0
        '--------------------------------------------------------'
        cbFillplane.Checked = True
        cbSquare.Checked = True
        valueVCZ = 0
        cbGridXZ.Checked = True
        cbAxisLines.Checked = True
        '--------------------------------------------------------'
        GridDataXZ = drawGridXZ(200, 20) : GridDataXY = drawGridXY(200, 20) : GridDataYZ = drawGridYZ(200, 20)
        '--------------------------------------------------------'
    End Sub
    Private Sub Int_Display()
        tbAdjustVX.Text = valueVAX.ToString("N3")
        tbAdjustVY.Text = valueVAY.ToString("N3")
        tbAdjustVZ.Text = valueVAZ.ToString("N3")
        tbAdjustSelect.Text = valueAS.ToString
        tbAdjustTotal.Text = valueAT.ToString
        tbAdjustMacro.Text = valueMacArray.ToString
        tbAdjustMicro.Text = valueMicArray.ToString
        tbOriginVX.Text = valueVOX.ToString("N3")
        tbOriginVY.Text = valueVOY.ToString("N3")
        tbOriginVZ.Text = valueVOZ.ToString("N3")
        tbMeasureVX.Text = valueVMX.ToString("N3")
        tbMeasureVY.Text = valueVMY.ToString("N3")
        tbMeasureVZ.Text = valueVMZ.ToString("N3")
        tbMicroV.Text = valueRule.ToString("N3")
        tbMacroV.Text = ValueMark.ToString("N3")
        tbSpinVX.Text = valueVSX.ToString("N3")
        tbSpinVY.Text = valueVSY.ToString("N3")
        tbSpinVZ.Text = valueVSZ.ToString("N3")
        tbTranslateVX.Text = valueVTX.ToString("N3")
        tbTranslateVY.Text = valueVTY.ToString("N3")
        tbTranslateVZ.Text = valueVTZ.ToString("N3")
        tbRadiusV.Text = valueVRD.ToString("N3")
        tbPaintVA.Text = valueVPA.ToString
        tbPaintVR.Text = valueVPR.ToString
        tbPaintVG.Text = valueVPG.ToString
        tbPaintVB.Text = valueVPB.ToString
        tbPaintVFA.Text = valueVPFA.ToString
        tbPaintVFR.Text = valueVPFR.ToString
        tbPaintVFG.Text = valueVPFG.ToString
        tbPaintVFB.Text = valueVPFB.ToString
        tbMicroV.Text = valueMicro.ToString("N3")
        tbMacroV.Text = valueMacro.ToString("N3")
        tbPeakVX.Text = valueVPX.ToString("N3")
        tbPeakVY.Text = valueVPY.ToString("N3")
        tbPeakVZ.Text = valueVPZ.ToString("N3")
        tbGridSize.Text = valueVGS.ToString
        tbGridDim.Text = valueVGD.ToString
        '--------------------------------------------------------'
    End Sub
    Private Sub Int_Model()
        If IO.Directory.GetFiles(fileMicroData).Length <> 0 Then
            M_DATA = IO.Directory.GetFiles(fileMicroData)
            ModelData = M_Switch(M_DATA(modC))
        End If
    End Sub
#End Region
#Region "Plot enviroment"
    Private Function rNum(ByVal min As Integer, ByVal max As Integer) As Integer
        Static r As Random = New Random
        Return r.Next(min, max)
    End Function
    Private Function rrNum(ByVal min As Integer, ByVal max As Integer) As Integer
        Static r As Random = New Random(rNum(0, 32767))
        Return r.Next(min, max)
    End Function
    Private Function drawXA(ByVal l As Integer) As XYZ()
        Dim tlV(1) As XYZ
        tlV(0).X = -l : tlV(0).Y = 0 : tlV(0).Z = 0 : tlV(0).W = 1
        tlV(1).X = l : tlV(1).Y = 0 : tlV(1).Z = 0 : tlV(1).W = 1
        Return tlV
    End Function
    Private Function drawYA(ByVal l As Integer) As XYZ()
        Dim tlV(1) As XYZ
        tlV(0).X = 0 : tlV(0).Y = -l : tlV(0).Z = 0 : tlV(0).W = 1
        tlV(1).X = 0 : tlV(1).Y = l : tlV(1).Z = 0 : tlV(1).W = 1
        Return tlV
    End Function
    Private Function drawZA(ByVal l As Integer) As XYZ()
        Dim tlV(1) As XYZ
        tlV(0).X = 0 : tlV(0).Y = 0 : tlV(0).Z = -l : tlV(0).W = 1
        tlV(1).X = 0 : tlV(1).Y = 0 : tlV(1).Z = l : tlV(1).W = 1
        Return tlV
    End Function
    Private Function drawAL(ByVal l As Integer) As XYZ()
        Dim tav(5) As XYZ
        tav(0).X = -l : tav(0).Y = 0 : tav(0).Z = 0 : tav(0).W = 1
        tav(1).X = l : tav(1).Y = 0 : tav(1).Z = 0 : tav(1).W = 1
        tav(2).X = 0 : tav(2).Y = -l : tav(2).Z = 0 : tav(2).W = 1
        tav(3).X = 0 : tav(3).Y = l : tav(3).Z = 0 : tav(3).W = 1
        tav(4).X = 0 : tav(4).Y = 0 : tav(4).Z = -l : tav(4).W = 1
        tav(5).X = 0 : tav(5).Y = 0 : tav(5).Z = l : tav(5).W = 1
        Return tav
    End Function

    Private Function tBuildXZ(ByVal s As Double, ByVal d As Double, ByVal L As Boolean, ByVal F As Boolean, _
                            ByVal gc1 As Color, ByVal gc2 As Color) As MDATA()
        Dim twa() As MDATA = Nothing
        Dim c As Integer = 0
        For i = -s To s - d Step d
            For x As Integer = -s To s - d Step d
                ReDim Preserve twa(c)
                twa(c).Mesh = xzTile(vCompile(i, 0, x), d)
                twa(c).LineColor = gc1
                twa(c).FillColor = gc2
                twa(c).Origin = vCompile(i, 0, x)
                twa(c).Center = vCompile(i, 0, x)
                If L = True Then twa(c).Line = True Else twa(c).Line = False
                If F = True Then twa(c).Fill = True Else twa(c).Fill = False
                twa(c).Curve = False
                c += 1
            Next x
        Next i
        Return twa
    End Function
    Private Function tBuildXY(ByVal s As Double, ByVal d As Double, ByVal L As Boolean, ByVal F As Boolean, _
                        ByVal gc1 As Color, ByVal gc2 As Color) As MDATA()
        Dim twa() As MDATA = Nothing
        Dim c As Integer = 0
        For i = -s To s - d Step d
            For x = -s To s - d Step d
                ReDim Preserve twa(c)
                twa(c).Mesh = xyTile(vCompile(i, x, 0), d)
                twa(c).LineColor = gc1
                twa(c).FillColor = gc2
                twa(c).Origin = vCompile(i, x, 0)
                twa(c).Center = vCompile(i, x, 0)
                If L = True Then twa(c).Line = True Else twa(c).Line = False
                If F = True Then twa(c).Fill = True Else twa(c).Fill = False
                twa(c).Curve = False
                c += 1
            Next x
        Next i
        Return twa
    End Function
    Private Function tBuildYZ(ByVal s As Double, ByVal d As Double, ByVal L As Boolean, ByVal F As Boolean, _
                        ByVal gc1 As Color, ByVal gc2 As Color) As MDATA()
        Dim twa() As MDATA = Nothing
        Dim c As Integer = 0
        For i = -s To s - d Step d
            For x = -s To s - d Step d
                ReDim Preserve twa(c)
                twa(c).Mesh = yzTile(vCompile(0, i, x), d)
                twa(c).LineColor = gc1
                twa(c).FillColor = gc2
                twa(c).Origin = vCompile(0, i, x)
                twa(c).Center = vCompile(0, i, x)
                If L = True Then twa(c).Line = True Else twa(c).Line = False
                If F = True Then twa(c).Fill = True Else twa(c).Fill = False
                twa(c).Curve = False
                c += 1
            Next x
        Next i
        Return twa
    End Function
    Private Function drawGridXZ(ByVal s As Double, ByVal d As Double) As XYZ()()
        Dim twa()() As XYZ = Nothing
        Dim c As Integer = 0
        For x As Integer = -s To s Step d
            ReDim Preserve twa(c)
            twa(c) = lineto(vCompile(-s, 0, x), vCompile(s, 0, x))
            c += 1
        Next x
        For x As Integer = -s To s Step d
            ReDim Preserve twa(c)
            twa(c) = lineto(vCompile(x, 0, -s), vCompile(x, 0, s))
            c += 1
        Next x
        Return twa
    End Function
    Private Function drawGridXY(ByVal s As Double, ByVal d As Double) As XYZ()()
        Dim twa()() As XYZ = Nothing
        Dim c As Integer = 0
        For x As Integer = -s To s Step d
            ReDim Preserve twa(c)
            twa(c) = lineto(vCompile(-s, x, 0), vCompile(s, x, 0))
            c += 1
        Next x
        For x As Integer = -s To s Step d
            ReDim Preserve twa(c)
            twa(c) = lineto(vCompile(x, -s, 0), vCompile(x, s, 0))
            c += 1
        Next x
        Return twa
    End Function
    Private Function drawGridYZ(ByVal s As Double, ByVal d As Double) As XYZ()()
        Dim twa()() As XYZ = Nothing
        Dim c As Integer = 0
        For x As Integer = -s To s Step d
            ReDim Preserve twa(c)
            twa(c) = lineto(vCompile(0, -s, x), vCompile(0, s, x))
            c += 1
        Next x
        For x As Integer = -s To s Step d
            ReDim Preserve twa(c)
            twa(c) = lineto(vCompile(0, x, -s), vCompile(0, x, s))
            c += 1
        Next x
        Return twa
    End Function
#End Region
#Region "Designer"
    Private Function vCompile(ByVal x As Double, ByVal y As Double, ByVal z As Double) As XYZ
        Dim tV As XYZ
        tV.X = x : tV.Y = y : tV.Z = z : tV.W = 1
        Return tV
    End Function
    Private Function lineto(ByVal vec1 As XYZ, ByVal vec2 As XYZ) As XYZ()
        Dim V(1) As XYZ
        V(0).X = vec1.X : V(0).Y = vec1.Y : V(0).Z = vec1.Z : V(0).W = 1
        V(1).X = vec2.X : V(1).Y = vec2.Y : V(1).Z = vec2.Z : V(1).W = 1
        Return V
    End Function
    Private Function xzTile(ByVal o As XYZ, ByVal d As Double) As XYZ()
        Dim V(4) As XYZ
        V(0) = o
        V(1).X = o.X + d : V(1).Y = o.Y : V(1).Z = o.Z : V(1).W = 1
        V(2).X = o.X + d : V(2).Y = o.Y : V(2).Z = o.Z + d : V(2).W = 1
        V(3).X = o.X : V(3).Y = o.Y : V(3).Z = o.Z + d : V(3).W = 1
        V(4) = o
        Return V
    End Function
    Private Function xyTile(ByVal o As XYZ, ByVal d As Double) As XYZ()
        Dim V(4) As XYZ
        V(0) = o
        V(1).X = o.X + d : V(1).Y = o.Y : V(1).Z = o.Z : V(1).W = 1
        V(2).X = o.X + d : V(2).Y = o.Y + d : V(2).Z = o.Z : V(2).W = 1
        V(3).X = o.X : V(3).Y = o.Y + d : V(3).Z = o.Z : V(3).W = 1
        V(4) = o
        Return V
    End Function
    Private Function yzTile(ByVal o As XYZ, ByVal d As Double) As XYZ()
        Dim V(4) As XYZ
        V(0) = o
        V(1).X = o.X : V(1).Y = o.Y + d : V(1).Z = o.Z : V(1).W = 1
        V(2).X = o.X : V(2).Y = o.Y + d : V(2).Z = o.Z + d : V(2).W = 1
        V(3).X = o.X : V(3).Y = o.Y : V(3).Z = o.Z + d : V(3).W = 1
        V(4) = o
        Return V
    End Function

    Private Function toolLine(ByVal m As XYZ, ByVal p As XYZ, ByVal r As Double, ByVal o As XYZ) As XYZ()
        Dim V(1) As XYZ
        V(0).X = o.X + m.X : V(0).Y = o.Y + m.Y : V(0).Z = o.Z + m.Z : V(0).W = 1
        V(1).X = o.X + p.X : V(1).Y = o.Y + p.Y : V(1).Z = o.Z + p.Z : V(1).W = 1
        '------------------------------------------------------------------------------------
        Return V
    End Function
    Private Function toolSquare(ByVal r As Double, ByVal o As XYZ, ByVal m As XYZ, ByVal p As XYZ) As XYZ()
        Dim tsq(4) As XYZ
        Dim x, y, z As Double
        If cbLockX.Checked = True Then
            x = r : y = 0 : z = 0
        End If
        If cbLockY.Checked = True Then
            x = 0 : y = r : z = 0
        End If
        If cbLockZ.Checked = True Then
            x = 0 : y = 0 : z = r
        End If
        tsq(0).X = o.X + m.X : tsq(0).Y = o.Y + m.Y : tsq(0).Z = o.Z + m.Z : tsq(0).W = 1
        tsq(1).X = o.X + p.X : tsq(1).Y = o.Y + p.Y : tsq(1).Z = o.Z + p.Z : tsq(1).W = 1
        tsq(2).X = o.X + p.X + x : tsq(2).Y = o.Y + p.Y + y : tsq(2).Z = o.Z + p.Z + z : tsq(2).W = 1
        tsq(3).X = o.X + m.X + x : tsq(3).Y = o.Y + m.Y + y : tsq(3).Z = o.Z + m.Z + z : tsq(3).W = 1
        tsq(4) = tsq(0)
        '------------------------------------------------------------------------------------
        Return tsq
    End Function
    Private Function toolCompass(ByVal o As XYZ, ByVal m As XYZ, ByVal p As XYZ) As XYZ()
        Dim tsq(2) As XYZ
        tsq(0).X = o.X : tsq(0).Y = o.Y : tsq(0).Z = o.Z : tsq(0).W = 1
        tsq(1).X = m.X : tsq(1).Y = m.Y : tsq(1).Z = m.Z : tsq(1).W = 1
        tsq(2).X = p.X : tsq(2).Y = p.Y : tsq(2).Z = p.Z : tsq(2).W = 1
        '------------------------------------------------------------------------------------
        Return tsq
    End Function
    Private Function toolFace(ByVal i As Integer, ByVal o As XYZ, ByVal x As Double, ByVal y As Double, ByVal z As Double) As XYZ()
        Dim tsq(4) As XYZ
        If i = 1 Then
            tsq(0).X = o.X : tsq(0).Y = o.Y : tsq(0).Z = o.Z : tsq(0).W = 1
            tsq(1).X = o.X + x : tsq(1).Y = o.Y : tsq(1).Z = o.Z : tsq(1).W = 1
            tsq(2).X = o.X + x : tsq(2).Y = o.Y : tsq(2).Z = o.Z + z : tsq(2).W = 1
            tsq(3).X = o.X : tsq(3).Y = o.Y : tsq(3).Z = o.Z + z : tsq(3).W = 1
            tsq(4) = tsq(0)
        ElseIf i = 2 Then
            tsq(0).X = o.X : tsq(0).Y = o.Y : tsq(0).Z = o.Z : tsq(0).W = 1
            tsq(1).X = o.X + x : tsq(1).Y = o.Y : tsq(1).Z = o.Z : tsq(1).W = 1
            tsq(2).X = o.X + x : tsq(2).Y = o.Y + y : tsq(2).Z = o.Z : tsq(2).W = 1
            tsq(3).X = o.X : tsq(3).Y = o.Y + y : tsq(3).Z = o.Z : tsq(3).W = 1
            tsq(4) = tsq(0)
        ElseIf i = 3 Then
            tsq(0).X = o.X : tsq(0).Y = o.Y : tsq(0).Z = o.Z : tsq(0).W = 1
            tsq(1).X = o.X : tsq(1).Y = o.Y + y : tsq(1).Z = o.Z : tsq(1).W = 1
            tsq(2).X = o.X : tsq(2).Y = o.Y + y : tsq(2).Z = o.Z + z : tsq(2).W = 1
            tsq(3).X = o.X : tsq(3).Y = o.Y : tsq(3).Z = o.Z + z : tsq(3).W = 1
            tsq(4) = tsq(0)
        End If

        '------------------------------------------------------------------------------------
        Return tsq
    End Function
    Private Function toolCube(ByVal o As XYZ, ByVal x As Double, ByVal y As Double, ByVal z As Double) As MDATA()
        Dim cube As MDATA = Nothing
        Dim tC(5) As MDATA
        '-------------------------------------'
        tC(0).Mesh = toolFace(1, vCompile(valueVOX, valueVOY, valueVOZ), x, 0, z)
        tC(1).Mesh = toolFace(1, vCompile(valueVOX, valueVOY + y, valueVOZ), x, 0, z)
        tC(2).Mesh = toolFace(2, vCompile(valueVOX, valueVOY, valueVOZ), x, y, 0)
        tC(3).Mesh = toolFace(2, vCompile(valueVOX, valueVOY, valueVOZ + z), x, y, 0)
        tC(4).Mesh = toolFace(3, vCompile(valueVOX, valueVOY, valueVOZ), 0, y, z)
        tC(5).Mesh = toolFace(3, vCompile(valueVOX + x, valueVOY, valueVOZ), 0, y, z)
        '-------------------------------------'
        For n = 0 To 5
            tC(n).LineColor = cColor(valueVPA, valueVPR, valueVPG, valueVPB)
            tC(n).FillColor = cColor(valueVPA, valueVPR, valueVPG, valueVPB)
            tC(n).Origin = o
            tC(n).Center = o
            tC(n).Line = True
            tC(n).Fill = True
            tC(n).Curve = False
        Next
        Return tC
    End Function
    Private Function toolOverlay(ByVal x As Single, ByVal y As Single, ByVal i1 As Single, ByVal i2 As Single) As PointF()
        Dim pF(8) As PointF
        pF(0).X = i1 : pF(0).Y = i2
        pF(1).X = i2 : pF(1).Y = i1
        pF(2).X = x - i2 : pF(2).Y = i1
        pF(3).X = x - i1 : pF(3).Y = i2
        pF(4).X = x - i1 : pF(4).Y = y - i2
        pF(5).X = x - i2 : pF(5).Y = y - i1
        pF(6).X = i2 : pF(6).Y = y - i1
        pF(7).X = i1 : pF(7).Y = y - i2
        pF(8).X = i1 : pF(8).Y = i2
        'pF(9).X = 0 : pF(9).Y = 0
        'pF(10).X = 0 : pF(10).Y = 0
        Return pF
    End Function
    Private Function toolTriangle(ByVal r As Double, ByVal o As XYZ, ByVal p As XYZ, ByVal m As XYZ) As XYZ()
        Dim tsq(3) As XYZ
        Dim x, y, z As Double
        If cbLockX.Checked = True Then
            x = r : y = 0 : z = 0
        End If
        If cbLockY.Checked = True Then
            x = 0 : y = r : z = 0
        End If
        If cbLockZ.Checked = True Then
            x = 0 : y = 0 : z = r
        End If
        tsq(0).X = o.X + m.X : tsq(0).Y = o.Y + m.Y : tsq(0).Z = o.Z + m.Z : tsq(0).W = 1
        tsq(1).X = o.X + p.X : tsq(1).Y = o.Y + p.Y : tsq(1).Z = o.Z + p.Z : tsq(1).W = 1
        tsq(2).X = o.X + x : tsq(2).Y = o.Y + y : tsq(2).Z = o.Z + z : tsq(2).W = 1
        tsq(3).X = o.X + m.X : tsq(3).Y = o.Y + m.Y : tsq(3).Z = o.Z + m.Z : tsq(3).W = 1
        '------------------------------------------------------------------------------------
        Return tsq
    End Function
#End Region
#Region "Simplification"
    Private Function ValueC(ByVal V As Double, ByVal i As Double, ByVal min As Double, ByVal max As Double) As Double
        Dim tV As Double = V
        tV += i
        If tV < min Then tV = min
        If tV > max Then tV = max
        Return tV
    End Function
    Private Function cColor(ByVal a As Integer, ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As Color
        Dim C As Color = Color.FromArgb(a, r, g, b)
        Return C
    End Function
#End Region
#Region "Calc"
    Private Function cross(ByVal a As XYZ, ByVal b As XYZ) As XYZ
        Dim cp As XYZ
        cp.X = a.Y * b.Z - b.Y * a.Z
        cp.Y = a.Z * b.X - b.Z * a.X
        cp.Z = a.X * b.Y - b.X * a.Y
        cp.W = 1
        Return cp
    End Function
    Private Function dot(ByVal a As XYZ, ByVal b As XYZ) As Double
        Dim dp As Double
        dp = a.X * b.X + a.Y * b.Y + a.Z * b.Z
        Return dp
    End Function
    Private Function minus(ByVal A As XYZ, ByVal B As XYZ) As XYZ
        Dim C As XYZ
        C.X = A.X - B.X
        C.Y = A.Y - B.Y
        C.Z = A.Z - B.Z
        C.W = 1
        Return C
    End Function
    Private Function norm(ByVal A As XYZ) As XYZ
        Dim C As XYZ
        C.X = A.X / A.W
        C.Y = A.Y / A.W
        C.Z = A.Z / A.W
        C.W = 1
        Return C
    End Function
    Private Function distance(ByVal p1 As XYZ, ByVal p2 As XYZ) As Double
        Dim d As Double
        Dim td As Double
        td = (Math.Sqrt(((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2)))
        d = Abs(td)
        Return d
    End Function
    Private Function mXv(ByVal A As MTX, ByVal B As XYZ) As XYZ
        Dim v As XYZ
        v.X = (A.c1r1 * B.X) + (A.c2r1 * B.Y) + (A.c3r1 * B.Z) + (A.c4r1 * B.W)
        v.Y = (A.c1r2 * B.X) + (A.c2r2 * B.Y) + (A.c3r2 * B.Z) + (A.c4r2 * B.W)
        v.Z = (A.c1r3 * B.X) + (A.c2r3 * B.Y) + (A.c3r3 * B.Z) + (A.c4r3 * B.W)
        v.W = (A.c1r4 * B.X) + (A.c2r4 * B.Y) + (A.c3r4 * B.Z) + (A.c4r4 * B.W)
        Return v
    End Function
    Private Function mXm(ByVal A As MTX, ByVal B As MTX) As MTX
        Dim C As MTX
        C.c1r1 = (A.c1r1 * B.c1r1) + (A.c2r1 * B.c1r2) + (A.c3r1 * B.c1r3) + (A.c4r1 * B.c1r4)
        C.c1r2 = (A.c1r2 * B.c1r1) + (A.c2r2 * B.c1r2) + (A.c3r2 * B.c1r3) + (A.c4r2 * B.c1r4)
        C.c1r3 = (A.c1r3 * B.c1r1) + (A.c2r3 * B.c1r2) + (A.c3r3 * B.c1r3) + (A.c4r3 * B.c1r4)
        C.c1r4 = (A.c1r4 * B.c1r1) + (A.c2r4 * B.c1r2) + (A.c3r4 * B.c1r3) + (A.c4r4 * B.c1r4)
        C.c2r1 = (A.c1r1 * B.c2r1) + (A.c2r1 * B.c2r2) + (A.c3r1 * B.c2r3) + (A.c4r1 * B.c2r4)
        C.c2r2 = (A.c1r2 * B.c2r1) + (A.c2r2 * B.c2r2) + (A.c3r2 * B.c2r3) + (A.c4r2 * B.c2r4)
        C.c2r3 = (A.c1r3 * B.c2r1) + (A.c2r3 * B.c2r2) + (A.c3r3 * B.c2r3) + (A.c4r3 * B.c2r4)
        C.c2r4 = (A.c1r4 * B.c2r1) + (A.c2r4 * B.c2r2) + (A.c3r4 * B.c2r3) + (A.c4r4 * B.c2r4)
        C.c3r1 = (A.c1r1 * B.c3r1) + (A.c2r1 * B.c3r2) + (A.c3r1 * B.c3r3) + (A.c4r1 * B.c3r4)
        C.c3r2 = (A.c1r2 * B.c3r1) + (A.c2r2 * B.c3r2) + (A.c3r2 * B.c3r3) + (A.c4r2 * B.c3r4)
        C.c3r3 = (A.c1r3 * B.c3r1) + (A.c2r3 * B.c3r2) + (A.c3r3 * B.c3r3) + (A.c4r3 * B.c3r4)
        C.c3r4 = (A.c1r4 * B.c3r1) + (A.c2r4 * B.c3r2) + (A.c3r4 * B.c3r3) + (A.c4r4 * B.c3r4)
        C.c4r1 = (A.c1r1 * B.c4r1) + (A.c2r1 * B.c4r2) + (A.c3r1 * B.c4r3) + (A.c4r1 * B.c4r4)
        C.c4r2 = (A.c1r2 * B.c4r1) + (A.c2r2 * B.c4r2) + (A.c3r2 * B.c4r3) + (A.c4r2 * B.c4r4)
        C.c4r3 = (A.c1r3 * B.c4r1) + (A.c2r3 * B.c4r2) + (A.c3r3 * B.c4r3) + (A.c4r3 * B.c4r4)
        C.c4r4 = (A.c1r4 * B.c4r1) + (A.c2r4 * B.c4r2) + (A.c3r4 * B.c4r3) + (A.c4r4 * B.c4r4)
        Return C
    End Function
    Private Function project(ByVal V As XYZ, ByVal E As XYZ) As PointF
        Dim P As PointF
        Dim px, py As Double
        py = (distance(V, E) * V.Y) / (distance(V, E) + V.Z)
        px = (distance(V, E) * V.X) / (distance(V, E) + V.Z)
        P = New PointF(px, py)
        Return P
    End Function
    Private Function MI() As MTX
        Dim C As MTX
        C.c1r1 = 1 : C.c2r1 = 0 : C.c3r1 = 0 : C.c4r1 = 0
        C.c1r2 = 0 : C.c2r2 = 1 : C.c3r2 = 0 : C.c4r2 = 0
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = 1 : C.c4r3 = 0
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
    Private Function MPO(ByVal v As XYZ) As MTX
        Dim C As MTX
        C.c1r1 = 1 : C.c2r1 = 0 : C.c3r1 = 0 : C.c4r1 = 0
        C.c1r2 = 0 : C.c2r2 = 1 : C.c3r2 = 0 : C.c4r2 = 0
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = 0 : C.c4r3 = 1 / distance(v, wEye)
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
    Private Function ROX(ByVal a As Double)
        Dim C As MTX
        C.c1r1 = 1 : C.c2r1 = 0 : C.c3r1 = 0 : C.c4r1 = 0
        C.c1r2 = 0 : C.c2r2 = Cos(a) : C.c3r2 = Sin(a) : C.c4r2 = 0
        C.c1r3 = 0 : C.c2r3 = -Sin(a) : C.c3r3 = Cos(a) : C.c4r3 = 0
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
    Private Function ROY(ByVal a As Double)
        Dim C As MTX
        C.c1r1 = Cos(a) : C.c2r1 = 0 : C.c3r1 = -Sin(a) : C.c4r1 = 0
        C.c1r2 = 0 : C.c2r2 = 1 : C.c3r2 = 0 : C.c4r2 = 0
        C.c1r3 = Sin(a) : C.c2r3 = 0 : C.c3r3 = Cos(a) : C.c4r3 = 0
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
    Private Function ROZ(ByVal a As Double)
        Dim C As MTX
        C.c1r1 = Cos(a) : C.c2r1 = Sin(a) : C.c3r1 = 0 : C.c4r1 = 0
        C.c1r2 = -Sin(a) : C.c2r2 = Cos(a) : C.c3r2 = 0 : C.c4r2 = 0
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = 1 : C.c4r3 = 0
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
    Private Function MTR(ByVal V As XYZ) As MTX
        Dim C As MTX
        C.c1r1 = 1 : C.c2r1 = 0 : C.c3r1 = 0 : C.c4r1 = V.X
        C.c1r2 = 0 : C.c2r2 = 1 : C.c3r2 = 0 : C.c4r2 = V.Y
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = 1 : C.c4r3 = V.Z
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
    Private Function MTB(ByVal V As XYZ) As MTX
        Dim C As MTX
        C.c1r1 = 1 : C.c2r1 = 0 : C.c3r1 = 0 : C.c4r1 = V.X * -1
        C.c1r2 = 0 : C.c2r2 = 1 : C.c3r2 = 0 : C.c4r2 = V.Y * -1
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = 1 : C.c4r3 = V.Z * -1
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
    Private Function MSC(ByVal V As XYZ) As MTX
        Dim C As MTX
        C.c1r1 = V.X : C.c2r1 = 0 : C.c3r1 = 0 : C.c4r1 = 0
        C.c1r2 = 0 : C.c2r2 = V.Y : C.c3r2 = 0 : C.c4r2 = 0
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = V.Z : C.c4r3 = 0
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
    Private Function MPR(ByVal a As Double, ByVal w As Double, ByVal h As Double, ByVal zn As Double, ByVal zf As Double) As MTX
        Dim C As MTX
        C.c1r1 = 1 / w * AR : C.c2r1 = 0 : C.c3r1 = 0 : C.c4r1 = 0
        C.c1r2 = 0 : C.c2r2 = 1 / h : C.c3r2 = 0 : C.c4r2 = 0
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = 1 / (zf - zn) : C.c4r3 = -(zn / (zf - zn))
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = C.c4r3 = 0 : C.c4r4 = 1
        Return C
    End Function

    Private Function MFR(ByVal n As Double, ByVal f As Double, ByVal t As Double, ByVal b As Double, ByVal l As Double, ByVal r As Double) As MTX
        Dim C As MTX
        C.c1r1 = 2 * n / (r - l) : C.c2r1 = 0 : C.c3r1 = (r + l) / (r - l) : C.c4r1 = 0
        C.c1r2 = 0 : C.c2r2 = 2 * n / (t - b) : C.c3r2 = (t + b) / (t - b) : C.c4r2 = 0
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = -(f + n) / (f - n) : C.c4r3 = -2 * f * n / (f - n)
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = -1 : C.c4r4 = 1
        Return C
    End Function
    '--------------------------------------------------'
    '| 2/(r-1) |    0    |    0    |   -((r+1)/(r-1)) |'
    '--------------------------------------------------'
    '|    0    | 2(t-b)  |    0    |   -((t+b)/(t-b)) |'
    '--------------------------------------------------'
    '|    0    |    0    |  1(f-n) |   -(n/(f-n))     |'
    '--------------------------------------------------'
    '|    0    |    0    |    0    |        1         |'
    '--------------------------------------------------'
    '___________________________________________________________________'
    '_   Matrix                                                        _'
    '___________________________________________________________________'
    '      1       |       0       |       0       |       0       | X  '
    '_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _'
    '      0       |       1       |       0       |       0       | Y  '
    '_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _'
    '      0       |       0       |       1       |       0       | Z  '
    '_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _'
    '      0       |       0       |       0       |       1       | W  '
    '___________________________________________________________________'
    '_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _'
    '___________________________________________________________________'
    'float r_width  = 1.0f / (right - left);
    'float r_height = 1.0f / (top - bottom);
    'float r_depth  = 1.0f / (far - near);
    'float x =  2.0f * (r_width);
    'float y =  2.0f * (r_height);
    'float z =  2.0f * (r_depth);
    'float A = (right + left) * r_width;
    'float B = (top + bottom) * r_height;
    'float C = (far + near) * r_depth;
    'm[offset + 0] = x;
    'm[offset + 3] = -A;
    'm[offset + 5] = y;
    'm[offset + 7] = -B;
    'm[offset + 10] = -z;
    'm[offset + 11] = -C;

    Private Function MPP(ByVal t As Double, ByVal b As Double, ByVal r As Double, ByVal n As Double, ByVal f As Double) As MTX
        Dim C As MTX
        C.c1r1 = 2 / (r - 1) : C.c2r1 = 0 : C.c3r1 = 0 : C.c4r1 = -((r + 1) / (r - 1))
        C.c1r2 = 0 : C.c2r2 = 2 / (t - b) : C.c3r2 = 0 : C.c4r2 = -((t + b) / (t - b))
        C.c1r3 = 0 : C.c2r3 = 0 : C.c3r3 = 1 / (f - n) : C.c4r3 = -(n / (f - n))
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function

    Private Function MTP(ByVal M As MTX) As MTX
        Dim C As MTX
        C.c1r1 = M.c1r1 : C.c2r1 = M.c1r2 : C.c3r1 = M.c1r3 : C.c4r1 = M.c1r4
        C.c1r2 = M.c2r1 : C.c2r2 = M.c2r2 : C.c3r2 = M.c2r3 : C.c4r2 = M.c2r4
        C.c1r3 = M.c3r1 : C.c2r3 = M.c3r2 : C.c3r3 = M.c3r3 : C.c4r3 = M.c3r4
        C.c1r4 = M.c4r1 : C.c2r4 = M.c4r2 : C.c3r4 = M.c4r3 : C.c4r4 = M.c4r4
        Return C
    End Function

    Private Function MLA(ByVal x As XYZ, ByVal y As XYZ, ByVal z As XYZ, ByVal e As XYZ) As MTX
        Dim C As MTX
        Dim V As XYZ = vCompile(-e.X, -e.Y, -e.Z)
        C.c1r1 = x.X : C.c2r1 = y.X : C.c3r1 = z.X : C.c4r1 = dot(x, V)
        C.c1r2 = x.Y : C.c2r2 = y.Y : C.c3r2 = z.Y : C.c4r2 = dot(y, V)
        C.c1r3 = x.Z : C.c2r3 = y.Z : C.c3r3 = z.Z : C.c4r3 = dot(z, V)
        C.c1r4 = 0 : C.c2r4 = 0 : C.c3r4 = 0 : C.c4r4 = 1
        Return C
    End Function
#End Region
#Region "Runtime"
    Private Sub gPress_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gPress.Tick
        clock.Start()
        CamMotion()
        MainPress()
        ModMotion()
        SpeedDraw()
        MainDisplay.Image = gImage
        SpeedModel()
        ModelDisplay.Image = mImage
        proTime = clock.ElapsedMilliseconds
        'If proTime = 0 Then proTime = 1
        clock.Reset()
    End Sub
#End Region
#Region "Data Press"
    Private Sub StepMicro(ByVal la As XYZ())
        If IsNothing(MicroData) = False Then
            ReDim Preserve MicroData(MicroData.Length)
            MicroData(MicroData.Length - 1).Mesh = la
            MicroData(MicroData.Length - 1).LineColor = cColor(valueVPA, valueVPR, valueVPG, valueVPB)
            MicroData(MicroData.Length - 1).FillColor = cColor(valueVPFA, valueVPFR, valueVPFG, valueVPFB)
            MicroData(MicroData.Length - 1).Origin = vCompile(valueVOX, valueVOY, valueVOZ)
            MicroData(MicroData.Length - 1).Label = tbLabel.Text
            MicroData(MicroData.Length - 1).Font = tbLabel.Font
            MicroData(MicroData.Length - 1).Radian = valueVRD
            MicroData(MicroData.Length - 1).SpinAngleX = valueVSX
            MicroData(MicroData.Length - 1).SpinAngleY = valueVSY
            MicroData(MicroData.Length - 1).SpinAngleZ = valueVSZ
            MicroData(MicroData.Length - 1).OrbitAngleX = valueVTX
            MicroData(MicroData.Length - 1).OrbitAngleY = valueVTY
            MicroData(MicroData.Length - 1).OrbitAngleZ = valueVTZ
            If cbApplySpin.Checked = True Then
                MicroData(MicroData.Length - 1).SpinIO = True
            End If
            If cbApplyOrbit.Checked = True Then
                MicroData(MicroData.Length - 1).OrbitIO = True
            End If
            If cbFillplane.Checked = True Then
                MicroData(MicroData.Length - 1).Fill = True
            Else
                MicroData(MicroData.Length - 1).Fill = False
            End If
            If cbDrawmesh.Checked = True Then
                MicroData(MicroData.Length - 1).Line = True
            Else
                MicroData(MicroData.Length - 1).Line = False
            End If
            If cbComp.Checked = True Then
                MicroData(MicroData.Length - 1).Curve = True
            Else
                MicroData(MicroData.Length - 1).Curve = False
            End If
            If cbLabel.Checked = True Then
                MicroData(MicroData.Length - 1).LabelIO = True
            Else
                MicroData(MicroData.Length - 1).LabelIO = False
            End If
            MicroData = aPress(MicroData)
            '------------------------------------------
            valueMacArray = MicroData.Length
            valueMicArray = 1
            valueAT = MicroData(0).Mesh.Length
            valueAS = 1
            tbAdjustSelect.Text = valueAS.ToString
            tbAdjustTotal.Text = valueAT.ToString
            tbAdjustMacro.Text = valueMacArray.ToString
            tbAdjustMicro.Text = valueMicArray.ToString
            valueVAX = MicroData(0).Mesh(0).X
            valueVAY = MicroData(0).Mesh(0).Y
            valueVAZ = MicroData(0).Mesh(0).Z
            tbAdjustVX.Text = valueVAX.ToString
            tbAdjustVY.Text = valueVAY.ToString
            tbAdjustVZ.Text = valueVAZ.ToString
        Else
            ReDim Preserve MicroData(0)
            MicroData(0).Mesh = la
            MicroData(0).LineColor = cColor(valueVPA, valueVPR, valueVPG, valueVPB)
            MicroData(0).FillColor = cColor(valueVPFA, valueVPFR, valueVPFG, valueVPFB)
            MicroData(0).Origin = vCompile(valueVOX, valueVOY, valueVOZ)
            MicroData(0).Label = tbLabel.Text
            MicroData(0).Font = tbLabel.Font
            MicroData(0).Radian = valueVRD
            MicroData(0).SpinAngleX = valueVSX
            MicroData(0).SpinAngleY = valueVSY
            MicroData(0).SpinAngleZ = valueVSZ
            MicroData(0).OrbitAngleX = valueVTX
            MicroData(0).OrbitAngleY = valueVTY
            MicroData(0).OrbitAngleZ = valueVTZ
            If cbApplySpin.Checked = True Then
                MicroData(0).SpinIO = True
            End If
            If cbApplyOrbit.Checked = True Then
                MicroData(0).OrbitIO = True
            End If
            If cbFillplane.Checked = True Then
                MicroData(0).Fill = True
            Else
                MicroData(0).Fill = False
            End If
            If cbDrawmesh.Checked = True Then
                MicroData(0).Line = True
            Else
                MicroData(0).Line = False
            End If
            If cbComp.Checked = True Then
                MicroData(0).Curve = True
            Else
                MicroData(0).Curve = False
            End If
            If cbLabel.Checked = True Then
                MicroData(0).LabelIO = True
            Else
                MicroData(0).LabelIO = False
            End If
            MicroData = aPress(MicroData)
            '------------------------------------------
            valueMacArray = MicroData.Length
            valueMicArray = 1
            valueAT = MicroData(0).Mesh.Length
            valueAS = 1
            tbAdjustSelect.Text = valueAS.ToString
            tbAdjustTotal.Text = valueAT.ToString
            tbAdjustMacro.Text = valueMacArray.ToString
            tbAdjustMicro.Text = valueMicArray.ToString
            valueVAX = MicroData(0).Mesh(0).X
            valueVAY = MicroData(0).Mesh(0).Y
            valueVAZ = MicroData(0).Mesh(0).Z
            tbAdjustVX.Text = valueVAX.ToString
            tbAdjustVY.Text = valueVAY.ToString
            tbAdjustVZ.Text = valueVAZ.ToString
        End If
    End Sub
    Private Sub PressMacro(ByVal wa As MDATA())
        If IsNothing(MacroData) = False Then
            ReDim Preserve MacroData(MacroData.Length)
            MacroData(MacroData.Length - 1) = wa
            valueMacArray = 0
            valueMicArray = 0
            valueAT = 0
            valueAS = 0
            tbAdjustSelect.Text = valueAS.ToString
            tbAdjustTotal.Text = valueAT.ToString
            tbAdjustMacro.Text = valueMacArray.ToString
            tbAdjustMicro.Text = valueMicArray.ToString
            valueVAX = 0
            valueVAY = 0
            valueVAZ = 0
            tbAdjustVX.Text = valueVAX.ToString
            tbAdjustVY.Text = valueVAY.ToString
            tbAdjustVZ.Text = valueVAZ.ToString
        Else
            ReDim Preserve MacroData(0)
            MacroData(0) = wa
            valueMacArray = 0
            valueMicArray = 0
            valueAT = 0
            valueAS = 0
            tbAdjustSelect.Text = valueAS.ToString
            tbAdjustTotal.Text = valueAT.ToString
            tbAdjustMacro.Text = valueMacArray.ToString
            tbAdjustMicro.Text = valueMicArray.ToString
            valueVAX = 0
            valueVAY = 0
            valueVAZ = 0
            tbAdjustVX.Text = valueVAX.ToString
            tbAdjustVY.Text = valueVAY.ToString
            tbAdjustVZ.Text = valueVAZ.ToString
        End If
    End Sub
#End Region
#Region "camera control"
    Private Sub ModMotion()
        modView = mXm(modView, ROY(rad))
        mEye = norm(mXv(ROY(-rad), mEye))
    End Sub
    Private Sub CamMotion()
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        Dim tM As MTX = mView
        If bViewUp = True Then
            mView = mXm(ROX(rad), mView)
            wEye = norm(mXv((ROX(-rad)), wEye))
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If bViewDown = True Then
            mView = mXm(ROX(-rad), mView)
            wEye = norm(mXv((ROX(rad)), wEye))
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If bViewLeft = True Then
            mView = mXm(mView, ROY(rad))
            wEye = norm(mXv(ROY(-rad), wEye))
            'wEye = norm(mXv(MTP(ROY(rad)), wEye))
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If bViewRight = True Then
            mView = mXm(mView, ROY(-rad))
            wEye = norm(mXv(ROY(rad), wEye))
            'wEye = norm(mXv(MTP(ROY(-rad)), wEye))
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If bZoomIn = True Then
            mView = mXm(MSC(vCompile(1.04, 1.04, 1.04)), mView)
            'camDist += 10
        End If
        If bZoomOut = True Then
            mView = mXm(MSC(vCompile(0.96, 0.96, 0.96)), mView)
            'camDist -= 10
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If bShiftU = True Then
            ScrY += 20
        End If
        If bShiftD = True Then
            ScrY -= 20
        End If
        If bShiftL = True Then
            ScrX += 20
        End If
        If bShiftR = True Then
            ScrX -= 20
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    End Sub
    Private Sub btnCamTarget_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCamTarget.Click
        Static i As Integer = 0
        ScrX = MainDisplay.Width
        ScrY = MainDisplay.Height
        mView = MI()
        wEye = vCompile(0, 0, camDist)
        If i = 0 Then
            mView = MI()
            wEye = vCompile(0, 0, camDist)
            tbLog.AppendText(vbNewLine + "Quick view front.")
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If i = 1 Then
            mView = mXm(ROX(_Rad90), mView)
            wEye = norm(mXv((ROX(-_Rad90)), wEye))
            tbLog.AppendText(vbNewLine + "Quick view top")
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If i = 2 Then
            mView = mXm(mView, ROY(-_Rad90))
            wEye = norm(mXv(ROY(_Rad90), wEye))
            tbLog.AppendText(vbNewLine + "Quick view right")
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If i = 3 Then
            mView = mXm(mView, ROY(_Rad180))
            wEye = norm(mXv(ROY(-_Rad180), wEye))
            tbLog.AppendText(vbNewLine + "Quick view back")
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If i = 4 Then
            mView = mXm(mView, ROY(_Rad90))
            wEye = norm(mXv(ROY(-_Rad90), wEye))
            tbLog.AppendText(vbNewLine + "Quick view left")
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        If i = 5 Then
            mView = mXm(mView, ROY(-_Rad45))
            wEye = norm(mXv(ROY(_Rad45), wEye))
            mView = mXm(ROX(_Rad15), mView)
            wEye = norm(mXv((ROX(-_Rad15)), wEye))
            tbLog.AppendText(vbNewLine + "Quick view angle")
        End If
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        i += 1
        If i = 6 Then i = 0
        '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    End Sub

    Private Sub btnCamSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCamSave.Click
        SaveImage()
    End Sub
#End Region
#Region "Graphics Press"
    Private Function oPress(ByVal v As XYZ, ByVal w As XYZ()) As XYZ()
        Dim M As MTX = mXm(ROX(v.X), mXm(ROY(v.Y), (ROZ(v.Z))))
        Dim tWv() As XYZ = w
        For x = 0 To w.Length - 1
            tWv(x) = mXv(M, tWv(x))
        Next
        Return tWv
    End Function
    Private Function rPress(ByVal v As XYZ, ByVal w As XYZ(), ByVal o As XYZ) As XYZ()
        Dim M As MTX = mXm(ROX(v.X), mXm(ROY(v.Y), (ROZ(v.Z))))
        Dim tWv() As XYZ = w
        For x = 0 To w.Length - 1
            tWv(x) = mXv(MTR(o), mXv(M, mXv(MTB(o), tWv(x))))
        Next
        Return tWv
    End Function
    Private Function tPress(ByVal v As XYZ, ByVal w As XYZ()) As XYZ()
        Dim tWv(w.Length - 1) As XYZ
        For x = 0 To w.Length - 1
            tWv(x) = norm(mXv(MTR(v), w(x)))
        Next
        Return tWv
    End Function
    Private Function vPress(ByVal m As MTX, ByVal w As XYZ()) As XYZ()
        Dim tWv(w.Length - 1) As XYZ
        For x = 0 To w.Length - 1
            tWv(x) = mXv(m, norm(w(x)))
        Next
        Return tWv
    End Function
    Private Function mvPress(ByVal w As XYZ()) As XYZ()
        Dim tWv(w.Length - 1) As XYZ
        For x = 0 To w.Length - 1
            tWv(x) = mXv(modView, norm(w(x)))
        Next
        Return tWv
    End Function
    Private Function pPress(ByVal v As XYZ()) As PointF()
        Dim tPF(v.Length - 1) As PointF
        For x = 0 To v.Length - 1
            tPF(x) = project(v(x), wEye)
            tPF(x).X = tPF(x).X + ((ScrX) / 2)
            tPF(x).Y = -tPF(x).Y + ((ScrY) / 2)
        Next
        Return tPF
    End Function
    Private Function mPress(ByVal v As XYZ()) As PointF()
        Dim tPF(v.Length - 1) As PointF
        For x = 0 To v.Length - 1
            tPF(x) = project(v(x), mEye)
            tPF(x).X = tPF(x).X + (ScrMX / 2)
            tPF(x).Y = -tPF(x).Y + (ScrMY / 2)
        Next
        Return tPF
    End Function
    Private Function qSort(ByVal list As MDATA())
        Return list.OrderByDescending(Function(c) c.Distance).ToArray
    End Function

    Private Function nPress(ByVal m As MDATA(), ByVal p As XYZ) As MDATA()
        Dim t() As MDATA = m
        For i = 0 To m.Length - 1
            t(i).Distance = distance(dPress(m(i).Mesh), p)
        Next
        Return t
    End Function

    Private Function dPress(ByVal m As XYZ()) As XYZ
        Dim R As XYZ
        Dim i, j, k As Double
        i = 0 : j = 0 : k = 0
        For x = 0 To m.Length - 1
            i += m(x).X : j += m(x).Y : k += m(x).Z
        Next x
        R = vCompile(i / m.Length, j / m.Length, k / m.Length)
        Return R
    End Function
    Private Function aPress(ByVal m As MDATA()) As MDATA()
        Dim t() As MDATA = qSort(nPress(MicroData, wOrigin))
        Dim q = (t.Length / 2)
        t(q).Center = t(q).Origin
        For x = 0 To m.Length - 1
            If x <> q Then
                t(x).Center = t(q).Center
            End If
        Next
        Return t
    End Function
    Private Function cPress(ByVal u As MDATA()()) As MDATA()
        Dim MD()() As MDATA = u
        Dim n As Integer = 0
        Dim S() As MDATA = Nothing
        For x = 0 To MD.Length - 1
            For y = 0 To MD(x).Length - 1
                ReDim Preserve S(n)
                S(n) = MD(x)(y)
                n += 1
            Next y
        Next x
        Return S
    End Function
    Private Sub MainPress()
        If IsNothing(MacroData) = False Then
            Dim mac() As MDATA = cPress(MacroData)
            For Each m As MDATA In mac
                If m.SpinIO = True Then
                    m.Mesh = rPress(vCompile(m.SpinAngleX, m.SpinAngleY, m.SpinAngleZ), (m.Mesh), m.Center)
                End If
                If m.OrbitIO = True Then
                    m.Mesh = oPress(vCompile(m.OrbitAngleX, m.OrbitAngleY, m.OrbitAngleZ), m.Mesh)
                End If
            Next
        End If
    End Sub
#End Region
#Region "Graphics Paint"
    Private Sub SpeedDraw()
        gr.Clear(bcC)
        gr.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        gr.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        gr.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        '-----------------------------------------------------'
        Dim aPen As Pen = New Pen(cColor(200, 0, 0, 255), 0)
        Dim gridPen As Pen = New Pen(cColor(150, 0, 0, 150), 0)
        Dim tPen As Pen = New Pen(fcC)
        tPen.DashStyle = DashStyle.DashDot
        Dim gBrush As SolidBrush = New SolidBrush(cColor(valueVPFA, valueVPFR, valueVPFG, valueVPFB))
        Dim gPen As Pen = New Pen(cColor(valueVPA, valueVPR, valueVPG, valueVPB), 0)
        Dim aBrush As SolidBrush = New SolidBrush(cColor(valueVPA, valueVPR, valueVPG, valueVPB))
        Dim tBrush As SolidBrush = New SolidBrush(fcC)
        VertexCount = 0
        '-----------------------------------------------------'
        If cbGridXZ.Checked = True Then
            gridPen.Color = cColor(150, 0, 0, 150)
            For Each m As XYZ() In GridDataXZ
                gr.DrawLines(gridPen, pPress(vPress(mView, m)))
                VertexCount += m.Length
            Next
        End If
        If cbGridXY.Checked = True Then
            gridPen.Color = cColor(150, 0, 150, 0)
            For Each m As XYZ() In GridDataXY
                gr.DrawLines(gridPen, pPress(vPress(mView, m)))
                VertexCount += m.Length
            Next
        End If
        If cbGridYZ.Checked = True Then
            gridPen.Color = cColor(150, 150, 0, 0)
            For Each m As XYZ() In GridDataYZ
                gr.DrawLines(gridPen, pPress(vPress(mView, m)))
                VertexCount += m.Length
            Next
        End If
        '-----------------------------------------------------'
        If cbAxisLines.Checked = True Then
            Dim tempL() As PointF
            tempL = pPress(vPress(mView, drawAL(valueVGS + 30)))
            aPen.Color = cColor(200, 0, 0, 150)
            aBrush.Color = cColor(200, 0, 0, 150)
            gr.DrawLines(aPen, pPress(vPress(mView, drawXA(valueVGS + 25))))
            gr.DrawString("- X", New Font("Lucida Console", 10), aBrush, tempL(0))
            gr.DrawString("+ X", New Font("Lucida Console", 10), aBrush, tempL(1))
            aPen.Color = cColor(200, 0, 150, 0)
            aBrush.Color = cColor(200, 0, 150, 0)
            gr.DrawLines(aPen, pPress(vPress(mView, drawYA(valueVGS + 25))))
            gr.DrawString("- Y", New Font("Lucida Console", 10), aBrush, tempL(2))
            gr.DrawString("+ Y", New Font("Lucida Console", 10), aBrush, tempL(3))
            aPen.Color = cColor(200, 150, 0, 0)
            aBrush.Color = cColor(200, 150, 0, 0)
            gr.DrawLines(aPen, pPress(vPress(mView, drawZA(valueVGS + 25))))
            gr.DrawString("- Z", New Font("Lucida Console", 10), aBrush, tempL(4))
            gr.DrawString("+ Z", New Font("Lucida Console", 10), aBrush, tempL(5))
            VertexCount += 6
        End If
        '-----------------------------------------------------'
        If IsNothing(MacroData) = False Then
            Dim TZD() As MDATA = qSort(nPress(cPress(MacroData), wEye))
            For Each m As MDATA In TZD
                gPen.Color = m.LineColor
                gBrush.Color = m.FillColor
                If m.Curve = False Then
                    If m.Fill = True Then
                        gr.FillPolygon(gBrush, pPress(vPress(mView, m.Mesh)))
                        VertexCount += m.Mesh.Length
                    End If
                    If m.Line = True Then
                        gr.DrawPolygon(gPen, pPress(vPress(mView, m.Mesh)))
                        VertexCount += m.Mesh.Length
                    End If
                    If m.LabelIO = True Then
                        gr.DrawPolygon(gPen, pPress(vPress(mView, m.Mesh)))
                        gr.DrawString(m.Label, m.Font, gBrush, pPress(vPress(mView, m.Mesh))(1))
                        VertexCount += m.Mesh.Length
                    End If
                ElseIf m.Curve = True Then
                    If m.Fill = True Then
                        gr.FillClosedCurve(gBrush, pPress(vPress(mView, m.Mesh)), FillMode.Alternate, m.Radian)
                        VertexCount += m.Mesh.Length
                    End If
                    If m.Line = True Then
                        gr.DrawCurve(gPen, pPress(vPress(mView, m.Mesh)), m.Radian)
                        VertexCount += m.Mesh.Length
                    End If
                End If
            Next
        End If
        '-----------------------------------------------------'
        If IsNothing(MicroData) = False Then
            Dim TMD() As MDATA = qSort(nPress(MicroData, wEye))
            For Each m As MDATA In TMD
                gPen.Color = m.LineColor
                gBrush.Color = m.FillColor
                If m.Curve = False Then
                    If m.Fill = True Then
                        gr.FillPolygon(gBrush, pPress(vPress(mView, m.Mesh)))
                        VertexCount += m.Mesh.Length
                    End If
                    If m.Line = True Then
                        gr.DrawPolygon(gPen, pPress(vPress(mView, m.Mesh)))
                        VertexCount += m.Mesh.Length
                    End If
                    If m.LabelIO = True Then
                        gr.DrawPolygon(gPen, pPress(vPress(mView, m.Mesh)))
                        gr.DrawString(m.Label, m.Font, gBrush, pPress(vPress(mView, m.Mesh))(1))
                        VertexCount += m.Mesh.Length
                    End If
                ElseIf m.Curve = True Then
                    If m.Fill = True Then
                        gr.FillClosedCurve(gBrush, pPress(vPress(mView, m.Mesh)), FillMode.Alternate, m.Radian)
                        VertexCount += m.Mesh.Length
                    End If
                    If m.Line = True Then
                        gr.DrawCurve(gPen, pPress(vPress(mView, m.Mesh)), m.Radian)
                        VertexCount += m.Mesh.Length
                    End If
                End If
            Next
        End If
        '-----------------------------------------------------'
        gPen.Color = cColor(valueVPA, valueVPR, valueVPG, valueVPB)
        gBrush.Color = cColor(valueVPFA, valueVPFR, valueVPFG, valueVPFB)
        If cbLine.Checked = True Then
            If cbFillplane.Checked = True Then
                gr.FillPolygon(gBrush, pPress(vPress(mView, toolLine(vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ), valueVRD, vCompile(valueVOX, valueVOY, valueVOZ)))))
                VertexCount += 2
            End If
            If cbDrawmesh.Checked = True Then
                gr.DrawLines(gPen, pPress(vPress(mView, toolLine(vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ), valueVPZ, vCompile(valueVOX, valueVOY, valueVOZ)))))
                VertexCount += 2
            End If
        End If
        '-----------------------------------------------------'
        If cbSquare.Checked = True Then
            If cbFillplane.Checked = True Then
                gr.FillPolygon(gBrush, pPress(vPress(mView, toolSquare(valueVRD, vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))))
                VertexCount += 5
            End If
            If cbDrawmesh.Checked = True Then
                gr.DrawPolygon(gPen, pPress(vPress(mView, toolSquare(valueVRD, vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))))
                VertexCount += 5
            End If
        End If
        '-----------------------------------------------------'
        If cbTriangle.Checked = True Then
            If cbFillplane.Checked = True Then
                gr.FillPolygon(gBrush, pPress(vPress(mView, toolTriangle(valueVRD, vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))))
                VertexCount += 4
            End If
            If cbDrawmesh.Checked = True Then
                gr.DrawPolygon(gPen, pPress(vPress(mView, toolTriangle(valueVRD, vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))))
                VertexCount += 4
            End If
        End If
        '-----------------------------------------------------'
        If cbComp.Checked = True Then
            If cbFillplane.Checked = True Then
                gr.FillClosedCurve(gBrush, pPress(vPress(mView, toolCompass(vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))), FillMode.Alternate, valueVRD)
                VertexCount += 3
            End If
            If cbDrawmesh.Checked = True Then
                gr.DrawCurve(gPen, pPress(vPress(mView, toolCompass(vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))), valueVRD)
                VertexCount += 3
            End If
        End If
        '-----------------------------------------------------'
        If cbPrism.Checked = True Then
            If cbFillplane.Checked = True Then
                For Each M As MDATA In toolCube(vCompile(valueVOX, valueVOY, valueVOZ), valueVMX, valueVMY, valueVMZ)
                    gr.FillPolygon(gBrush, pPress(vPress(mView, M.Mesh)))
                    VertexCount += M.Mesh.Length
                Next
            End If
            If cbDrawmesh.Checked = True Then
                For Each M As MDATA In toolCube(vCompile(valueVOX, valueVOY, valueVOZ), valueVMX, valueVMY, valueVMZ)
                    gr.DrawPolygon(gPen, pPress(vPress(mView, M.Mesh)))
                    VertexCount += M.Mesh.Length
                Next
            End If
        End If
        '-----------------------------------------------------'
        If cbLabel.Checked = True Then
            If cbDrawmesh.Checked = True Then
                gr.DrawLines(gPen, pPress(vPress(mView, toolLine(vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ), valueVPZ, vCompile(valueVOX, valueVOY, valueVOZ)))))
                VertexCount += 2
            End If
            Dim tP As PointF = pPress(vPress(mView, toolLine(vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ), valueVPZ, vCompile(valueVOX, valueVOY, valueVOZ))))(1)
            Dim sT As String = tbLabel.Text
            Dim fT As Font = tbLabel.Font
            gr.DrawString(sT, fT, gBrush, tP)
        End If
        '-----------------------------------------------------'
        If FrameInfoToolStripMenuItem.Checked = True Then
            Dim tSP(0) As PointF
            Dim tSV(0) As XYZ
            tSV(0) = vCompile(valueVAX, valueVAY, valueVAZ)
            tSP = pPress(vPress(mView, tSV))
            If IsNothing(MicroData) = False Then gr.DrawEllipse(tPen, tSP(0).X - 10, tSP(0).Y - 10, 20, 20)
            gr.DrawPolygon(tPen, toolOverlay(MainDisplay.Width, MainDisplay.Height, 20, 50))
            gr.DrawPolygon(tPen, toolOverlay(MainDisplay.Width, MainDisplay.Height, 3, 33))
            Dim headline As String
            headline = "Vertices: " + VertexCount.ToString + " Refresh Rate: " + gPress.Interval.ToString + " Process Time: " + proTime.ToString + " FPS: " + (1000 / proTime).ToString("N3")
            '-----------------------------------------------------'
            gr.DrawString(headline, New Font("Lucida Console", 10), tBrush, New PointF(60, 5))
        End If
        '-----------------------------------------------------'
        gridPen.Dispose()
        gPen.Dispose()
        aPen.Dispose()
        tPen.Dispose()
        gBrush.Dispose()
        aBrush.Dispose()
        tBrush.Dispose()
    End Sub
    Private Sub SpeedModel()
        gm.Clear(cB4)
        gm.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        gm.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        Dim aPen As Pen = New Pen(cColor(200, 0, 0, 255), 0)
        Dim gridPen As Pen = New Pen(cColor(150, 0, 0, 150), 0)
        Dim gBrush As SolidBrush = New SolidBrush(cColor(valueVPA, valueVPR, valueVPG, valueVPB))
        Dim gPen As Pen = New Pen(cColor(valueVPA, valueVPR, valueVPG, valueVPB), 0)
        '-----------------------------------------------------'
        If IsNothing(ModelData) = False Then
            Dim TMD() As MDATA = qSort(nPress(ModelData, mEye))
            For Each m As MDATA In TMD
                gPen.Color = m.LineColor
                gBrush.Color = m.FillColor
                If m.Curve = False Then
                    If m.Fill = True Then
                        gm.FillPolygon(gBrush, mPress(mvPress(m.Mesh)))
                    End If
                    If m.Line = True Then
                        gm.DrawPolygon(gPen, mPress(mvPress(m.Mesh)))
                    End If
                    If m.LabelIO = True Then
                        gm.DrawPolygon(gPen, mPress(mvPress(m.Mesh)))
                        gm.DrawString(m.Label, m.Font, gBrush, mPress(mvPress(m.Mesh))(1))
                        VertexCount += m.Mesh.Length
                    End If
                ElseIf m.Curve = True Then
                    If m.Fill = True Then
                        gm.FillClosedCurve(gBrush, mPress(mvPress(m.Mesh)), FillMode.Alternate, m.Radian)
                    End If
                    If m.Line = True Then
                        gm.DrawCurve(gPen, mPress(mvPress(m.Mesh)), m.Radian)
                    End If
                End If
            Next
        End If
        '-----------------------------------------------------'
        gridPen.Dispose()
        gPen.Dispose()
        aPen.Dispose()
        gBrush.Dispose()
    End Sub
    Private Sub MainDisplay_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MainDisplay.Paint
        
    End Sub
    Private Sub ModelDisplay_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles ModelDisplay.Paint
        
    End Sub
    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.Refresh()
    End Sub
#End Region
#Region "checkboard"
    Private Sub cbMicro_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMicro.CheckedChanged
        If cbMicro.Checked = True Then
            cbMacro.Checked = False
            ValueMark = valueMicro
        ElseIf cbMicro.Checked = False Then
            cbMacro.Checked = True
            ValueMark = valueMacro
        End If
    End Sub

    Private Sub cbMacro_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbMacro.CheckedChanged
        If cbMacro.Checked = True Then
            cbMicro.Checked = False
            ValueMark = valueMacro
        ElseIf cbMacro.Checked = False Then
            cbMicro.Checked = True
            ValueMark = valueMicro
        End If
    End Sub
    Private Sub cbSquare_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSquare.CheckedChanged
        If cbSquare.Checked = True Then
            cbComp.Checked = False
            cbPrism.Checked = False
            cbLine.Checked = False
            cbTriangle.Checked = False
            cbLabel.Checked = False
        End If
    End Sub

    Private Sub cbComp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbComp.CheckedChanged
        If cbComp.Checked = True Then
            cbSquare.Checked = False
            cbPrism.Checked = False
            cbLine.Checked = False
            cbTriangle.Checked = False
            cbLabel.Checked = False
        End If
    End Sub

    Private Sub cbPrism_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbPrism.CheckedChanged
        If cbPrism.Checked = True Then
            cbComp.Checked = False
            cbSquare.Checked = False
            cbLine.Checked = False
            cbTriangle.Checked = False
            cbLabel.Checked = False
        End If
    End Sub

    Private Sub cbLine_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbLine.CheckedChanged
        If cbLine.Checked = True Then
            cbComp.Checked = False
            cbSquare.Checked = False
            cbPrism.Checked = False
            cbTriangle.Checked = False
            cbLabel.Checked = False
        End If
    End Sub

    Private Sub cbTriangle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTriangle.CheckedChanged
        If cbTriangle.Checked = True Then
            cbComp.Checked = False
            cbSquare.Checked = False
            cbPrism.Checked = False
            cbLine.Checked = False
            cbLabel.Checked = False
        End If
    End Sub

    Private Sub cbCylinder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbLabel.CheckedChanged
        If cbLabel.Checked = True Then
            cbComp.Checked = False
            cbSquare.Checked = False
            cbPrism.Checked = False
            cbLine.Checked = False
            cbTriangle.Checked = False
        End If
    End Sub
    Private Sub cbLockX_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbLockX.CheckedChanged
        If cbLockX.Checked = True Then
            cbLockY.Checked = False
            cbLockZ.Checked = False
        End If
    End Sub

    Private Sub cbLockY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbLockY.CheckedChanged
        If cbLockY.Checked = True Then
            cbLockX.Checked = False
            cbLockZ.Checked = False
        End If
    End Sub

    Private Sub cbLockZ_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbLockZ.CheckedChanged
        If cbLockZ.Checked = True Then
            cbLockY.Checked = False
            cbLockX.Checked = False
        End If
    End Sub
#End Region
#Region "Main Controls"
    Private Sub SaveImage()
        Dim SD As New SaveFileDialog
        SD.Title = "Save Image"
        SD.Filter = "JPEG Image|*.jpg"
        SD.InitialDirectory = fileImageData
        SD.RestoreDirectory = True
        SD.ShowDialog()
        If SD.FileName <> "" Then
            Try
                gImage.Save(SD.FileName, System.Drawing.Imaging.ImageFormat.Bmp)
            Catch ex As Exception
                If ex.Message.ToString <> "" Then
                    tbLog.AppendText(vbNewLine + "Image save error.")
                End If
                tbLog.AppendText(ex.Message.ToString)
            Finally
                If DialogResult = DialogResult.OK Then
                    tbLog.AppendText(vbNewLine + "Image save successful.")
                End If
                If DialogResult = DialogResult.Cancel Then
                    tbLog.AppendText(vbNewLine + "Image save cancelled.")
                End If
            End Try
        End If
    End Sub
    Private Sub SaveZone()
        If IsNothing(MacroData) = False Then
            Dim SD As New SaveFileDialog
            Dim TZ As New ZONE
            TZ.Z = MacroData
            SD.Filter = "Zone Data|*.dat"
            SD.Title = "Save Zone"
            SD.InitialDirectory = fileMacroData
            SD.RestoreDirectory = True
            SD.ShowDialog()
            If SD.FileName <> "" Then
                Try
                    SaveMacroData(SD.FileName, TZ)
                Catch ex As Exception
                    If ex.Message.ToString <> "" Then
                        tbLog.AppendText(vbNewLine + "Zone save error.")
                    End If
                    tbLog.AppendText(ex.Message.ToString)
                Finally
                    If DialogResult = DialogResult.OK Then
                        tbLog.AppendText(vbNewLine + "Zone save successful.")
                    End If
                    If DialogResult = DialogResult.Cancel Then
                        tbLog.AppendText(vbNewLine + "Zone save cancelled.")
                    End If
                End Try
            End If
        End If
    End Sub
    Private Sub LoadZone()
        Dim OD As New OpenFileDialog
        Dim TZ As New ZONE
        OD.Filter = "Zone Data|*.dat"
        OD.Title = "Load Zone"
        OD.InitialDirectory = fileMacroData
        OD.RestoreDirectory = True
        OD.ShowDialog()
        If OD.FileName <> "" Then
            Try
                TZ = LoadMacroData(OD.FileName)
                MacroData = TZ.Z
                valueVCZ = 0

                camLX = -(MainDisplay.Width / 2) : camRX = (MainDisplay.Width / 2) : camTY = (MainDisplay.Height / 2) : camBY = -(MainDisplay.Height / 2)
                mView = MI()
                wEye = vCompile(0, 0, camDist)
            Catch ex As Exception
                If ex.Message.ToString <> "" Then
                    tbLog.AppendText(vbNewLine + "Zone load error.")
                End If
                tbLog.AppendText(vbNewLine + ex.Message.ToString)
            Finally
                If DialogResult = DialogResult.OK Then
                    tbLog.AppendText(vbNewLine + "Zone load successful.")
                End If
                If DialogResult = DialogResult.Cancel Then
                    tbLog.AppendText(vbNewLine + "Zone load cancelled.")
                End If
                valueMacArray = 0
                valueMicArray = 0
                valueAT = 0
                valueAS = 0
                tbAdjustSelect.Text = valueAS.ToString
                tbAdjustTotal.Text = valueAT.ToString
                tbAdjustMacro.Text = valueMacArray.ToString
                tbAdjustMicro.Text = valueMicArray.ToString
                valueVAX = 0
                valueVAY = 0
                valueVAZ = 0
                tbAdjustVX.Text = valueVAX.ToString
                tbAdjustVY.Text = valueVAY.ToString
                tbAdjustVZ.Text = valueVAZ.ToString
            End Try
        End If
    End Sub
    Private Sub ExportModel()
        If IsNothing(MacroData) = False Then
            Dim SD As New SaveFileDialog
            Dim TM As New MODEL

            TM.M = cPress(MacroData)
            SD.Filter = "Model Data|*.dat"
            SD.Title = "Export Model"
            SD.InitialDirectory = fileMicroData
            SD.RestoreDirectory = True
            SD.ShowDialog()
            If SD.FileName <> "" Then
                Try
                    SaveMicroData(SD.FileName, TM)
                Catch ex As Exception
                    If ex.Message.ToString <> "" Then
                        tbLog.AppendText(vbNewLine + "Model save error.")
                    End If
                    tbLog.AppendText(ex.Message.ToString)
                Finally
                    M_DATA = IO.Directory.GetFiles(fileMicroData)
                    ModelData = M_Switch(M_DATA(modC))
                    If DialogResult = DialogResult.OK Then
                        tbLog.AppendText(vbNewLine + "Model save successful.")
                    End If
                    If DialogResult = DialogResult.Cancel Then
                        tbLog.AppendText(vbNewLine + "Model save cancelled.")
                    End If
                End Try
            End If
        End If
    End Sub
    Private Sub ImportModel()
        Try
            MicroData = ModelData
            ModelData = M_Switch(M_DATA(modC))
            modView = MI()
            mEye = vCompile(0, 0, camDist)
            modView = mXm(ROX(rad * 3), modView)
            mEye = norm(mXv((ROX(-rad * 3)), mEye))
        Catch ex As Exception

        Finally
            valueMacArray = MicroData.Length
            valueMicArray = 1
            valueAT = MicroData(0).Mesh.Length
            valueAS = 1
            tbAdjustSelect.Text = valueAS.ToString
            tbAdjustTotal.Text = valueAT.ToString
            tbAdjustMacro.Text = valueMacArray.ToString
            tbAdjustMicro.Text = valueMicArray.ToString
            valueVAX = MicroData(0).Mesh(0).X
            valueVAY = MicroData(0).Mesh(0).Y
            valueVAZ = MicroData(0).Mesh(0).Z
            tbAdjustVX.Text = valueVAX.ToString("N3")
            tbAdjustVY.Text = valueVAY.ToString("N3")
            tbAdjustVZ.Text = valueVAZ.ToString("N3")
            '---------------------------------------------'
            valueVPA = MicroData(0).LineColor.A
            valueVPR = MicroData(0).LineColor.R
            valueVPG = MicroData(0).LineColor.G
            valueVPB = MicroData(0).LineColor.B
            valueVPFA = MicroData(0).FillColor.A
            valueVPFR = MicroData(0).FillColor.R
            valueVPFG = MicroData(0).FillColor.G
            valueVPFB = MicroData(0).FillColor.B
            tbPaintVA.Text = valueVPA.ToString
            tbPaintVR.Text = valueVPR.ToString
            tbPaintVG.Text = valueVPG.ToString
            tbPaintVB.Text = valueVPB.ToString
            tbPaintVFA.Text = valueVPFA.ToString
            tbPaintVFR.Text = valueVPFR.ToString
            tbPaintVFG.Text = valueVPFG.ToString
            tbPaintVFB.Text = valueVPFB.ToString
        End Try
    End Sub
    Private Sub UndoLast()
        If IsNothing(MicroData) = False Then
            If MicroData.Length = 1 Then
                MicroData = Nothing
                valueMacArray = 0
                valueMicArray = 0
                valueAT = 0
                valueAS = 0
                tbAdjustSelect.Text = valueAS.ToString
                tbAdjustTotal.Text = valueAT.ToString
                tbAdjustMacro.Text = valueMacArray.ToString
                tbAdjustMicro.Text = valueMicArray.ToString
                valueVAX = 0
                valueVAY = 0
                valueVAZ = 0
                tbAdjustVX.Text = valueVAX.ToString("N3")
                tbAdjustVY.Text = valueVAY.ToString("N3")
                tbAdjustVZ.Text = valueVAZ.ToString("N3")
                tbLog.AppendText(vbNewLine + "Undo last set.")
            Else
                ReDim Preserve MicroData(MicroData.Length - 2)
                valueMacArray = 0
                valueMicArray = 0
                valueAT = 0
                valueAS = 0
                tbAdjustSelect.Text = valueAS.ToString
                tbAdjustTotal.Text = valueAT.ToString
                tbAdjustMacro.Text = valueMacArray.ToString
                tbAdjustMicro.Text = valueMicArray.ToString
                valueVAX = 0
                valueVAY = 0
                valueVAZ = 0
                tbAdjustVX.Text = valueVAX.ToString("N3")
                tbAdjustVY.Text = valueVAY.ToString("N3")
                tbAdjustVZ.Text = valueVAZ.ToString("N3")
                tbLog.AppendText(vbNewLine + "Undo last set.")
            End If
        End If
    End Sub
    Private Sub ResetZone()
        MacroData = Nothing
        MicroData = Nothing
        valueMacArray = 0
        valueMicArray = 0
        valueAT = 0
        valueAS = 0
        tbAdjustSelect.Text = valueAS.ToString
        tbAdjustTotal.Text = valueAT.ToString
        tbAdjustMacro.Text = valueMacArray.ToString
        tbAdjustMicro.Text = valueMicArray.ToString
        valueVAX = 0
        valueVAY = 0
        valueVAZ = 0
        tbAdjustVX.Text = valueVAX.ToString("N3")
        tbAdjustVY.Text = valueVAY.ToString("N3")
        tbAdjustVZ.Text = valueVAZ.ToString("N3")
        valueVCZ = 0
        ScrX = MainDisplay.Width
        ScrY = MainDisplay.Height
        If ScrX >= ScrY Then
            AR = ScrX / ScrY
        ElseIf ScrX < ScrY Then
            AR = ScrY / ScrX
        End If
        Int_Data()
        Int_Display()
        tbLog.AppendText(vbNewLine + "Zone cleared.")
    End Sub
    Private Sub UpdateAdjust()
        valueMacArray = MicroData.Length
        valueMicArray = 1
        valueAT = MicroData(0).Mesh.Length
        valueAS = 1
        tbAdjustSelect.Text = valueAS.ToString
        tbAdjustTotal.Text = valueAT.ToString
        tbAdjustMacro.Text = valueMacArray.ToString
        tbAdjustMicro.Text = valueMicArray.ToString
        valueVAX = MicroData(0).Mesh(0).X
        valueVAY = MicroData(0).Mesh(0).Y
        valueVAZ = MicroData(0).Mesh(0).Z
        tbAdjustVX.Text = valueVAX.ToString("N3")
        tbAdjustVY.Text = valueVAY.ToString("N3")
        tbAdjustVZ.Text = valueVAZ.ToString("N3")
        '---------------------------------------------'
        valueVPA = MicroData(0).LineColor.A
        valueVPR = MicroData(0).LineColor.R
        valueVPG = MicroData(0).LineColor.G
        valueVPB = MicroData(0).LineColor.B
        valueVPFA = MicroData(0).FillColor.A
        valueVPFR = MicroData(0).FillColor.R
        valueVPFG = MicroData(0).FillColor.G
        valueVPFB = MicroData(0).FillColor.B
        tbPaintVA.Text = valueVPA.ToString
        tbPaintVR.Text = valueVPR.ToString
        tbPaintVG.Text = valueVPG.ToString
        tbPaintVB.Text = valueVPB.ToString
        tbPaintVFA.Text = valueVPFA.ToString
        tbPaintVFR.Text = valueVPFR.ToString
        tbPaintVFG.Text = valueVPFG.ToString
        tbPaintVFB.Text = valueVPFB.ToString
    End Sub
    Private Sub SetVector()
        If cbLine.Checked = True Then
            StepMicro(toolLine(vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ), valueVRD, vCompile(valueVOX, valueVOY, valueVOZ)))
        End If
        If cbSquare.Checked = True Then
            StepMicro(toolSquare(valueVRD, vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))
        End If
        If cbTriangle.Checked = True Then
            StepMicro(toolTriangle(valueVRD, vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))
        End If
        If cbComp.Checked = True Then
            StepMicro(toolCompass(vCompile(valueVOX, valueVOY, valueVOZ), vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ)))
        End If
        If cbPrism.Checked = True Then
            For Each M As MDATA In toolCube(vCompile(valueVOX, valueVOY, valueVOZ), valueVMX, valueVMY, valueVMZ)
                StepMicro(M.Mesh)
            Next
        End If
        If cbLabel.Checked = True Then
            StepMicro(toolLine(vCompile(valueVMX, valueVMY, valueVMZ), vCompile(valueVPX, valueVPY, valueVPZ), valueVPZ, vCompile(valueVOX, valueVOY, valueVOZ)))
        End If
        tbLog.AppendText(vbNewLine + "Shape set.")
    End Sub
    Private Sub PressModel()
        If IsNothing(MicroData) = False Then
            PressMacro(MicroData)
            MicroData = Nothing
            tbLog.AppendText(vbNewLine + "Model pressed.")
        Else
            tbLog.AppendText(vbNewLine + "Micro buffer empty.")
        End If
    End Sub
#End Region
#Region "switchboard"
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveZone()
    End Sub

    Private Sub btnLoad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLoad.Click
        LoadZone()
    End Sub

    Private Sub btnUndo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUndo.Click
        UndoLast()
    End Sub

    Private Sub btnRedo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRedo.Click
        ResetZone()
    End Sub

    Private Sub btnStep_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnStep.Click
        SetVector()
    End Sub

    Private Sub btnPress_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPress.Click
        PressModel()
    End Sub

    Private Sub btnImport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnImport.Click
        ImportModel()
    End Sub

    Private Sub btnExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExport.Click
        ExportModel()
    End Sub
#End Region
#Region "File Handle"
    Private Sub SaveMicroData(ByVal f As String, ByVal d As Object)
        Dim fs As Stream = New FileStream(f, FileMode.Create)
        Dim bf As BinaryFormatter = New BinaryFormatter()
        bf.Serialize(fs, d)
        fs.Close()
    End Sub
    Private Function LoadMicroData(ByVal f As String) As MODEL
        Dim rc As MODEL
        Dim bf As New BinaryFormatter()
        Dim fs As New FileStream(f, IO.FileMode.Open)
        rc = DirectCast(bf.Deserialize(fs), MODEL)
        Return rc
    End Function
    Private Sub SaveMacroData(ByVal f As String, ByVal d As Object)
        Dim fs As Stream = New FileStream(f, FileMode.Create)
        Dim bf As BinaryFormatter = New BinaryFormatter()
        bf.Serialize(fs, d)
        fs.Close()
    End Sub
    Private Function LoadMacroData(ByVal f As String) As ZONE
        Dim rc As ZONE
        Dim bf As New BinaryFormatter()
        Dim fs As New FileStream(f, IO.FileMode.Open)
        rc = DirectCast(bf.Deserialize(fs), ZONE)
        Return rc
    End Function
#End Region
#Region "switch number down/up"
    Private Sub btnAdjustVleft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdjustVleft.Click
        If valueAT <> 0 Then
            valueAS = ValueC(valueAS, -1, 1, valueAT)
            tbAdjustSelect.Text = valueAS.ToString
            valueVAX = MicroData(valueMicArray - 1).Mesh(valueAS - 1).X
            valueVAY = MicroData(valueMicArray - 1).Mesh(valueAS - 1).Y
            valueVAZ = MicroData(valueMicArray - 1).Mesh(valueAS - 1).Z
            tbAdjustVX.Text = valueVAX.ToString("N3")
            tbAdjustVY.Text = valueVAY.ToString("N3")
            tbAdjustVZ.Text = valueVAZ.ToString("N3")
            valueVPA = MicroData(valueMicArray - 1).LineColor.A
            valueVPR = MicroData(valueMicArray - 1).LineColor.R
            valueVPG = MicroData(valueMicArray - 1).LineColor.G
            valueVPB = MicroData(valueMicArray - 1).LineColor.B
            valueVPFA = MicroData(valueMicArray - 1).FillColor.A
            valueVPFR = MicroData(valueMicArray - 1).FillColor.R
            valueVPFG = MicroData(valueMicArray - 1).FillColor.G
            valueVPFB = MicroData(valueMicArray - 1).FillColor.B
            tbPaintVA.Text = valueVPA.ToString
            tbPaintVR.Text = valueVPR.ToString
            tbPaintVG.Text = valueVPG.ToString
            tbPaintVB.Text = valueVPB.ToString
            tbPaintVFA.Text = valueVPFA.ToString
            tbPaintVFR.Text = valueVPFR.ToString
            tbPaintVFG.Text = valueVPFG.ToString
            tbPaintVFB.Text = valueVPFB.ToString
        End If
    End Sub

    Private Sub btnAdjustVright_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdjustVright.Click
        If valueAT <> 0 Then
            valueAS = ValueC(valueAS, 1, 1, valueAT)
            tbAdjustSelect.Text = valueAS.ToString
            valueVAX = MicroData(valueMicArray - 1).Mesh(valueAS - 1).X
            valueVAY = MicroData(valueMicArray - 1).Mesh(valueAS - 1).Y
            valueVAZ = MicroData(valueMicArray - 1).Mesh(valueAS - 1).Z
            tbAdjustVX.Text = valueVAX.ToString("N3")
            tbAdjustVY.Text = valueVAY.ToString("N3")
            tbAdjustVZ.Text = valueVAZ.ToString("N3")
            valueVPA = MicroData(valueMicArray - 1).LineColor.A
            valueVPR = MicroData(valueMicArray - 1).LineColor.R
            valueVPG = MicroData(valueMicArray - 1).LineColor.G
            valueVPB = MicroData(valueMicArray - 1).LineColor.B
            valueVPFA = MicroData(valueMicArray - 1).FillColor.A
            valueVPFR = MicroData(valueMicArray - 1).FillColor.R
            valueVPFG = MicroData(valueMicArray - 1).FillColor.G
            valueVPFB = MicroData(valueMicArray - 1).FillColor.B
            tbPaintVA.Text = valueVPA.ToString
            tbPaintVR.Text = valueVPR.ToString
            tbPaintVG.Text = valueVPG.ToString
            tbPaintVB.Text = valueVPB.ToString
            tbPaintVFA.Text = valueVPFA.ToString
            tbPaintVFR.Text = valueVPFR.ToString
            tbPaintVFG.Text = valueVPFG.ToString
            tbPaintVFB.Text = valueVPFB.ToString
        End If
    End Sub
    Private Sub btnAdjustARight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdjustARight.Click
        If valueMacArray <> 0 Then
            valueMicArray = ValueC(valueMicArray, 1, 1, valueMacArray)
            tbAdjustMicro.Text = valueMicArray.ToString
            valueAT = MicroData(valueMicArray - 1).Mesh.Length
            valueAS = 1
            tbAdjustSelect.Text = valueAS.ToString
            tbAdjustTotal.Text = valueAT.ToString

            valueVAX = MicroData(valueMicArray - 1).Mesh(0).X
            valueVAY = MicroData(valueMicArray - 1).Mesh(0).Y
            valueVAZ = MicroData(valueMicArray - 1).Mesh(0).Z
            tbAdjustVX.Text = valueVAX.ToString("N3")
            tbAdjustVY.Text = valueVAY.ToString("N3")
            tbAdjustVZ.Text = valueVAZ.ToString("N3")
            valueVPA = MicroData(valueMicArray - 1).LineColor.A
            valueVPR = MicroData(valueMicArray - 1).LineColor.R
            valueVPG = MicroData(valueMicArray - 1).LineColor.G
            valueVPB = MicroData(valueMicArray - 1).LineColor.B
            valueVPFA = MicroData(valueMicArray - 1).FillColor.A
            valueVPFR = MicroData(valueMicArray - 1).FillColor.R
            valueVPFG = MicroData(valueMicArray - 1).FillColor.G
            valueVPFB = MicroData(valueMicArray - 1).FillColor.B
            tbPaintVA.Text = valueVPA.ToString
            tbPaintVR.Text = valueVPR.ToString
            tbPaintVG.Text = valueVPG.ToString
            tbPaintVB.Text = valueVPB.ToString
            tbPaintVFA.Text = valueVPFA.ToString
            tbPaintVFR.Text = valueVPFR.ToString
            tbPaintVFG.Text = valueVPFG.ToString
            tbPaintVFB.Text = valueVPFB.ToString
        End If
    End Sub

    Private Sub btnAdjustALeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdjustALeft.Click
        If valueMacArray <> 0 Then
            valueMicArray = ValueC(valueMicArray, -1, 1, valueMacArray)
            tbAdjustMicro.Text = valueMicArray.ToString
            valueAT = MicroData(0).Mesh.Length
            valueAS = 1
            tbAdjustSelect.Text = valueAS.ToString
            tbAdjustTotal.Text = valueAT.ToString

            valueVAX = MicroData(valueMicArray - 1).Mesh(0).X
            valueVAY = MicroData(valueMicArray - 1).Mesh(0).Y
            valueVAZ = MicroData(valueMicArray - 1).Mesh(0).Z
            tbAdjustVX.Text = valueVAX.ToString("N3")
            tbAdjustVY.Text = valueVAY.ToString("N3")
            tbAdjustVZ.Text = valueVAZ.ToString("N3")
            valueVPA = MicroData(valueMicArray - 1).LineColor.A
            valueVPR = MicroData(valueMicArray - 1).LineColor.R
            valueVPG = MicroData(valueMicArray - 1).LineColor.G
            valueVPB = MicroData(valueMicArray - 1).LineColor.B
            valueVPFA = MicroData(valueMicArray - 1).FillColor.A
            valueVPFR = MicroData(valueMicArray - 1).FillColor.R
            valueVPFG = MicroData(valueMicArray - 1).FillColor.G
            valueVPFB = MicroData(valueMicArray - 1).FillColor.B
            tbPaintVA.Text = valueVPA.ToString
            tbPaintVR.Text = valueVPR.ToString
            tbPaintVG.Text = valueVPG.ToString
            tbPaintVB.Text = valueVPB.ToString
            tbPaintVFA.Text = valueVPFA.ToString
            tbPaintVFR.Text = valueVPFR.ToString
            tbPaintVFG.Text = valueVPFG.ToString
            tbPaintVFB.Text = valueVPFB.ToString
        End If
    End Sub
    Private Sub btnAdjustXdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdjustXdown.Click
        valueVAX = ValueC(valueVAX, -ValueMark, -99999, 99999)
        If valueAS <> 0 Then
            MicroData(valueMicArray - 1).Mesh(valueAS - 1).X = valueVAX
        End If
        tbAdjustVX.Text = valueVAX.ToString("N3")
    End Sub

    Private Sub btnAdjustXup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdjustXup.Click
        valueVAX = ValueC(valueVAX, ValueMark, -99999, 99999)
        If valueAS <> 0 Then
            MicroData(valueMicArray - 1).Mesh(valueAS - 1).X = valueVAX
        End If
        tbAdjustVX.Text = valueVAX.ToString("N3")
    End Sub

    Private Sub btnAdjustYdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdjustYdown.Click
        valueVAY = ValueC(valueVAY, -ValueMark, -99999, 99999)
        If valueAS <> 0 Then
            MicroData(valueMicArray - 1).Mesh(valueAS - 1).Y = valueVAY
        End If
        tbAdjustVY.Text = valueVAY.ToString("N3")
    End Sub

    Private Sub btnAdjustYup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdjustYup.Click
        valueVAY = ValueC(valueVAY, ValueMark, -99999, 99999)
        If valueAS <> 0 Then
            MicroData(valueMicArray - 1).Mesh(valueAS - 1).Y = valueVAY
        End If
        tbAdjustVY.Text = valueVAY.ToString("N3")
    End Sub

    Private Sub btnAdjustZdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdjustZdown.Click
        valueVAZ = ValueC(valueVAZ, -ValueMark, -99999, 99999)
        If valueAS <> 0 Then
            MicroData(valueMicArray - 1).Mesh(valueAS - 1).Z = valueVAZ
        End If
        tbAdjustVZ.Text = valueVAZ.ToString("N3")
    End Sub

    Private Sub btnAdjustZup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdjustZup.Click
        valueVAZ = ValueC(valueVAZ, ValueMark, -99999, 99999)
        If valueAS <> 0 Then
            MicroData(valueMicArray - 1).Mesh(valueAS - 1).Z = valueVAZ
        End If
        tbAdjustVZ.Text = valueVAZ.ToString("N3")
    End Sub

    Private Sub btnMeasureXdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMeasureXdown.Click
        valueVMX = ValueC(valueVMX, -ValueMark, -99999, 99999)
        tbMeasureVX.Text = valueVMX.ToString("N3")
    End Sub

    Private Sub btnMeasureXup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMeasureXup.Click
        valueVMX = ValueC(valueVMX, ValueMark, -99999, 99999)
        tbMeasureVX.Text = valueVMX.ToString("N3")
    End Sub

    Private Sub btnMeasureYdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMeasureYdown.Click
        valueVMY = ValueC(valueVMY, -ValueMark, -99999, 99999)
        tbMeasureVY.Text = valueVMY.ToString("N3")
    End Sub

    Private Sub btnMeasureYup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMeasureYup.Click
        valueVMY = ValueC(valueVMY, ValueMark, -99999, 99999)
        tbMeasureVY.Text = valueVMY.ToString("N3")
    End Sub

    Private Sub btnMeasureZdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMeasureZdown.Click
        valueVMZ = ValueC(valueVMZ, -ValueMark, -99999, 99999)
        tbMeasureVZ.Text = valueVMZ.ToString("N3")
    End Sub

    Private Sub btnMeasureZup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMeasureZup.Click
        valueVMZ = ValueC(valueVMZ, ValueMark, -99999, 99999)
        tbMeasureVZ.Text = valueVMZ.ToString("N3")
    End Sub

    Private Sub btnOriginXdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOriginXdown.Click
        valueVOX = ValueC(valueVOX, -ValueMark, -99999, 99999)
        tbOriginVX.Text = valueVOX.ToString("N3")
    End Sub

    Private Sub btnOriginXup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOriginXup.Click
        valueVOX = ValueC(valueVOX, ValueMark, -99999, 99999)
        tbOriginVX.Text = valueVOX.ToString("N3")
    End Sub

    Private Sub btnOriginYdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOriginYdown.Click
        valueVOY = ValueC(valueVOY, -ValueMark, -99999, 99999)
        tbOriginVY.Text = valueVOY.ToString("N3")
    End Sub

    Private Sub btnOriginYup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOriginYup.Click
        valueVOY = ValueC(valueVOY, ValueMark, -99999, 99999)
        tbOriginVY.Text = valueVOY.ToString("N3")
    End Sub

    Private Sub btnOriginZdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOriginZdown.Click
        valueVOZ = ValueC(valueVOZ, -ValueMark, -99999, 99999)
        tbOriginVZ.Text = valueVOZ.ToString("N3")
    End Sub

    Private Sub btnOriginZup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOriginZup.Click
        valueVOZ = ValueC(valueVOZ, ValueMark, -99999, 99999)
        tbOriginVZ.Text = valueVOZ.ToString("N3")
    End Sub
    '-----------------------------------------------------------
    Private Sub btnPaintAdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintAdown.Click
        valueVPA = ValueC(valueVPA, -5, 0, 255)
        tbPaintVA.Text = valueVPA.ToString
    End Sub

    Private Sub btnPaintAup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintAup.Click
        valueVPA = ValueC(valueVPA, 5, 0, 255)
        tbPaintVA.Text = valueVPA.ToString
    End Sub

    Private Sub btnPaintBdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintBdown.Click
        valueVPB = ValueC(valueVPB, -5, 0, 255)
        tbPaintVB.Text = valueVPB.ToString
    End Sub

    Private Sub btnPaintBup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintBup.Click
        valueVPB = ValueC(valueVPB, 5, 0, 255)
        tbPaintVB.Text = valueVPB.ToString
    End Sub

    Private Sub btnPaintGdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintGdown.Click
        valueVPG = ValueC(valueVPG, -5, 0, 255)
        tbPaintVG.Text = valueVPG.ToString
    End Sub

    Private Sub btnPaintGup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintGup.Click
        valueVPG = ValueC(valueVPG, 5, 0, 255)
        tbPaintVG.Text = valueVPG.ToString
    End Sub

    Private Sub btnPaintRdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintRdown.Click
        valueVPR = ValueC(valueVPR, -5, 0, 255)
        tbPaintVR.Text = valueVPR.ToString
    End Sub

    Private Sub btnPaintRup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintRup.Click
        valueVPR = ValueC(valueVPR, 5, 0, 255)
        tbPaintVR.Text = valueVPR.ToString
    End Sub
    '-----------------------------------------------------------
    Private Sub btnPaintFAdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintFAdown.Click
        valueVPFA = ValueC(valueVPFA, -5, 0, 255)
        tbPaintVFA.Text = valueVPFA.ToString
    End Sub

    Private Sub btnPaintFAup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintFAup.Click
        valueVPFA = ValueC(valueVPFA, 5, 0, 255)
        tbPaintVFA.Text = valueVPFA.ToString
    End Sub

    Private Sub btnPaintFBdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintFBdown.Click
        valueVPFB = ValueC(valueVPFB, -5, 0, 255)
        tbPaintVFB.Text = valueVPFB.ToString
    End Sub

    Private Sub btnPaintFBup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintFBup.Click
        valueVPFB = ValueC(valueVPFB, 5, 0, 255)
        tbPaintVFB.Text = valueVPFB.ToString
    End Sub

    Private Sub btnPaintFGdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintFGdown.Click
        valueVPFG = ValueC(valueVPFG, -5, 0, 255)
        tbPaintVFG.Text = valueVPFG.ToString
    End Sub

    Private Sub btnPaintFGup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintFGup.Click
        valueVPFG = ValueC(valueVPFG, 5, 0, 255)
        tbPaintVFG.Text = valueVPFG.ToString
    End Sub

    Private Sub btnPaintFRdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintFRdown.Click
        valueVPFR = ValueC(valueVPFR, -5, 0, 255)
        tbPaintVFR.Text = valueVPFR.ToString
    End Sub

    Private Sub btnPaintFRup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPaintFRup.Click
        valueVPFR = ValueC(valueVPFR, 5, 0, 255)
        tbPaintVFR.Text = valueVPFR.ToString
    End Sub
    '-----------------------------------------------------------
    Private Sub btnColorSwitchA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorSwitchA.Click
        If valueVPA <= 122 Then
            valueVPA = 255
        ElseIf valueVPA >= 123 Then
            valueVPA = 0
        End If
        tbPaintVA.Text = valueVPA.ToString
    End Sub

    Private Sub btnColorSwitchR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorSwitchR.Click
        If valueVPR <= 122 Then
            valueVPR = 255
        ElseIf valueVPR >= 123 Then
            valueVPR = 0
        End If
        tbPaintVR.Text = valueVPR.ToString
    End Sub

    Private Sub btnColorSwitchG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorSwitchG.Click
        If valueVPG <= 122 Then
            valueVPG = 255
        ElseIf valueVPG >= 123 Then
            valueVPG = 0
        End If
        tbPaintVG.Text = valueVPG.ToString
    End Sub

    Private Sub btnColorSwitchB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorSwitchB.Click
        If valueVPB <= 122 Then
            valueVPB = 255
        ElseIf valueVPB >= 123 Then
            valueVPB = 0
        End If
        tbPaintVB.Text = valueVPB.ToString
    End Sub
    '-----------------------------------------------------------
    Private Sub btnColorSwitchFA_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorSwitchFA.Click
        If valueVPFA <= 122 Then
            valueVPFA = 255
        ElseIf valueVPFA >= 123 Then
            valueVPFA = 0
        End If
        tbPaintVFA.Text = valueVPFA.ToString
    End Sub

    Private Sub btnColorSwitchFR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorSwitchFR.Click
        If valueVPFR <= 122 Then
            valueVPFR = 255
        ElseIf valueVPFR >= 123 Then
            valueVPFR = 0
        End If
        tbPaintVFR.Text = valueVPFR.ToString
    End Sub

    Private Sub btnColorSwitchFG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorSwitchFG.Click
        If valueVPFG <= 122 Then
            valueVPFG = 255
        ElseIf valueVPFG >= 123 Then
            valueVPFG = 0
        End If
        tbPaintVFG.Text = valueVPFG.ToString
    End Sub

    Private Sub btnColorSwitchFB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorSwitchFB.Click
        If valueVPFB <= 122 Then
            valueVPFB = 255
        ElseIf valueVPFB >= 123 Then
            valueVPFB = 0
        End If
        tbPaintVFB.Text = valueVPFB.ToString
    End Sub
    Private Sub btnRadiusRdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRadiusRdown.Click
        valueVRD = ValueC(valueVRD, -ValueMark, -99999, 99999)
        tbRadiusV.Text = valueVRD.ToString("N3")
    End Sub

    Private Sub btnRadiusRup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRadiusRup.Click
        valueVRD = ValueC(valueVRD, ValueMark, -99999, 99999)
        tbRadiusV.Text = valueVRD.ToString("N3")
    End Sub

    Private Sub btnSpinXdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSpinXdown.Click
        valueVSX = ValueC(valueVSX, -ValueMark, -99999, 99999)
        tbSpinVX.Text = valueVSX.ToString("N3")
    End Sub

    Private Sub btnSpinXup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSpinXup.Click
        valueVSX = ValueC(valueVSX, ValueMark, -99999, 99999)
        tbSpinVX.Text = valueVSX.ToString("N3")
    End Sub

    Private Sub btnSpinYdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSpinYdown.Click
        valueVSY = ValueC(valueVSY, -ValueMark, -99999, 99999)
        tbSpinVY.Text = valueVSY.ToString("N3")
    End Sub

    Private Sub btnSpinYup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSpinYup.Click
        valueVSY = ValueC(valueVSY, ValueMark, -99999, 99999)
        tbSpinVY.Text = valueVSY.ToString("N3")
    End Sub

    Private Sub btnSpinZdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSpinZdown.Click
        valueVSZ = ValueC(valueVSZ, -ValueMark, -99999, 99999)
        tbSpinVZ.Text = valueVSZ.ToString("N3")
    End Sub

    Private Sub btnSpinZup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSpinZup.Click
        valueVSZ = ValueC(valueVSZ, ValueMark, -99999, 99999)
        tbSpinVZ.Text = valueVSZ.ToString("N3")
    End Sub

    Private Sub btnTranslateXdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTranslateXdown.Click
        valueVTX = ValueC(valueVTX, -ValueMark, -99999, 99999)
        tbTranslateVX.Text = valueVTX.ToString("N3")
    End Sub

    Private Sub btnTranslateXup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTranslateXup.Click
        valueVTX = ValueC(valueVTX, ValueMark, -99999, 99999)
        tbTranslateVX.Text = valueVTX.ToString("N3")
    End Sub

    Private Sub btnTranslateYdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTranslateYdown.Click
        valueVTY = ValueC(valueVTY, -ValueMark, -99999, 99999)
        tbTranslateVY.Text = valueVTY.ToString("N3")
    End Sub

    Private Sub btnTranslateYup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTranslateYup.Click
        valueVTY = ValueC(valueVTY, ValueMark, -99999, 99999)
        tbTranslateVY.Text = valueVTY.ToString("N3")
    End Sub

    Private Sub btnTranslateZdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTranslateZdown.Click
        valueVTZ = ValueC(valueVTZ, -ValueMark, -99999, 99999)
        tbTranslateVZ.Text = valueVTZ.ToString("N3")
    End Sub

    Private Sub btnTranslateZup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTranslateZup.Click
        valueVTZ = ValueC(valueVTZ, ValueMark, -99999, 99999)
        tbTranslateVZ.Text = valueVTZ.ToString("N3")
    End Sub

    Private Sub btnMicrodown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMicrodown.Click
        valueMicro = ValueC(valueMicro, -0.01, -99999, 99999)
        tbMicroV.Text = valueMicro.ToString("N3")
        If cbMicro.Checked = True Then ValueMark = valueMicro
    End Sub

    Private Sub btnMicroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMicroup.Click
        valueMicro = ValueC(valueMicro, 0.01, -99999, 99999)
        tbMicroV.Text = valueMicro.ToString("N3")
        If cbMicro.Checked = True Then ValueMark = valueMicro
    End Sub

    Private Sub btnMacrodown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMacrodown.Click
        valueMacro = ValueC(valueMacro, -1, -99999, 99999)
        tbMacroV.Text = valueMacro.ToString("N3")
        If cbMacro.Checked = True Then ValueMark = valueMacro
    End Sub
    Private Sub btnMacroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMacroup.Click
        valueMacro = ValueC(valueMacro, 1, -99999, 99999)
        tbMacroV.Text = valueMacro.ToString("N3")
        If cbMacro.Checked = True Then ValueMark = valueMacro
    End Sub
    Private Sub btnPeakXdown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPeakXdown.Click
        valueVPX = ValueC(valueVPX, -ValueMark, -99999, 99999)
        tbPeakVX.Text = valueVPX.ToString("N3")
    End Sub

    Private Sub btnPeakXup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPeakXup.Click
        valueVPX = ValueC(valueVPX, ValueMark, -99999, 99999)
        tbPeakVX.Text = valueVPX.ToString("N3")
    End Sub

    Private Sub btnPeakYdown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPeakYdown.Click
        valueVPY = ValueC(valueVPY, -ValueMark, -99999, 99999)
        tbPeakVY.Text = valueVPY.ToString("N3")
    End Sub

    Private Sub btnPeakYup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPeakYup.Click
        valueVPY = ValueC(valueVPY, ValueMark, -99999, 99999)
        tbPeakVY.Text = valueVPY.ToString("N3")
    End Sub

    Private Sub btnPeakZdown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPeakZdown.Click
        valueVPZ = ValueC(valueVPZ, -ValueMark, -99999, 99999)
        tbPeakVZ.Text = valueVPZ.ToString("N3")
    End Sub

    Private Sub btnPeakZup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPeakZup.Click
        valueVPZ = ValueC(valueVPZ, ValueMark, -99999, 99999)
        tbPeakVZ.Text = valueVPZ.ToString("N3")
    End Sub

    Private Sub btnGridSizeDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGridSizeDown.Click
        valueVGS = ValueC(valueVGS, -5, 10, 1000)
        tbGridSize.Text = valueVGS.ToString
        GridDataXZ = drawGridXZ(valueVGS, valueVGD)
        GridDataXY = drawGridXY(valueVGS, valueVGD)
        GridDataYZ = drawGridYZ(valueVGS, valueVGD)
    End Sub

    Private Sub btnGridSizeUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGridSizeUp.Click
        valueVGS = ValueC(valueVGS, 5, 10, 1000)
        tbGridSize.Text = valueVGS.ToString
        GridDataXZ = drawGridXZ(valueVGS, valueVGD)
        GridDataXY = drawGridXY(valueVGS, valueVGD)
        GridDataYZ = drawGridYZ(valueVGS, valueVGD)
    End Sub

    Private Sub btnGridDimDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGridDimDown.Click
        valueVGD = ValueC(valueVGD, -1, 1, 1000)
        tbGridDim.Text = valueVGD.ToString
        GridDataXZ = drawGridXZ(valueVGS, valueVGD)
        GridDataXY = drawGridXY(valueVGS, valueVGD)
        GridDataYZ = drawGridYZ(valueVGS, valueVGD)
    End Sub

    Private Sub btnGridDimUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGridDimUp.Click
        valueVGD = ValueC(valueVGD, 1, 1, 1000)
        tbGridDim.Text = valueVGD.ToString
        GridDataXZ = drawGridXZ(valueVGS, valueVGD)
        GridDataXY = drawGridXY(valueVGS, valueVGD)
        GridDataYZ = drawGridYZ(valueVGS, valueVGD)
    End Sub

#End Region
#Region "boolean I/O"
    Private Sub btnViewup_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnViewup.MouseDown
        bViewUp = True
    End Sub

    Private Sub btnViewup_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnViewup.MouseUp
        bViewUp = False
    End Sub

    Private Sub btnViewdown_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnViewdown.MouseDown
        bViewDown = True
    End Sub

    Private Sub btnViewdown_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnViewdown.MouseUp
        bViewDown = False
    End Sub

    Private Sub btnViewleft_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnViewleft.MouseDown
        bViewLeft = True
    End Sub

    Private Sub btnViewleft_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnViewleft.MouseUp
        bViewLeft = False
    End Sub

    Private Sub btnViewright_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnViewright.MouseDown
        bViewRight = True
    End Sub

    Private Sub btnViewright_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnViewright.MouseUp
        bViewRight = False
    End Sub
    '---------------------------------------------------------'
    Private Sub btnZoomIn_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnZoomIn.MouseDown
        bZoomIn = True
    End Sub

    Private Sub btnZoomIn_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnZoomIn.MouseUp
        bZoomIn = False
    End Sub

    Private Sub btnZoomOut_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnZoomOut.MouseDown
        bZoomOut = True
    End Sub

    Private Sub btnZoomOut_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnZoomOut.MouseUp
        bZoomOut = False
    End Sub

    Private Sub btnCamShiftLeft_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnCamShiftLeft.MouseDown
        bShiftL = True
    End Sub

    Private Sub btnCamShiftLeft_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnCamShiftLeft.MouseUp
        bShiftL = False
    End Sub

    Private Sub btnCamShiftRight_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnCamShiftRight.MouseDown
        bShiftR = True
    End Sub

    Private Sub btnCamShiftRight_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnCamShiftRight.MouseUp
        bShiftR = False
    End Sub

    Private Sub btnCamShiftUp_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnCamShiftUp.MouseDown
        bShiftU = True
    End Sub

    Private Sub btnCamShiftUp_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnCamShiftUp.MouseUp
        bShiftU = False
    End Sub

    Private Sub btnCamShiftDown_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnCamShiftDown.MouseDown
        bShiftD = True
    End Sub

    Private Sub btnCamShiftDown_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnCamShiftDown.MouseUp
        bShiftD = False
    End Sub
#End Region
#Region "motion control"
    '....................................................................'
    Private Sub btnGhostXleft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGhostXleft.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = rPress(vCompile(_Rad15, 0, 0), (MicroData(z).Mesh), MicroData(z).Center)
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostXright_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGhostXright.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = rPress(vCompile(-_Rad15, 0, 0), (MicroData(z).Mesh), MicroData(z).Center)
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostYleft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGhostYleft.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = rPress(vCompile(0, _Rad15, 0), (MicroData(z).Mesh), MicroData(z).Center)
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostYright_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGhostYright.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = rPress(vCompile(0, -_Rad15, 0), (MicroData(z).Mesh), MicroData(z).Center)
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostZleft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGhostZleft.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = rPress(vCompile(0, 0, -_Rad15), (MicroData(z).Mesh), MicroData(z).Center)
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostZright_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGhostZright.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = rPress(vCompile(0, 0, _Rad15), (MicroData(z).Mesh), MicroData(z).Center)
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostXdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGhostXdown.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = tPress(vCompile(-ValueMark, 0, 0), (MicroData(z).Mesh))
                MicroData(z).Origin = mXv(MTR(vCompile(-ValueMark, 0, 0)), norm(MicroData(z).Origin))
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostXup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGhostXup.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = tPress(vCompile(ValueMark, 0, 0), (MicroData(z).Mesh))
                MicroData(z).Origin = mXv(MTR(vCompile(ValueMark, 0, 0)), norm(MicroData(z).Origin))
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostYdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGhostYdown.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = tPress(vCompile(0, -ValueMark, 0), (MicroData(z).Mesh))
                MicroData(z).Origin = mXv(MTR(vCompile(0, -ValueMark, 0)), norm(MicroData(z).Origin))
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostYup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGhostYup.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = tPress(vCompile(0, ValueMark, 0), (MicroData(z).Mesh))
                MicroData(z).Origin = mXv(MTR(vCompile(0, ValueMark, 0)), norm(MicroData(z).Origin))
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostZdown_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGhostZdown.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = tPress(vCompile(0, 0, -ValueMark), (MicroData(z).Mesh))
                MicroData(z).Origin = mXv(MTR(vCompile(0, 0, -ValueMark)), norm(MicroData(z).Origin))
            Next z
            UpdateAdjust()
        End If
    End Sub

    Private Sub btnGhostZup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGhostZup.Click
        If IsNothing(MicroData) = False Then
            For z As Integer = 0 To MicroData.Length - 1
                MicroData(z).Mesh = tPress(vCompile(0, 0, ValueMark), (MicroData(z).Mesh))
                MicroData(z).Origin = mXv(MTR(vCompile(0, 0, ValueMark)), norm(MicroData(z).Origin))
            Next z
            UpdateAdjust()
        End If
    End Sub
    '....................................................................'
#End Region
#Region "Paint Styles"
    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        Dim g As Graphics = e.Graphics
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        Dim b1 As Pen = New Pen(cB1, 0)
        Dim b2 As Pen = New Pen(cB2, 0)
        Dim b3 As Pen = New Pen(cB3, 0)
        Dim b4 As Pen = New Pen(cB4, 0)
        g.DrawRectangle(b4, MainDisplay.Left - 1, MainDisplay.Top - 1, MainDisplay.Width + 1, MainDisplay.Height + 1)
        g.DrawRectangle(b3, MainDisplay.Left - 2, MainDisplay.Top - 2, MainDisplay.Width + 3, MainDisplay.Height + 3)
        g.DrawRectangle(b2, MainDisplay.Left - 3, MainDisplay.Top - 3, MainDisplay.Width + 5, MainDisplay.Height + 5)
        g.DrawRectangle(b1, MainDisplay.Left - 4, MainDisplay.Top - 4, MainDisplay.Width + 7, MainDisplay.Height + 7)
        b1.Dispose()
        b2.Dispose()
        b3.Dispose()
        b4.Dispose()
        g.Dispose()
    End Sub
    Private Sub pControl_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pControl.Paint
        Dim g As Graphics = e.Graphics
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        Dim b1 As Pen = New Pen(cB1, 0)
        Dim b2 As Pen = New Pen(cB2, 0)
        Dim b3 As Pen = New Pen(cB3, 0)
        Dim b4 As Pen = New Pen(cB4, 0)
        g.DrawRectangle(b1, 0, 0, pControl.Width - 1, pControl.Height - 1)
        g.DrawRectangle(b2, 1, 1, pControl.Width - 3, pControl.Height - 3)
        g.DrawRectangle(b3, 2, 2, pControl.Width - 5, pControl.Height - 5)
        g.DrawRectangle(b4, 3, 3, pControl.Width - 7, pControl.Height - 7)
        b1.Dispose()
        b2.Dispose()
        b3.Dispose()
        b4.Dispose()
        g.Dispose()
    End Sub
    Private Sub pViewControl_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pViewControl.Paint
        Dim g As Graphics = e.Graphics
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias
        Dim b1 As Pen = New Pen(cB1, 0)
        Dim b2 As Pen = New Pen(cB2, 0)
        Dim b3 As Pen = New Pen(cB3, 0)
        Dim b4 As Pen = New Pen(cB4, 0)
        g.DrawRectangle(b1, tbLog.Left - 1, tbLog.Top - 1, tbLog.Width + 1, tbLog.Height + 1)
        g.DrawRectangle(b2, tbLog.Left - 2, tbLog.Top - 2, tbLog.Width + 3, tbLog.Height + 3)
        g.DrawRectangle(b3, tbLog.Left - 3, tbLog.Top - 3, tbLog.Width + 5, tbLog.Height + 5)
        g.DrawRectangle(b4, tbLog.Left - 4, tbLog.Top - 4, tbLog.Width + 7, tbLog.Height + 7)
        g.DrawRectangle(b1, tbLabel.Left - 1, tbLabel.Top - 1, tbLabel.Width + 1, tbLabel.Height + 1)
        g.DrawRectangle(b2, tbLabel.Left - 2, tbLabel.Top - 2, tbLabel.Width + 3, tbLabel.Height + 3)
        g.DrawRectangle(b3, tbLabel.Left - 3, tbLabel.Top - 3, tbLabel.Width + 5, tbLabel.Height + 5)
        g.DrawRectangle(b4, tbLabel.Left - 4, tbLabel.Top - 4, tbLabel.Width + 7, tbLabel.Height + 7)
        g.DrawRectangle(b1, 0, 0, pViewControl.Width - 1, pViewControl.Height - 1)
        g.DrawRectangle(b2, 1, 1, pViewControl.Width - 3, pViewControl.Height - 3)
        g.DrawRectangle(b3, 2, 2, pViewControl.Width - 5, pViewControl.Height - 5)
        g.DrawRectangle(b4, 3, 3, pViewControl.Width - 7, pViewControl.Height - 7)
        b1.Dispose()
        b2.Dispose()
        b3.Dispose()
        b4.Dispose()
        g.Dispose()
    End Sub
#End Region
#Region "model control"

    Private Function M_Switch(ByVal f As String) As MDATA()
        Dim M As MDATA() = Nothing
        Dim SA() As String
        SA = M_DATA(modC).Split("\")
        M = LoadMicroData(f).M
        tbMString.Text = SA(SA.Length - 1).ToString
        Return M
    End Function
    Private Sub btnModelRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModelRight.Click
        If IsNothing(M_DATA) = False Then
            Try
                modC += 1
                If modC > M_DATA.Length - 1 Then
                    modC = 0
                End If
                ModelData = M_Switch(M_DATA(modC))
                modView = MI()
                mEye = vCompile(0, 0, camDist)
                modView = mXm(ROX(_Rad15), modView)
                mEye = norm(mXv((ROX(-_Rad15)), mEye))
            Catch ex As Exception
                If ex.Message.ToString <> "" Then
                    tbLog.AppendText(vbNewLine + "Model load error.")
                End If
                tbLog.AppendText(vbNewLine + ex.Message.ToString)
            Finally
                If DialogResult = DialogResult.OK Then
                    tbLog.AppendText(vbNewLine + "Model load successful.")
                End If
                If DialogResult = DialogResult.Cancel Then
                    tbLog.AppendText(vbNewLine + "Model load cancelled.")
                End If
            End Try
        End If
    End Sub
    Private Sub btnModelLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModelLeft.Click
        If IsNothing(M_DATA) = False Then
            Try
                modC -= 1
                If modC < 0 Then
                    modC = M_DATA.Length - 1
                End If
                ModelData = M_Switch(M_DATA(modC))
                modView = MI()
                mEye = vCompile(0, 0, camDist)
                modView = mXm(ROX(_Rad15), modView)
                mEye = norm(mXv((ROX(-_Rad15)), mEye))
            Catch ex As Exception
                If ex.Message.ToString <> "" Then
                    tbLog.AppendText(vbNewLine + "Model load error.")
                End If
                tbLog.AppendText(vbNewLine + ex.Message.ToString)
            Finally
                If DialogResult = DialogResult.OK Then
                    tbLog.AppendText(vbNewLine + "Model load successful.")
                End If
                If DialogResult = DialogResult.Cancel Then
                    tbLog.AppendText(vbNewLine + "Model load cancelled.")
                End If
            End Try
        End If
    End Sub
#End Region
#Region "Menu Switchboard"
    '---------------------------------------------------'
    Private Sub BackColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackColorToolStripMenuItem.Click
        Dim bc As ColorDialog = New ColorDialog
        bc.ShowDialog()
        bcC = bc.Color
        bc.Dispose()
    End Sub

    Private Sub ForeColorToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ForeColorToolStripMenuItem.Click
        Dim fc As ColorDialog = New ColorDialog
        fc.ShowDialog()
        fcC = fc.Color
        fc.Dispose()
    End Sub
    '---------------------------------------------------'
    Private Sub XZTileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles XZTileToolStripMenuItem.Click
        PressMacro(tBuildXZ(1000, 200, True, True, cColor(100, 0, 150, 0), cColor(100, 0, 150, 0)))
    End Sub

    Private Sub XYTileToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles XYTileToolStripMenuItem.Click
        PressMacro(tBuildXY(1000, 200, True, True, cColor(100, 150, 0, 0), cColor(100, 150, 0, 0)))
    End Sub

    Private Sub YZTileToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles YZTileToolStripMenuItem.Click
        PressMacro(tBuildYZ(1000, 200, True, True, cColor(100, 150, 150, 0), cColor(100, 150, 150, 0)))
    End Sub
    '---------------------------------------------------'
    Private Sub ClearToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        ResetZone()
    End Sub

    Private Sub ExportToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExportToolStripMenuItem.Click
        ExportModel()
    End Sub

    Private Sub LoadToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoadToolStripMenuItem.Click
        LoadZone()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        SaveZone()
    End Sub

    Private Sub ImportToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ImportToolStripMenuItem.Click
        ImportModel()
    End Sub
    '---------------------------------------------------'
    Private Sub LabelFontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LabelFontToolStripMenuItem.Click
        Dim bc As FontDialog = New FontDialog
        bc.ShowDialog()
        tbLabel.Font = bc.Font
        bc.Dispose()
    End Sub
#End Region
End Class

