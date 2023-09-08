Attribute VB_Name = "ModDisReModel"
Option Explicit

Public Sub ModelRun(ByVal Aimat As String)
    Dim Rd As New ADODB.Recordset, CheckFile As String
    Dim SSName As String
    Dim EStarttime() As Date, EEndtime() As Date
    Dim EStartday() As Date, EEndday() As Date, StartDay As Date, EndDay As Date
    Dim ETime() As Date, LongFDTime As Long, RFlN() As Integer
    Dim DTS As Integer, GGE() As Single, HTS As Integer, GEE() As Single
    Dim DRMOrderNo() As Integer
    Dim DRMWM() As Single
    Dim DRMSM() As Single
    Dim IDRMMNfS() As Single, DRMMNfS() As Single, NfS As Single
    Dim DRMMNfC() As Single, NfC As Single
    Dim DRMAlpha() As Single, DRMBmax() As Single
    Dim NoEvents  As Integer
    Dim PSCol() As Integer, PSRow() As Integer
    Dim PObs() As Single, QObs() As Single, EObs() As Single, AvgP() As Single
    Dim NextNoIJ() As Long, DRMGCNo As Long
    Dim DRMSSlope() As Single, DRMCSlope() As Single, DRMfc() As Single
    Dim TSteps As Long, TimeSeries() As Date
    Dim DRMET() As Single, DRMPO() As Single
    Dim LAI() As Single
    Dim DatumTIs As Single
    Dim SortingOrder() As Long, SortingRow() As Integer, SortingCol() As Integer
    Dim DRMQoutS() As Single, DRMHS() As Single
    Dim DRMQoutCh() As Single, DRMHCh() As Single
    Dim SumDRMQinS() As Single, SumDRMQinCh() As Single, SumDRMQinSCh() As Single
    Dim SType030() As Integer, SType30100() As Integer, LCover() As Integer
    Dim DRMIca() As Single, SumDRMIca As Single
    Dim DRMIch() As Single
    Dim Cellsize As Integer
    Dim DRMW() As Single, DRMS() As Single
    Dim OC As Single, ROC As Single
    Dim DRMKg() As Single, DRMKi() As Single
    Dim STSWC() As Single, STFC() As Single, STWP() As Single
    Dim FRTI() As Integer, DT As Integer
    Dim CSBeta As Single, CSAlpha As Single
    Dim QSim() As Single, SimQ() As Single, HSim() As Single, SimH() As Single
    Dim i As Integer, j As Integer, ii As Integer, jj As Integer, k As Long
    Dim i1 As Integer, j1 As Integer, m As Integer
    Dim DEMPrecision As Integer, ENo As Integer, ENo0 As Integer
    Dim IW As Single
    Dim Dis() As Single, Sumdis As Single, Sump As Single
    Dim SaP As Boolean, Kp As Integer
    Dim GLAI As Single, Scmax As Single, Cp As Single, Cvd As Single, Pcum() As Single
    Dim Icum() As Single
    Dim RFC() As Integer, Grfc As Integer
    Dim Gswc As Integer
    Dim Pnet() As Single
    Dim GW As Single, GS As Single, GWM As Single, GSM As Single
    Dim HumousT() As Single, ThickoVZ() As Single
    Dim ThitaS As Single, ThitaF As Single, ThitaW As Single, Thita As Single
    Dim SumKgKi As Single, KgKi As Single
    Dim Kg As Single, Ki As Single, KKg As Single, KKi As Single
    Dim Pe As Single
    Dim R As Single, Rs As Single, Ri As Single, Rg As Single, Qs As Single, Qi As Single, Qg As Single
    Dim FRSQinS As Single
    Dim FRQoutS As Single, HXS As Single, HdXS As Single, HS As Single, HdS As Single
    Dim FRQoutCh As Single, FRSQinCh As Single, HXCh As Single, HdXCh As Single, HCh As Single, HdCh As Single
    Dim AXCh As Single, ACH As Single, AdXCh As Single, AdCh As Single
    Dim BXCh As Single, BCh As Single, BdXCh As Single, BdCh As Single
    Dim DWD As Single
    Dim SS0 As Single, SSf As Single
    Dim DWU As Single
    Dim DRMW0() As Single, DRMS0() As Single, RFC0() As Integer
    Dim NCoe As Single, MNfCo As Single, HIndex As Single
    Dim AlUpper As Single, AlLower As Single, AlDeeper As Single
    Dim ZUpper As Single, ZLower As Single, ZDeeper As Single
    Dim GWUM As Single, GWU As Single, GWLM As Single, GWL As Single, GWDM As Single, GWD As Single
    Dim GridWU() As Single, GridWL() As Single, GridWD() As Single
    Dim GridWU0() As Single, GridWL0() As Single, GridWD0() As Single
    Dim GE As Single, GEU As Single, GEL As Single, ged As Single
    Dim Ek As Single, KEpC As Single, DeeperC As Single, Div As Integer
    Dim GridQi() As Single, GridQg() As Single, GGEE() As Single
    Dim GridFLC() As Single
    Dim JYear As Integer, JMonth As Integer, JDay As Integer, nd As Integer
    Dim GR As Single, GRs As Single, GRi As Single, GRg As Single
    Dim Cg As Single, Ci As Single, GridQi0() As Single, GridQg0() As Single
    Dim SumOQ As Single, SumSQ As Single, OPeak As Single, SPeak As Single, SumEE As Single, ET0out() As Single
    Dim OPeakTime As Integer, SPeakTime As Integer, SumPre As Single, DArea As Single
    Dim ORunoff As Single, SRunoff As Single, ONC As Single, SNC As Single, AvgOQ As Single
    Dim Qoutch As Single, Qouts As Single, Qouti As Single, Qoutg As Single, Qobs1 As Single
'    Dim MPeakE As Single, MRunoffE As Single, MTimeE As Single, MNashC As Single
    Dim UPRow() As Integer, UPCol() As Integer, UPName() As String, UPSimQ() As Single, UPQSim() As Single
    Dim UPSimH() As Single, UPHSim() As Single
    Dim GridWM() As Integer, GridSM() As Integer, GridQ() As Single
    Dim kk1 As Integer, kk2 As Integer, GridVch() As Single, GridVs() As Single
    
    With Rd
        .Open "select * from [WholeCatchPara] where YesOrNo= Yes", ConnectSys, adOpenStatic, adLockReadOnly
        If IsNull(Rd("Shortening")) Or IsNull(Rd("Time Interval(h)")) Or IsNull(Rd("Time Interval(h)")) Then
            MsgBox "请核实研究流域相关参数，程序中断！"
            .Close
            Exit Sub
        End If
        If IsNull(Rd("Ratio of OC")) Or IsNull(Rd("Outflow Coefficients")) Or IsNull(Rd("Hmax2")) Then
            MsgBox "请核实研究流域相关参数，程序中断！"
            .Close
            Exit Sub
        End If
        If IsNull(Rd("K")) Or IsNull(Rd("C")) Or IsNull(Rd("LUM")) Or IsNull(Rd("LLM")) Or IsNull(Rd("CG")) Or IsNull(Rd("CI")) Then
            MsgBox "请核实研究流域相关参数，程序中断！"
            .Close
            Exit Sub
        End If
        SSName = Rd("Shortening")
        DatumTIs = Rd("Time Interval(h)")
        OC = Rd("Outflow Coefficients")
        ROC = Rd("Ratio of OC")
        CSBeta = Rd("Beta")
        HIndex = Rd("Hmax2")
        DeeperC = Rd("C")
        KEpC = Rd("K")
        AlUpper = Rd("LUM")
        AlLower = Rd("LLM")
        AlDeeper = 1 - AlUpper - AlLower
        Cg = Rd("CG")
        Ci = Rd("CI")
        .Close
        
        .Open "select * from [HFlood Events-" & SSName & "] where [Purpose]='" & Aimat & "' order by [FloodNo]", ConnectSys, adOpenStatic, adLockReadOnly
        i = 0
        If IsNull(Rd("Purpose")) Or IsNull(Rd("Start time")) Or IsNull(Rd("End time")) Or IsNull(Rd("TI for FR(s)")) Then
            MsgBox "请核实率定洪水的相关参数，程序中断！"
            .Close
            Exit Sub
        End If
        NoEvents = .RecordCount
        If NoEvents = 0 Then
            MsgBox "数据表中没有率定的洪水，程序中断！"
            .Close
            Exit Sub
        End If
        .MoveFirst
        Do
            i = i + 1
            ReDim Preserve EStarttime(1 To i), EEndtime(1 To i), FRTI(1 To i), EStartday(1 To i), EEndday(1 To i), RFlN(1 To i) ', MoNfS(1 To i) ', EMonth(1 To i)
            RFlN(i) = Rd("FloodNo")
            EStarttime(i) = Rd("Start time")
            EEndtime(i) = Rd("End time")
            EStartday(i) = Format(EStarttime(i), "YYYY-MM-DD")
            EEndday(i) = Format(EEndtime(i), "YYYY-MM-DD")
            FRTI(i) = Rd("TI for FR(s)")
            .MoveNext
        Loop Until i = NoEvents
        .Close
        
        If NPStation = 0 Then
            MsgBox "P-Station缺乏站点信息，程序中断！"
            Exit Sub
        End If
        DEMPrecision = DDem * 3600
        .Open "select * from [P-Station] where [Watersheds]='" & StationName & "' order by [SubbasinNo]", ConnectSys, adOpenStatic, adLockReadOnly
        .MoveFirst
        ReDim SubbasinName(1 To NPStation), StrLon(1 To NPStation), StrLat(1 To NPStation), PSCol(1 To NPStation), PSRow(1 To NPStation)
        For i = 1 To NPStation
            SubbasinName(i) = Rd("PStationName")
            StrLat(i) = Rd("Latitude")
            StrLon(i) = Rd("Longitude")
            PSCol(i) = Int(((StrLon(i) - XllCorner) * 60 * (60 / DEMPrecision))) + 1
            PSRow(i) = Nx - Int(((StrLat(i) - YllCorner) * 60 * (60 / DEMPrecision)))
            .MoveNext
        Next i
        .Close
        
'        .Open "select * from [StationXY] where [Watersheds]='" & StationName & "' order by [SubbasinNo]", ConnectSys, adOpenStatic, adLockReadOnly
'        ReDim PSCol(1 To NPStation), PSRow(1 To NPStation)
'        .MoveFirst
'        For i = 1 To NPStation
'            PSCol(i) = Rd("Col")
'            PSRow(i) = Rd("Row")
'            .MoveNext
'        Next i
'        .Close
        
        If NUPoints > 0 Then
            ReDim UPRow(1 To NUPoints), UPCol(1 To NUPoints), UPName(1 To NUPoints)
            DEMPrecision = DDem * 3600
            .Open "select * from [Upstream outlets] where Watersheds='" & StationName & "'order by Points", ConnectSys, adOpenStatic, adLockReadOnly
            If .RecordCount <> NUPoints Then
                MsgBox "表Upstream outlets中上游出口点个数与表WholeCatchPara中Upstream points值不一致，请核实！"
                .Close
                Exit Sub
            End If
            .MoveFirst
            For i = 1 To NUPoints
                UPName(i) = Rd("Name")
                StrLat(i) = Rd("Latitude")
                StrLon(i) = Rd("Longitude")
                UPRow(i) = Nx - Int(((StrLat(i) - YllCorner) * 60 * (60 / DEMPrecision)))
                UPCol(i) = Int(((StrLon(i) - XllCorner) * 60 * (60 / DEMPrecision))) + 1
                .MoveNext
            Next i
            .Close
        End If
    End With

    ReDim STSWC(0 To 12), STFC(0 To 12), STWP(0 To 12)
    Rd.Open "select * from  [Soil Types] order by Category", ConnectSys, adOpenStatic, adLockReadOnly
    If Rd.RecordCount = 0 Then
        MsgBox "请在Soil Types表中输入土壤类型相关参数！"
        Rd.Close
        Exit Sub
    End If
    Rd.MoveFirst
    i = 0
    Do
        STSWC(i) = Rd("SWC")
        STFC(i) = Rd("FC")
        STWP(i) = Rd("WP")
        i = i + 1
        Rd.MoveNext
    Loop Until Rd.EOF
    Rd.Close

    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "DEM.asc")
    If CheckFile = "" Then
        MsgBox "缺少" & StationName & "流域DEM高程ASC文件！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "累积汇水面积.asc")
    If CheckFile = "" Then
        MsgBox "缺少" & StationName & "流域累积汇水面积ASC文件！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "水系.asc")
    If CheckFile = "" Then
        MsgBox "缺少" & StationName & "流域水系ASC文件！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "栅格流向.asc")
    If CheckFile = "" Then
        MsgBox "缺少" & StationName & "流域栅格流向ASC文件！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "栅格演算次序.asc")
    If CheckFile = "" Then
        MsgBox "请先进行栅格间汇流演算次序计算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "张力水蓄水容量.asc")
    If CheckFile = "" Then
        MsgBox "请先进行包气带厚度估算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "自由水蓄水容量.asc")
    If CheckFile = "" Then
        MsgBox "请先进行包气带厚度估算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "河道汇流糙率.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "河道宽度指数.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "坡面汇流糙率.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "径流分配比例.asc")
    If CheckFile = "" Then
        MsgBox "请先进行栅格间汇流演算次序计算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "最陡坡度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行最陡坡度提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "河道最大宽度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\input\" & StationName & "\" & StationName & "植被类型.asc")
    If CheckFile = "" Then
        MsgBox "请先给定植被类型数据！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\input\" & StationName & "\" & StationName & "0-30cm土壤类型.asc")
    If CheckFile = "" Then
        MsgBox "请先给定0-30cm土壤类型数据！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\input\" & StationName & "\" & StationName & "30-100cm土壤类型.asc")
    If CheckFile = "" Then
        MsgBox "请先给定30-100cm土壤类型数据！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "腐殖质土厚度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行包气带厚度估算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "栅格河道坡度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "包气带厚度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行包气带厚度估算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If

    ReDim E(1 To Nx, 1 To Ny), WaterArea(1 To Nx, 1 To Ny), RiverPoint(1 To Nx, 1 To Ny), FlowDirection(1 To Nx, 1 To Ny)
    ReDim DRMOrderNo(1 To Nx, 1 To Ny), DRMWM(1 To Nx, 1 To Ny), DRMSM(1 To Nx, 1 To Ny)
    ReDim DRMMNfC(1 To Nx, 1 To Ny), DRMAlpha(1 To Nx, 1 To Ny), DRMMNfS(1 To Nx, 1 To Ny)
    ReDim NextNoIJ(1 To Nx, 1 To Ny), DRMBmax(1 To Nx, 1 To Ny), IDRMMNfS(1 To Nx, 1 To Ny)
    ReDim DRMSSlope(1 To Nx, 1 To Ny), DRMCSlope(1 To Nx, 1 To Ny), DRMfc(1 To Nx, 1 To Ny)
    ReDim SType030(1 To Nx, 1 To Ny), SType30100(1 To Nx, 1 To Ny), TopIndex(1 To Nx, 1 To Ny), LCover(1 To Nx, 1 To Ny)
    ReDim DRMKg(1 To Nx, 1 To Ny), DRMKi(1 To Nx, 1 To Ny), HumousT(1 To Nx, 1 To Ny), ThickoVZ(1 To Nx, 1 To Ny)
    ReDim GridWM(1 To Nx, 1 To Ny), GridSM(1 To Nx, 1 To Ny), GridVch(1 To Nx, 1 To Ny), GridVs(1 To Nx, 1 To Ny), GridQ(1 To Nx, 1 To Ny)
    
    Open App.Path & "\Input\" & StationName & "\" & StationName & "DEM.asc" For Input As #1
    Input #1, Str
    Input #1, Str
    Input #1, Str
    Input #1, Str
    Input #1, Str
    Input #1, Str
    Open App.Path & "\Input\" & StationName & "\" & StationName & "累积汇水面积.asc" For Input As #2
    Input #2, Str
    Input #2, Str
    Input #2, Str
    Input #2, Str
    Input #2, Str
    Input #2, Str
    Open App.Path & "\Input\" & StationName & "\" & StationName & "水系.asc" For Input As #3
    Input #3, Str
    Input #3, Str
    Input #3, Str
    Input #3, Str
    Input #3, Str
    Input #3, Str
    Open App.Path & "\Input\" & StationName & "\" & StationName & "栅格流向.asc" For Input As #4
    Input #4, Str
    Input #4, Str
    Input #4, Str
    Input #4, Str
    Input #4, Str
    Input #4, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "栅格演算次序.asc" For Input As #5
    Input #5, Str
    Input #5, Str
    Input #5, Str
    Input #5, Str
    Input #5, Str
    Input #5, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "张力水蓄水容量.asc" For Input As #6
    Input #6, Str
    Input #6, Str
    Input #6, Str
    Input #6, Str
    Input #6, Str
    Input #6, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "自由水蓄水容量.asc" For Input As #7
    Input #7, Str
    Input #7, Str
    Input #7, Str
    Input #7, Str
    Input #7, Str
    Input #7, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "河道汇流糙率.asc" For Input As #8
    Input #8, Str
    Input #8, Str
    Input #8, Str
    Input #8, Str
    Input #8, Str
    Input #8, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "河道宽度指数.asc" For Input As #9
    Input #9, Str
    Input #9, Str
    Input #9, Str
    Input #9, Str
    Input #9, Str
    Input #9, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "坡面汇流糙率.asc" For Input As #10
    Input #10, Str
    Input #10, Str
    Input #10, Str
    Input #10, Str
    Input #10, Str
    Input #10, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "径流分配比例.asc" For Input As #11
    Input #11, Str
    Input #11, Str
    Input #11, Str
    Input #11, Str
    Input #11, Str
    Input #11, Str
    Open App.Path & "\Input\" & StationName & "\" & StationName & "最陡坡度.asc" For Input As #12
    Input #12, Str
    Input #12, Str
    Input #12, Str
    Input #12, Str
    Input #12, Str
    Input #12, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "河道最大宽度.asc" For Input As #13
    Input #13, Str
    Input #13, Str
    Input #13, Str
    Input #13, Str
    Input #13, Str
    Input #13, Str
    Open App.Path & "\input\" & StationName & "\" & StationName & "0-30cm土壤类型.asc" For Input As #14
    Input #14, Str
    Input #14, Str
    Input #14, Str
    Input #14, Str
    Input #14, Str
    Input #14, Str
    Open App.Path & "\input\" & StationName & "\" & StationName & "30-100cm土壤类型.asc" For Input As #15
    Input #15, Str
    Input #15, Str
    Input #15, Str
    Input #15, Str
    Input #15, Str
    Input #15, Str
    Open App.Path & "\input\" & StationName & "\" & StationName & "植被类型.asc" For Input As #16
    Input #16, Str
    Input #16, Str
    Input #16, Str
    Input #16, Str
    Input #16, Str
    Input #16, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "腐殖质土厚度.asc" For Input As #17
    Input #17, Str
    Input #17, Str
    Input #17, Str
    Input #17, Str
    Input #17, Str
    Input #17, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "栅格河道坡度.asc" For Input As #18
    Input #18, Str
    Input #18, Str
    Input #18, Str
    Input #18, Str
    Input #18, Str
    Input #18, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "包气带厚度.asc" For Input As #20
    Input #20, Str
    Input #20, Str
    Input #20, Str
    Input #20, Str
    Input #20, Str
    Input #20, Str
    DRMGCNo = 0
    SumKgKi = 0
    Cellsize = Int(DDem / (3 / 3600) * 90)
    For i = 1 To Nx
        For j = 1 To Ny
            Input #1, E(i, j)
            Input #2, WaterArea(i, j)
            Input #3, RiverPoint(i, j)
            Input #4, FlowDirection(i, j)
            Input #5, DRMOrderNo(i, j)
            Input #6, DRMWM(i, j)
            Input #7, DRMSM(i, j)
            Input #8, DRMMNfC(i, j)
            Input #9, DRMAlpha(i, j)
            Input #10, IDRMMNfS(i, j)
            Input #11, DRMfc(i, j)
            Input #12, DRMSSlope(i, j)
            Input #13, DRMBmax(i, j)
            Input #14, SType030(i, j)
            Input #15, SType30100(i, j)
            Input #16, LCover(i, j)
            Input #17, HumousT(i, j)
            Input #18, DRMCSlope(i, j)
            Input #20, ThickoVZ(i, j)
            If E(i, j) <> Nodata Then
                DRMGCNo = DRMGCNo + 1
                NextNoIJ(i, j) = DRMGCNo
                If HumousT(i, j) <= 300 Then
                    ThitaS = STSWC(SType030(i, j))
                    ThitaF = STFC(SType030(i, j))
                    ThitaW = STWP(SType030(i, j))
                Else
                    ThitaS = STSWC(SType030(i, j)) * (300 / HumousT(i, j)) + STSWC(SType30100(i, j)) * (1 - 300 / HumousT(i, j))
                    ThitaF = STFC(SType030(i, j)) * (300 / HumousT(i, j)) + STFC(SType30100(i, j)) * (1 - 300 / HumousT(i, j))
                    ThitaW = STWP(SType030(i, j)) * (300 / HumousT(i, j)) + STWP(SType30100(i, j)) * (1 - 300 / HumousT(i, j))
                End If
                DRMKi(i, j) = ((ThitaF / ThitaS) ^ OC) / (1 + ROC / (1 + 2 * (1 - ThitaW)))
                DRMKg(i, j) = (ThitaF / ThitaS) ^ OC - DRMKi(i, j)
                SumKgKi = SumKgKi + (ThitaF / ThitaS) ^ OC
            End If
        Next j
    Next i
    KgKi = SumKgKi / DRMGCNo
    Close #1
    Close #2
    Close #3
    Close #4
    Close #5
    Close #6
    Close #7
    Close #8
    Close #9
    Close #10
    Close #11
    Close #12
    Close #13
    Close #14
    Close #15
    Close #16
    Close #17
    Close #18
    Close #20
    DArea = Garea * DRMGCNo
    
    ReDim SortingOrder(1 To DRMGCNo), SortingRow(1 To DRMGCNo), SortingCol(1 To DRMGCNo)
    With Rd
        .Open "select * from [CalSorting] where [Watersheds]='" & StationName & "' order by [CalOrder],[Row],[Col] asc", ConnectSys, adOpenStatic, adLockReadOnly
        If .RecordCount <> DRMGCNo Then
            Rd.Close
            MsgBox "请确认栅格演算次序表！程序中断！"
            Exit Sub
        End If
        .MoveFirst
        i = 1
        Do
          SortingOrder(i) = Rd("CalOrder")
          SortingRow(i) = Rd("Row")
          SortingCol(i) = Rd("Col")
          .MoveNext
          i = i + 1
        Loop Until .EOF
        .Close
    End With
    IW = 2
    For ENo = 1 To NoEvents
        ENo0 = RFlN(ENo)
        Rd.Open "select * from [HObserved-" & SSName & "] where [时间] between  #" & Format(EStarttime(ENo), "YYYY-MM-DD HH:NN:SS") & "# and #" & Format(EEndtime(ENo), "YYYY-MM-DD HH:NN:SS") & "# order by [时间]", ConnectSys, adOpenStatic, adLockReadOnly
        TSteps = Rd.RecordCount
        If TSteps <> (DateDiff("H", EStarttime(ENo), EEndtime(ENo)) / DatumTIs + 1) Then
            MsgBox "第" & ENo0 & "场洪水资料有误，程序中断！"
            Rd.Close
           ' Exit Sub
           Stop
        End If
        DT = FRTI(ENo)
        ReDim PObs(1 To ((TSteps - 1) * Int(DatumTIs * 3600 / DT) + 1), 1 To NPStation), EObs(1 To TSteps), TimeSeries(1 To TSteps), AvgP(1 To TSteps)
        ReDim QObs(1 To TSteps), QSim(1 To TSteps), HSim(1 To TSteps), GEE(1 To TSteps * Int(DatumTIs * 3600 / DT))
        If NUPoints > 0 Then
            ReDim UPQSim(1 To TSteps, 1 To NUPoints), UPHSim(1 To TSteps, 1 To NUPoints)
        End If
        Rd.MoveFirst
        For i = 1 To TSteps
            If i = 1 Then
                For j = 1 To NPStation
                    If IsNull(Rd(SubbasinName(j))) Then
                        PObs(1, j) = 0
                        AvgP(i) = AvgP(i)
                    Else
                        PObs(1, j) = Rd(SubbasinName(j))
                        AvgP(i) = AvgP(i) + Rd(SubbasinName(j))
                    End If
                Next j
            Else
                For j = 1 To NPStation
                    If IsNull(Rd(SubbasinName(j))) Then
                        AvgP(i) = AvgP(i)
                    Else
                        AvgP(i) = AvgP(i) + Rd(SubbasinName(j))
                    End If
                    For m = 1 To Int(DatumTIs * 3600 / DT)
                        If IsNull(Rd(SubbasinName(j))) Then
                            PObs((i - 2) * Int(DatumTIs * 3600 / DT) + m + 1, j) = 0
                        Else
                            PObs((i - 2) * Int(DatumTIs * 3600 / DT) + m + 1, j) = Rd(SubbasinName(j)) / Int(DatumTIs * 3600 / DT)
                        End If
                    Next m
                Next j
            End If
            AvgP(i) = AvgP(i) / NPStation
            TimeSeries(i) = Rd("时间")
            If IsNull(Rd("实测流量")) Then
                QObs(i) = 0
            Else
                QObs(i) = Rd("实测流量")
            End If
            Rd.MoveNext
        Next i
        Rd.Close
        StartDay = DateAdd("D", -1, EStartday(ENo))
        EndDay = DateAdd("D", 1, EEndday(ENo))
        DTS = DateDiff("D", StartDay, EndDay) + 1
        Rd.Open "select * from [DObserved-" & SSName & "] where [时间] between  #" & Format(StartDay, "YYYY-MM-DD") & "# and #" & Format(EndDay, "YYYY-MM-DD") & "# order by [时间]", ConnectSys, adOpenStatic, adLockReadOnly
        If Rd.RecordCount = 0 Then
            MsgBox "第" & ENo0 & "场洪水资料缺失，程序中断！"
            Rd.Close
            Exit Sub
        End If
        ReDim GGEE(1 To DTS)
        Rd.MoveFirst
        For i = 1 To DTS
            If IsNull(Rd("蒸发")) Then
                GGEE(i) = 0
            Else
                GGEE(i) = Rd("蒸发")
            End If
            Rd.MoveNext
        Next i
        HTS = (DTS - 1) * 24 + 8
        ReDim GGE(1 To HTS)
        For i = 1 To 8
            GGE(i) = GGEE(1) / 24
        Next i
        For i = 1 To DTS - 1
            For k = 1 To 24
                GGE((i - 1) * 24 + 8 + k) = GGEE(i + 1) / 24
            Next k
        Next i
        HTS = DateDiff("H", StartDay, EStarttime(ENo)) - 1
        For i = 1 To TSteps
            EObs(i) = GGE(HTS + i)
            For k = 1 To Int(DatumTIs * 3600 / DT)
                GEE((i - 1) * Int(DatumTIs * 3600 / DT) + k) = EObs(i) / Int(DatumTIs * 3600 / DT)
            Next k
        Next i
        Rd.Close
        
        ReDim LAI(0 To 13, 1 To 12)
        Rd.Open "select * from [Land Cover] order by [Category]", ConnectSys, adOpenStatic, adLockReadOnly
        Rd.MoveFirst
        For i = 0 To 13
            LAI(i, 1) = Rd("LAI-Jan")
            LAI(i, 2) = Rd("LAI-Feb")
            LAI(i, 3) = Rd("LAI-Mar")
            LAI(i, 4) = Rd("LAI-Apr")
            LAI(i, 5) = Rd("LAI-May")
            LAI(i, 6) = Rd("LAI-Jun")
            LAI(i, 7) = Rd("LAI-Jul")
            LAI(i, 8) = Rd("LAI-Aug")
            LAI(i, 9) = Rd("LAI-Sep")
            LAI(i, 10) = Rd("LAI-Oct")
            LAI(i, 11) = Rd("LAI-Nov")
            LAI(i, 12) = Rd("LAI-Dec")
            Rd.MoveNext
        Next i
        Rd.Close
        
        ReDim GridQi(1 To Nx, 1 To Ny), gridQg2(1 To Nx, 1 To Ny), GridQi0(1 To Nx, 1 To Ny), GridQg0(1 To Nx, 1 To Ny)
        ReDim DRMW(1 To Nx, 1 To Ny), DRMS(1 To Nx, 1 To Ny), GridFLC(1 To Nx, 1 To Ny)
        ReDim GridWU(1 To Nx, 1 To Ny), GridWL(1 To Nx, 1 To Ny), GridWD(1 To Nx, 1 To Ny)
        ReDim GridWU0(1 To Nx, 1 To Ny), GridWL0(1 To Nx, 1 To Ny), GridWD0(1 To Nx, 1 To Ny)
        
        LongFDTime = DatePart("YYYY", StartDay) * 10000 + DatePart("M", StartDay) * 100 + DatePart("D", StartDay)
        JYear = DatePart("YYYY", StartDay)
        JMonth = DatePart("M", StartDay)
        CheckFile = Dir(App.Path & "\output\" & StationName & "\日模张力水容量\" & JYear & "\" & LongFDTime & ".asc")
        If CheckFile = "" Then
            MsgBox "请先进行第" & ENo0 & "次洪水的日洪模拟！程序中断！", vbExclamation + vbInformation, "警告："
            Exit Sub
        End If
        Open App.Path & "\output\" & StationName & "\日模张力水容量\" & JYear & "\" & LongFDTime & ".asc" For Input As #90
        Open App.Path & "\output\" & StationName & "\日模自由水容量\" & JYear & "\" & LongFDTime & ".asc" For Input As #91
        Open App.Path & "\output\" & StationName & "\日模上层张力水容量\" & JYear & "\" & LongFDTime & ".asc" For Input As #92
        Open App.Path & "\output\" & StationName & "\日模下层张力水容量\" & JYear & "\" & LongFDTime & ".asc" For Input As #93
        Open App.Path & "\output\" & StationName & "\日模植被覆盖率\" & JYear & "\" & LongFDTime & ".asc" For Input As #94
        Input #90, Str
        Input #90, Str
        Input #90, Str
        Input #90, Str
        Input #90, Str
        Input #90, Str
        Input #91, Str
        Input #91, Str
        Input #91, Str
        Input #91, Str
        Input #91, Str
        Input #91, Str
        Input #92, Str
        Input #92, Str
        Input #92, Str
        Input #92, Str
        Input #92, Str
        Input #92, Str
        Input #93, Str
        Input #93, Str
        Input #93, Str
        Input #93, Str
        Input #93, Str
        Input #93, Str
        Input #94, Str
        Input #94, Str
        Input #94, Str
        Input #94, Str
        Input #94, Str
        Input #94, Str
        For ii = 1 To Nx
            For jj = 1 To Ny
                Input #90, DRMW(ii, jj)
                Input #91, DRMS(ii, jj)
                Input #92, GridWU(ii, jj)
                Input #93, GridWL(ii, jj)
                Input #94, GridFLC(ii, jj)
                If DRMW(ii, jj) <> Nodata Then
                    GridQi(ii, jj) = QObs(1) / DRMGCNo / 2
                    GridQg(ii, jj) = QObs(1) / DRMGCNo / 2
                End If
                GridWM(ii, jj) = Nodata
                GridSM(ii, jj) = Nodata
                GridVch(ii, jj) = Nodata
                GridVs(ii, jj) = Nodata
                GridQ(ii, jj) = Nodata
            Next
        Next
        Close #90
        Close #91
        Close #92
        Close #93
        Close #94

        SumDRMIca = 0
        Div = 5
        Qobs1 = QObs(1) / 2
        TSteps = (TSteps - 1) * Int(DatumTIs * 3600 / DT) + 1
        ReDim DRMPO(1 To DRMGCNo), DRMET(1 To DRMGCNo)
        ReDim DRMQoutS(1 To DRMGCNo), DRMHS(0 To 1, 1 To DRMGCNo)
        ReDim SumDRMQinS(1 To DRMGCNo), SumDRMQinCh(1 To DRMGCNo), SumDRMQinSCh(1 To DRMGCNo)
        ReDim Dis(1 To NPStation), Pcum(1 To Nx, 1 To Ny)
        ReDim Icum(0 To 1, 1 To DRMGCNo), DRMIca(1 To DRMGCNo)
        ReDim DRMIch(1 To DRMGCNo), SimQ(0 To TSteps), SimH(0 To TSteps)
        ReDim RFC(1 To Nx, 1 To Ny), Pnet(1 To DRMGCNo), ETime(0 To TSteps)
        ReDim DRMQoutCh(1 To DRMGCNo), DRMHCh(0 To 1, 1 To DRMGCNo)
        ReDim DRMW0(1 To Nx, 1 To Ny), DRMS0(1 To Nx, 1 To Ny), RFC0(1 To Nx, 1 To Ny)
        If NUPoints > 0 Then
            ReDim UPSimQ(0 To TSteps, 1 To NUPoints), UPSimH(0 To TSteps, 1 To NUPoints)
        End If
        
        For i = 0 To TSteps - 1
            If i = 0 Then
                ETime(1) = EStarttime(ENo)
            Else
                ETime(i + 1) = DateAdd("S", DT, ETime(i))
            End If
            JDay = DatePart("M", ETime(i + 1))
            If JDay <> JMonth Then
                LongFDTime = DatePart("YYYY", ETime(i + 1)) * 10000 + DatePart("M", ETime(i + 1)) * 100 + DatePart("D", ETime(i + 1))
                Open App.Path & "\output\" & StationName & "\日模植被覆盖率\" & JYear & "\" & LongFDTime & ".asc" For Input As #94
                Input #94, Str
                Input #94, Str
                Input #94, Str
                Input #94, Str
                Input #94, Str
                Input #94, Str
                For ii = 1 To Nx
                    For jj = 1 To Ny
                        Input #94, GridFLC(ii, jj)
                    Next
                Next
                Close #94
                JMonth = JDay
            End If
            
            For k = 1 To DRMGCNo
                ii = SortingRow(k)
                jj = SortingCol(k)
                If i = 0 Then
                    DRMMNfS(ii, jj) = IDRMMNfS(ii, jj)
                End If
                Sumdis = 0
                Sump = 0
                SaP = False
                Kp = 1
                For j = 1 To NPStation
                    Dis(j) = ((ii - PSRow(j)) ^ 2 + (jj - PSCol(j)) ^ 2) ^ 0.5
                    If Dis(j) = 0 Then
                        Kp = j
                        SaP = True
                        Exit For
                    End If
                    Sump = Sump + PObs(i + 1, j) * Dis(j) ^ (-IW)
                    Sumdis = Sumdis + Dis(j) ^ (-IW)
                Next j
                If SaP = True Then
                    DRMPO(NextNoIJ(ii, jj)) = PObs(i + 1, Kp)
                Else
                    DRMPO(NextNoIJ(ii, jj)) = Sump / Sumdis
                End If
                DRMET(NextNoIJ(ii, jj)) = KEpC * GEE(i + 1)
                GLAI = LAI(LCover(ii, jj), JMonth)
                Scmax = 0.935 + 0.498 * GLAI - 0.00575 * GLAI ^ 2
                Cp = GridFLC(ii, jj)
                Cvd = 0.046 * GLAI
                Pcum(ii, jj) = Pcum(ii, jj) + DRMPO(NextNoIJ(ii, jj))
                Icum(1, NextNoIJ(ii, jj)) = Cp * Scmax * (1 - Exp((-Cvd) * Pcum(ii, jj) / Scmax))
                DRMIca(NextNoIJ(ii, jj)) = Icum(1, NextNoIJ(ii, jj)) - Icum(0, NextNoIJ(ii, jj))
                Icum(0, NextNoIJ(ii, jj)) = Icum(1, NextNoIJ(ii, jj))
                If DRMET(NextNoIJ(ii, jj)) >= DRMIca(NextNoIJ(ii, jj)) Then
                    DRMET(NextNoIJ(ii, jj)) = DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj))
                    DRMIca(NextNoIJ(ii, jj)) = 0
                    If DRMET(NextNoIJ(ii, jj)) >= SumDRMIca Then
                        DRMET(NextNoIJ(ii, jj)) = DRMET(NextNoIJ(ii, jj)) - SumDRMIca
                        SumDRMIca = 0
                    Else
                        SumDRMIca = SumDRMIca - DRMET(NextNoIJ(ii, jj))
                        DRMET(NextNoIJ(ii, jj)) = 0
                    End If
                Else
                    DRMIca(NextNoIJ(ii, jj)) = DRMIca(NextNoIJ(ii, jj)) - DRMET(NextNoIJ(ii, jj))
                    SumDRMIca = SumDRMIca + DRMIca(NextNoIJ(ii, jj))
                    DRMET(NextNoIJ(ii, jj)) = 0
                End If
                
                If RiverPoint(ii, jj) = 1 Then
                    If FlowDirection(ii, jj) = 0 Or FlowDirection(ii, jj) = 2 Or FlowDirection(ii, jj) = 4 Or FlowDirection(ii, jj) = 6 Then
                        DRMIch(NextNoIJ(ii, jj)) = (DRMPO(NextNoIJ(ii, jj))) * ((DRMBmax(ii, jj) * Cellsize ^ 0.5 * 10 ^ (-6)) / Garea)
                    Else
                        DRMIch(NextNoIJ(ii, jj)) = (DRMPO(NextNoIJ(ii, jj))) * ((DRMBmax(ii, jj) * Cellsize * 10 ^ (-6)) / Garea)
                    End If
                Else
                    DRMIch(NextNoIJ(ii, jj)) = 0
                End If
                
                If RFC(ii, jj) = 1 Then
                    Pnet(NextNoIJ(ii, jj)) = DRMPO(NextNoIJ(ii, jj)) - DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj)) - DRMIch(NextNoIJ(ii, jj))
                    FRSQinS = SumDRMQinS(NextNoIJ(ii, jj))
                Else
                    Pnet(NextNoIJ(ii, jj)) = DRMPO(NextNoIJ(ii, jj)) + (SumDRMQinS(NextNoIJ(ii, jj)) * DT * 0.001 / Garea) - DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj)) - DRMIch(NextNoIJ(ii, jj))
                    FRSQinS = 0
                End If
                Pe = Pnet(NextNoIJ(ii, jj))
                Ek = DRMET(NextNoIJ(ii, jj))
                GWM = DRMWM(ii, jj)
                GSM = DRMSM(ii, jj)
                GW = DRMW(ii, jj)
                GS = DRMS(ii, jj)
                GWU = GridWU(ii, jj)
                GWL = GridWL(ii, jj)
                GWD = GW - GWU - GWL
                Kg = DRMKg(ii, jj)
                Ki = DRMKi(ii, jj)
                KKi = (1 - ((1 - (Kg + Ki)) ^ ((DT / 3600) / 24))) / (1 + Kg / Ki)
                KKg = KKi * Kg / Ki
                Kg = KKg
                Ki = KKi
                DRMW0(ii, jj) = DRMW(ii, jj): DRMS0(ii, jj) = DRMS(ii, jj): RFC0(ii, jj) = RFC(ii, jj)
                GridWU0(ii, jj) = GridWU(ii, jj): GridWL0(ii, jj) = GridWL(ii, jj)
                Grfc = RFC(ii, jj)
                ZUpper = AlUpper * ThickoVZ(ii, jj)
                ZLower = AlLower * ThickoVZ(ii, jj)
                ZDeeper = AlDeeper * ThickoVZ(ii, jj)
                If ZUpper > 300 Then
                    GWUM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * 300 + (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * (ZUpper - 300)
                    GWLM = (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * ZLower
                Else
                    GWUM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * ZUpper
                    If ZUpper + ZLower > 300 Then
                        GWLM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * (300 - ZUpper) + (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * (ZLower - 300 + ZUpper)
                    Else
                        GWLM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * ZLower
                    End If
                End If
                GWDM = GWM - GWUM - GWLM
                If Pe <= 0 Then
                    Grfc = 0
                    Gswc = 0
                    If GWU + Pe < 0# Then
                        GEU = GWU + Ek + Pe
                        GWU = 0
                        GEL = (Ek - GEU) * GWL / GWLM
                        If GWL < DeeperC * GWLM Then
                            GEL = DeeperC * (Ek - GEU)
                        End If
                        If (GWL - GEL) < 0# Then
                            ged = GEL - GWL
                            GEL = GWL
                            GWL = 0
                            GWD = GWD - ged
                        Else
                            ged = 0
                            GWL = GWL - GEL
                        End If
                    Else
                        GEU = Ek
                        GEL = 0
                        ged = 0
                        GWU = GWU + Pe
                    End If
                    GW = GWU + GWL + GWD
                    If GW < 0 Then
                        GEU = 0
                        GEL = 0
                        ged = DRMW(ii, jj)
                        GE = ged
                        GWU = 0
                        GWL = 0
                        GWD = 0
                        GW = 0
                    End If
                    GE = GEU + GEL + ged
                    SumDRMIca = Ek - GE
                    If SumDRMIca < -0.0001 Then
                        SumDRMIca = 0
                    End If
                    GRs = 0
                    GRg = GS * Kg
                    GRi = GS * Ki
                    GS = GS * (1 - Kg - Ki)
                Else
                    GEU = Ek
                    GEL = 0
                    ged = 0
                    GE = Ek
                    nd = Int(Pe / Div) + 1
                    If Pe Mod Div = 0 Then
                        If Pe - Int(Pe) = 0 Then
                            nd = nd - 1
                        End If
                    End If
                    ReDim PPe(1 To nd)
                    For m = 1 To nd - 1
                        PPe(m) = Div
                    Next m
                    PPe(nd) = Pe - (nd - 1) * Div
                    GRs = 0
                    GRg = 0
                    GRi = 0
                    KKi = (1# - (1# - (Kg + Ki)) ^ (1# / nd)) / (Kg + Ki)
                    KKg = KKi * Kg
                    KKi = KKi * Ki
                    For m = 1 To nd
                        If PPe(m) + GW < GWM Then
                            Grfc = 0
                            Gswc = 0
                            If PPe(m) + GWU < GWUM Then
                                GWU = PPe(m) + GWU
                            ElseIf GWL + GWUM - GWU + PPe(m) < GWLM Then
                                GWL = GWL + GWUM - GWU + PPe(m)
                                GWU = GWUM
                            Else
                                GWD = GWD + PPe(m) - (GWUM - GWU) - (GWLM - GWL)
                                GWU = GWUM
                                GWL = GWLM
                            End If
                            GW = GWU + GWL + GWD
                            GRs = GRs
                            GRg = GRg + GS * KKg
                            GRi = GRi + GS * KKi
                            GS = GS * (1 - KKg - KKi)
                        Else
                            GR = PPe(m) + GW - GWM
                            GWU = GWUM
                            GWL = GWLM
                            GWD = GWDM
                            GW = GWM
                            Grfc = 1
                            If GR + GS <= GSM Then
                                GRs = GRs
                                GS = GS + GR
                                GRg = GRg + GS * KKg
                                GRi = GRi + GS * KKi
                                GS = GS * (1 - KKg - KKi)
                            Else
                                GRs = GRs + GR + GS - GSM
                                GS = GSM
                                Gswc = 1
                                GRg = GRg + GSM * KKg
                                GRi = GRi + GSM * KKi
                                GS = GSM * (1 - KKg - KKi)
                            End If
                        End If
                    Next m
                End If
                
                Qs = GRs * Garea * 1000 / DT
                Qi = GRi * Garea * 1000 / DT
                Qg = GRg * Garea * 1000 / DT
                GW = GWU + GWL + GWD
                DRMW(ii, jj) = GW
                DRMS(ii, jj) = GS
                GridWU(ii, jj) = GWU
                GridWL(ii, jj) = GWL
                GridWD(ii, jj) = GWD
                RFC(ii, jj) = Grfc
                GridWM(ii, jj) = Grfc
                GridSM(ii, jj) = Gswc
                DRMET(NextNoIJ(ii, jj)) = GE
                FRQoutS = DRMQoutS(NextNoIJ(ii, jj))
                HS = DRMHS(0, NextNoIJ(ii, jj))
                HXS = HS + (FRSQinS + Qs - FRQoutS) * 0.001 * DT / Garea
                HXS = IIf(HXS > 0, HXS, 0)
                DRMHS(1, NextNoIJ(ii, jj)) = HXS
                DWD = IIf(FlowDirection(ii, jj) = 0 Or FlowDirection(ii, jj) = 2 Or FlowDirection(ii, jj) = 4 Or FlowDirection(ii, jj) = 6, 2 ^ 0.5 * Cellsize, Cellsize)
                CSAlpha = DRMAlpha(ii, jj)
                If i = 0 Then
                    NfC = DRMMNfC(SortingRow(DRMGCNo), SortingCol(DRMGCNo))
                    SSf = DRMCSlope(SortingRow(DRMGCNo), SortingCol(DRMGCNo))
                    CSAlpha = DRMAlpha(SortingRow(DRMGCNo), SortingCol(DRMGCNo))
                    DRMHCh(0, NextNoIJ(SortingRow(DRMGCNo), SortingCol(DRMGCNo))) = (NfC ^ 3 * CSAlpha ^ (-3) * (1 + CSBeta) ^ 5 * SSf ^ (-3 / 2) * Qobs1 ^ 3) ^ (1 / (3 * CSBeta + 5))
                    If RiverPoint(ii, jj) = 1 Then
                        DRMHCh(0, NextNoIJ(ii, jj)) = DRMHCh(0, NextNoIJ(SortingRow(DRMGCNo), SortingCol(DRMGCNo))) * (WaterArea(ii, jj) * Garea / DArea) ^ HIndex
                    End If
                End If
                HCh = DRMHCh(0, NextNoIJ(ii, jj))
                ACH = (CSAlpha / (1 + CSBeta)) * HCh ^ (1 + CSBeta)
                BCh = CSAlpha * HCh ^ CSBeta
                FRQoutCh = DRMQoutCh(NextNoIJ(ii, jj))
                FRSQinCh = Qi + Qg + (DRMIch(NextNoIJ(ii, jj)) * DWD * DRMBmax(ii, jj) * 0.001 / DT) + SumDRMQinSCh(NextNoIJ(ii, jj)) + SumDRMQinCh(NextNoIJ(ii, jj))
                AXCh = ACH + (DT / DWD) * (FRSQinCh - FRQoutCh)
                AXCh = IIf(AXCh > 0, AXCh, 0)
                HXCh = ((1 + CSBeta) * AXCh / CSAlpha) ^ (1 / (1 + CSBeta))
                DRMHCh(1, NextNoIJ(ii, jj)) = HXCh
            Next k
            ReDim SumDRMQinS(1 To DRMGCNo), SumDRMQinCh(1 To DRMGCNo), SumDRMQinSCh(1 To DRMGCNo)
            For k = 1 To DRMGCNo
                ii = SortingRow(k)
                jj = SortingCol(k)
                Select Case FlowDirection(ii, jj)
                    Case 0
                        i1 = ii - 1
                        j1 = jj + 1
                        DWD = 2 ^ 0.5 * Cellsize
                    Case 1
                        i1 = ii
                        j1 = jj + 1
                        DWD = Cellsize
                    Case 2
                        i1 = ii + 1
                        j1 = jj + 1
                        DWD = 2 ^ 0.5 * Cellsize
                    Case 3
                        i1 = ii + 1
                        j1 = jj
                        DWD = Cellsize
                    Case 4
                        i1 = ii + 1
                        j1 = jj - 1
                        DWD = 2 ^ 0.5 * Cellsize
                    Case 5
                        i1 = ii
                        j1 = jj - 1
                        DWD = Cellsize
                    Case 6
                        i1 = ii - 1
                        j1 = jj - 1
                        DWD = 2 ^ 0.5 * Cellsize
                    Case 7
                        i1 = ii - 1
                        j1 = jj
                        DWD = Cellsize
                    Case 8
                        i1 = ii
                        j1 = jj
                        DWD = Cellsize
                    Case Else
                End Select
                
                SS0 = DRMSSlope(ii, jj)
                NfS = DRMMNfS(ii, jj)
                HXS = DRMHS(1, NextNoIJ(ii, jj)) * 0.001
                HdXS = DRMHS(1, NextNoIJ(i1, j1)) * 0.001
                SSf = IIf(SS0 - (HdXS - HXS) / DWD < 0, SS0, SS0 - (HdXS - HXS) / DWD)
                DWU = (1 / NfS) * HXS ^ (2 / 3) * SSf ^ 0.5
                FRQoutS = DWD * DWU * HXS
                If RFC0(ii, jj) = 1 Then
                    Pnet(NextNoIJ(ii, jj)) = DRMPO(NextNoIJ(ii, jj)) - DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj)) - DRMIch(NextNoIJ(ii, jj))
                    FRSQinS = SumDRMQinS(NextNoIJ(ii, jj))
                Else
                    Pnet(NextNoIJ(ii, jj)) = DRMPO(NextNoIJ(ii, jj)) + (SumDRMQinS(NextNoIJ(ii, jj)) * DT * 0.001 / Garea) - DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj)) - DRMIch(NextNoIJ(ii, jj))
                    FRSQinS = 0
                End If
                
                Pe = Pnet(NextNoIJ(ii, jj))
                Ek = DRMET(NextNoIJ(ii, jj))
                GWM = DRMWM(ii, jj)
                GSM = DRMSM(ii, jj)
                GW = DRMW0(ii, jj)
                GS = DRMS0(ii, jj)
                GWU = GridWU0(ii, jj)
                GWL = GridWL0(ii, jj)
                GWD = GW - GWU - GWL
                Kg = DRMKg(ii, jj)
                Ki = DRMKi(ii, jj)
                KKi = (1 - ((1 - (Kg + Ki)) ^ ((DT / 3600) / 24))) / (1 + Kg / Ki)
                KKg = KKi * Kg / Ki
                Kg = KKg
                Ki = KKi
                Grfc = RFC0(ii, jj)
                ZUpper = AlUpper * ThickoVZ(ii, jj)
                ZLower = AlLower * ThickoVZ(ii, jj)
                ZDeeper = AlDeeper * ThickoVZ(ii, jj)
                If ZUpper > 300 Then
                    GWUM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * 300 + (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * (ZUpper - 300)
                    GWLM = (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * ZLower
                Else
                    GWUM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * ZUpper
                    If ZUpper + ZLower > 300 Then
                        GWLM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * (300 - ZUpper) + (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * (ZLower - 300 + ZUpper)
                    Else
                        GWLM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * ZLower
                    End If
                End If
                GWDM = GWM - GWUM - GWLM
                If Pe <= 0 Then
                    If GWU + Pe < 0# Then
                        GEU = GWU + Ek + Pe
                        GWU = 0
                        GEL = (Ek - GEU) * GWL / GWLM
                        If GWL < DeeperC * GWLM Then
                            GEL = DeeperC * (Ek - GEU)
                        End If
                        If (GWL - GEL) < 0# Then
                            ged = GEL - GWL
                            GEL = GWL
                            GWL = 0
                            GWD = GWD - ged
                        Else
                            ged = 0
                            GWL = GWL - GEL
                        End If
                    Else
                        GEU = Ek
                        GEL = 0
                        ged = 0
                        GWU = GWU + Pe
                    End If
                    GW = GWU + GWL + GWD
                    If GW < 0# Then
                        GEU = 0
                        GEL = 0
                        ged = DRMW(ii, jj)
                        GE = ged
                        GWU = 0
                        GWL = 0
                        GWD = 0
                        GW = 0
                    End If
                    GE = GEU + GEL + ged
                    GRs = 0
                    GRg = GS * Kg
                    GRi = GS * Ki
                    GS = GS * (1 - Kg - Ki)
                Else
                    GEU = Ek
                    GEL = 0
                    ged = 0
                    GE = Ek
                    nd = Int(Pe / Div) + 1
                    If Pe Mod Div = 0 Then
                        If Pe - Int(Pe) = 0 Then
                            nd = nd - 1
                        End If
                    End If
                    ReDim PPe(1 To nd)
                    For m = 1 To nd - 1
                        PPe(m) = Div
                    Next m
                    PPe(nd) = Pe - (nd - 1) * Div
                    GRs = 0
                    GRg = 0
                    GRi = 0
                    KKi = (1# - (1# - (Kg + Ki)) ^ (1# / nd)) / (Kg + Ki)
                    KKg = KKi * Kg
                    KKi = KKi * Ki
                    For m = 1 To nd
                        If PPe(m) + GW < GWM Then
                            If PPe(m) + GWU < GWUM Then
                                GWU = PPe(m) + GWU
                            ElseIf GWL + GWUM - GWU + PPe(m) < GWLM Then
                                GWL = GWL + GWUM - GWU + PPe(m)
                                GWU = GWUM
                            Else
                                GWD = GWD + PPe(m) - (GWUM - GWU) - (GWLM - GWL)
                                GWU = GWUM
                                GWL = GWLM
                            End If
                            GW = GWU + GWL + GWD
                            GRs = GRs
                            GRg = GRg + GS * KKg
                            GRi = GRi + GS * KKi
                            GS = GS * (1 - KKg - KKi)
                        Else
                            GR = PPe(m) + GW - GWM
                            GWU = GWUM
                            GWL = GWLM
                            GWD = GWDM
                            GW = GWM
                            If GR + GS <= GSM Then
                                GRs = GRs
                                GS = GS + GR
                                GRg = GRg + GS * KKg
                                GRi = GRi + GS * KKi
                                GS = GS * (1 - KKg - KKi)
                            Else
                                GRs = GRs + GR + GS - GSM
                                GS = GSM
                                GRg = GRg + GSM * KKg
                                GRi = GRi + GSM * KKi
                                GS = GSM * (1 - KKg - KKi)
                            End If
                        End If
                    Next m
                End If
                Qs = GRs * Garea * 1000 / DT
                Qi = GRi * Garea * 1000 / DT
                Qg = GRg * Garea * 1000 / DT
                
                HS = DRMHS(0, NextNoIJ(ii, jj))
                HXS = DRMHS(1, NextNoIJ(ii, jj))
                HS = 0.5 * (HS + HXS + (FRSQinS + Qs - FRQoutS) * 0.001 * DT / Garea)
                HS = IIf(HS > 0, HS, 0)
                DRMHS(1, NextNoIJ(ii, jj)) = HS
                If HS = 0 And FRSQinS + Qs < FRQoutS Then
                    FRQoutS = FRSQinS + Qs
                    SumDRMQinS(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinS(NextNoIJ(i1, j1)) + FRQoutS * (1 - DRMfc(ii, jj)), SumDRMQinS(NextNoIJ(ii, jj)))
                    SumDRMQinSCh(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinSCh(NextNoIJ(i1, j1)) + FRQoutS * DRMfc(ii, jj), SumDRMQinSCh(NextNoIJ(ii, jj)))
                Else
                    SumDRMQinS(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinS(NextNoIJ(i1, j1)) + FRQoutS * (1 - DRMfc(ii, jj)), SumDRMQinS(NextNoIJ(ii, jj)))
                    SumDRMQinSCh(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinSCh(NextNoIJ(i1, j1)) + FRQoutS * DRMfc(ii, jj), SumDRMQinSCh(NextNoIJ(ii, jj)))
                End If
                
                NfC = DRMMNfC(ii, jj)
                SS0 = DRMCSlope(ii, jj)
                HXCh = DRMHCh(1, NextNoIJ(ii, jj))
                HdXCh = DRMHCh(1, NextNoIJ(i1, j1))
                SSf = IIf(SS0 - (HdXCh - HXCh) / DWD < 0, SS0, SS0 - (HdXCh - HXCh) / DWD)
                CSAlpha = DRMAlpha(ii, jj)
                AXCh = (CSAlpha / (1 + CSBeta)) * HXCh ^ (1 + CSBeta)
                BXCh = CSAlpha * HXCh ^ CSBeta
                If BXCh > 0 Then
                    FRQoutCh = (1 / NfC) * SSf ^ 0.5 * AXCh * (AXCh / BXCh) ^ (2 / 3)
                    FRQoutCh = (FRQoutCh + DRMQoutCh(NextNoIJ(ii, jj))) / 2
                Else
                    FRQoutCh = 0
                    FRQoutCh = (FRQoutCh + DRMQoutCh(NextNoIJ(ii, jj))) / 2
                End If
                FRSQinCh = Qi + Qg + (DRMIch(NextNoIJ(ii, jj)) * DWD * DRMBmax(ii, jj) * 0.001 / DT) + SumDRMQinSCh(NextNoIJ(ii, jj)) + SumDRMQinCh(NextNoIJ(ii, jj))
                HCh = DRMHCh(0, NextNoIJ(ii, jj))
                ACH = (CSAlpha / (1 + CSBeta)) * HCh ^ (1 + CSBeta)
                ACH = 0.5 * (ACH + AXCh + (DT / DWD) * (FRSQinCh - FRQoutCh))
                ACH = IIf(ACH > 0, ACH, 0)
                HCh = ((1 + CSBeta) * ACH / CSAlpha) ^ (1 / (1 + CSBeta))
                DRMHCh(1, NextNoIJ(ii, jj)) = HCh
                If ACH = 0 And FRSQinCh < FRQoutCh Then
                    FRQoutCh = FRSQinCh
                    SumDRMQinCh(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinCh(NextNoIJ(i1, j1)) + FRQoutCh, SumDRMQinCh(NextNoIJ(ii, jj)))
                Else
                    SumDRMQinCh(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinCh(NextNoIJ(i1, j1)) + FRQoutCh, SumDRMQinCh(NextNoIJ(ii, jj)))
                End If
                
            Next k
            ReDim SumDRMQinS(1 To DRMGCNo), SumDRMQinCh(1 To DRMGCNo), SumDRMQinSCh(1 To DRMGCNo)
            For k = 1 To DRMGCNo
                ii = SortingRow(k)
                jj = SortingCol(k)
                Select Case FlowDirection(ii, jj)
                    Case 0
                        i1 = ii - 1
                        j1 = jj + 1
                        DWD = 2 ^ 0.5 * Cellsize
                    Case 1
                        i1 = ii
                        j1 = jj + 1
                        DWD = Cellsize
                    Case 2
                        i1 = ii + 1
                        j1 = jj + 1
                        DWD = 2 ^ 0.5 * Cellsize
                    Case 3
                        i1 = ii + 1
                        j1 = jj
                        DWD = Cellsize
                    Case 4
                        i1 = ii + 1
                        j1 = jj - 1
                        DWD = 2 ^ 0.5 * Cellsize
                    Case 5
                        i1 = ii
                        j1 = jj - 1
                        DWD = Cellsize
                    Case 6
                        i1 = ii - 1
                        j1 = jj - 1
                        DWD = 2 ^ 0.5 * Cellsize
                    Case 7
                        i1 = ii - 1
                        j1 = jj
                        DWD = Cellsize
                    Case 8
                        i1 = ii
                        j1 = jj
                        DWD = Cellsize
                    Case Else
                End Select
                SS0 = DRMSSlope(ii, jj)
                NfS = DRMMNfS(ii, jj)
                HS = DRMHS(1, NextNoIJ(ii, jj)) * 0.001
                HdS = DRMHS(1, NextNoIJ(i1, j1)) * 0.001
                SSf = IIf(SS0 - (HdS - HS) / DWD < 0, SS0, SS0 - (HdS - HS) / DWD)
                DWU = (1 / NfS) * HS ^ (2 / 3) * SSf ^ 0.5
                FRQoutS = DWD * DWU * HS
                DRMQoutS(NextNoIJ(ii, jj)) = FRQoutS
                SumDRMQinS(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinS(NextNoIJ(i1, j1)) + FRQoutS * (1 - DRMfc(ii, jj)), SumDRMQinS(NextNoIJ(ii, jj)))
                SumDRMQinSCh(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinSCh(NextNoIJ(i1, j1)) + FRQoutS * DRMfc(ii, jj), SumDRMQinSCh(NextNoIJ(ii, jj)))
                NfC = DRMMNfC(ii, jj)
                SS0 = DRMCSlope(ii, jj)
                HCh = DRMHCh(1, NextNoIJ(ii, jj))
                HdCh = DRMHCh(1, NextNoIJ(i1, j1))
                SSf = IIf(SS0 - (HdCh - HCh) / DWD < 0, SS0, SS0 - (HdCh - HCh) / DWD)
                CSAlpha = DRMAlpha(ii, jj)
                ACH = (CSAlpha / (1 + CSBeta)) * HCh ^ (1 + CSBeta)
                BCh = CSAlpha * HCh ^ CSBeta
                If BCh > 0 Then
                    FRQoutCh = (1 / NfC) * SSf ^ 0.5 * ACH * (ACH / BCh) ^ (2 / 3)
                Else
                    FRQoutCh = 0
                End If
                DRMQoutCh(NextNoIJ(ii, jj)) = FRQoutCh
                SumDRMQinCh(NextNoIJ(i1, j1)) = IIf(k < DRMGCNo, SumDRMQinCh(NextNoIJ(i1, j1)) + FRQoutCh, SumDRMQinCh(NextNoIJ(ii, jj)))
                DRMHS(0, NextNoIJ(ii, jj)) = DRMHS(1, NextNoIJ(ii, jj))
                DRMHCh(0, NextNoIJ(ii, jj)) = DRMHCh(1, NextNoIJ(ii, jj))
                If BCh > 0 Then
                    GridVch(ii, jj) = (1 / NfC) * SSf ^ 0.5 * (ACH / BCh) ^ (2 / 3)
                Else
                    GridVch(ii, jj) = 0
                End If
                GridVs(ii, jj) = DWU
                GridQ(ii, jj) = FRQoutS + FRQoutCh
                If ((DT * (DWU + (9.81 * HS) ^ 0.5)) / DWD) > 1 Then
                    MsgBox "第" & ENo0 & "次洪水差分格式不稳定，程序中断！"
                    Exit Sub
                End If
                            
            Next k
            
            ii = SortingRow(DRMGCNo)
            jj = SortingCol(DRMGCNo)
            SimQ(i + 1) = SumDRMQinS(NextNoIJ(ii, jj)) + SumDRMQinSCh(NextNoIJ(ii, jj)) + SumDRMQinCh(NextNoIJ(ii, jj))
            SimH(i + 1) = DRMHCh(1, NextNoIJ(ii, jj))
            If NUPoints > 0 Then
                For m = 1 To NUPoints
                    ii = UPRow(m)
                    jj = UPCol(m)
                    UPSimQ(i + 1, m) = SumDRMQinS(NextNoIJ(ii, jj)) + SumDRMQinCh(NextNoIJ(ii, jj)) + SumDRMQinSCh(NextNoIJ(ii, jj))
                    UPSimH(i + 1, m) = DRMHCh(1, NextNoIJ(ii, jj))
                Next m
            End If
        Next i
        
        TSteps = DateDiff("H", EStarttime(ENo), EEndtime(ENo)) / DatumTIs + 1
        For j = 1 To TSteps
            QSim(j) = SimQ((j - 1) * Int(DatumTIs * 3600 / DT) + 1)
            HSim(j) = SimH((j - 1) * Int(DatumTIs * 3600 / DT) + 1)
        Next j
        If NUPoints > 0 Then
            On Error GoTo AccessError
            ConnectSys.Execute "delete * from [UpHResults-" & SSName & "] where [时间] between  #" & Format(EStarttime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and #" & Format(EEndtime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and [目的]= '" & Aimat & " '"
            Rd.Open "select * from [UpHResults-" & SSName & "] ", ConnectSys, adOpenDynamic, adLockOptimistic
            For i = 1 To TSteps
                Rd.AddNew
                Rd("洪号") = ENo
                Rd("目的") = Aimat
                Rd("时间") = TimeSeries(i)
                For j = 1 To NUPoints
                    UPQSim(i, j) = UPSimQ((i - 1) * Int(DatumTIs * 3600 / DT) + 1, j)
                    Rd(UPName(j)) = UPQSim(i, j)
                Next j
                Rd.Update
            Next i
            Rd.Close
            GoTo NoAccessError
AccessError:
            MsgBox "请在表UpHResultsMK-" & SSName & "中添加上游入流点[" & UPName(j) & "]字段！"
            Rd.Update
            Rd.Close
        End If
NoAccessError:
        ConnectSys.Execute "delete * from [HResults-" & SSName & "] where [Purpose]='" & Aimat & "' and [Time] between  #" & Format(EStarttime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and #" & Format(EEndtime(ENo), "YYYY-MM-DD HH:MM:SS") & "#"
        ConnectSys.Execute "delete * from [HCResults-" & SSName & "] where [Start Time] = #" & Format(EStarttime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and [Purpose]= '" & Aimat & " '"
        With Rd
            SumOQ = 0
            SumSQ = 0
            OPeak = 0
            SPeak = 0
            SumPre = 0
            SumEE = 0
            ONC = 0
            SNC = 0

            For i = 1 To TSteps
                SumOQ = SumOQ + QObs(i)
                SumSQ = SumSQ + QSim(i)
                If OPeak < QObs(i) Then
                    OPeak = QObs(i)
                    OPeakTime = i
                End If
                If SPeak < QSim(i) Then
                    SPeak = QSim(i)
                    SPeakTime = i
                End If
                SumPre = SumPre + AvgP(i)
            Next i

            AvgOQ = SumOQ / TSteps
            .Open "select * from [HResults-" & SSName & "] ", ConnectSys, adOpenDynamic, adLockOptimistic
            For i = 1 To TSteps
                ONC = ONC + (QObs(i) - AvgOQ) ^ 2
                SNC = SNC + (QSim(i) - QObs(i)) ^ 2
                .AddNew
                Rd("FloodNo") = ENo0
                Rd("Purpose") = Aimat
                Rd("Time") = TimeSeries(i)
                Rd("SimulatedQ") = QSim(i)
                Rd("ObservedQ") = QObs(i)
                Rd("AverageP") = AvgP(i)
                .Update
            Next i
            .Close

            ORunoff = SumOQ * 3.6 * DatumTIs / DArea
            SRunoff = SumSQ * 3.6 * DatumTIs / DArea

            .Open "select * from [HCResults-" & SSName & "] ", ConnectSys, adOpenDynamic, adLockOptimistic
                .AddNew
                Rd("FloodNo") = ENo0
                Rd("Purpose") = Aimat
                Rd("Start Time") = TimeSeries(1)
                Rd("Precipitation") = SumPre
                Rd("ObservedRO") = ORunoff
                Rd("SimulatedRO") = SRunoff
                Rd("RO Error(%)") = (SRunoff - ORunoff) / ORunoff * 100
                Rd("ObservedPeak") = OPeak
                Rd("SimulatedPeak") = SPeak
                Rd("Peak Error(%)") = (SPeak - OPeak) / OPeak * 100
                Rd("Time Error") = SPeakTime - OPeakTime
                Rd("NC") = 1 - SNC / ONC
'                If Abs(Rd("Peak Error(%)")) > 20 Then
'                    MPeakE = 0
'                Else
'                    MPeakE = 2.5 * (1 - (Abs(Rd("Peak Error(%)")) / 20))
'                End If
'                If Abs(Rd("RO Error(%)")) > 20 Then
'                    MRunoffE = 0
'                Else
'                    MRunoffE = 2.5 * (1 - (Abs(Rd("RO Error(%)")) / 20))
'                End If
'                If Abs(Rd("Time Error")) > 3 Then
'                    MTimeE = 0
'                Else
'                    MTimeE = 2.5 * (1 - (Abs(Rd("Time Error")) / 4))
'                End If
'                If Rd("NC") < 0.5 Then
'                    MNashC = 0
'                Else
'                    MNashC = 2.5 * Rd("NC")
'                End If
'                Rd("Grading-marks") = MPeakE + MRunoffE + MTimeE + MNashC
                .Update
            .Close
        End With
        Erase DRMPO, DRMET, DRMQoutS, DRMHS
        Erase SumDRMQinS, SumDRMQinCh, SumDRMQinSCh
        Erase Dis, Pcum, Icum, DRMIca
        Erase DRMIch, QSim, RFC, Pnet
        Erase DRMW, DRMS, DRMQoutCh, DRMHCh
        Erase DRMW0, DRMS0, RFC0
    Next ENo
    HourlyScale = True
    Exit Sub
NReRun:
    MsgBox "请输入正确的数据信息，程序中断！"
End Sub

Sub ModelRunMusk(Aimat As String)
    Dim Rd As New ADODB.Recordset, CheckFile As String
    Dim SSName As String
    Dim EStarttime() As Date, EEndtime() As Date
    Dim EStartday() As Date, EEndday() As Date, StartDay As Date, EndDay As Date
    Dim ETime() As Date, LongFDTime As Long
    Dim DTS As Integer, GGE() As Single, HTS As Integer, GEE() As Single
    Dim DRMOrderNo() As Integer
    Dim DRMWM() As Single
    Dim DRMSM() As Single
    Dim DRMMNfS() As Single, NfS As Single
    Dim DRMMNfC() As Single, NfC As Single
    Dim DRMAlpha() As Single, DRMBmax() As Single
    Dim NoEvents  As Integer
    Dim PSCol() As Integer, PSRow() As Integer
    Dim PObs() As Single, QObs() As Single, EObs() As Single, AvgP() As Single
    Dim NextNoIJ() As Long, DRMGCNo As Long
    Dim DRMSSlope() As Single, DRMCSlope() As Single, DRMfc() As Single
    Dim TSteps As Long, TimeSeries() As Date
    Dim DRMET() As Single, DRMPO() As Single
    Dim LAI() As Single
    Dim DT As Single
    Dim SortingOrder() As Long, SortingRow() As Integer, SortingCol() As Integer
    Dim SType030() As Integer, SType30100() As Integer, LCover() As Integer
    Dim DRMIca() As Single, SumDRMIca As Single
    Dim DRMIch() As Single
    Dim Cellsize As Integer
    Dim GridW() As Single, GridS() As Single
    Dim OC As Single, ROC As Single
    Dim DRMKg() As Single, DRMKi() As Single
    Dim STSWC() As Single, STFC() As Single, STWP() As Single
    Dim CSBeta As Single, CSAlpha As Single
    Dim QSim() As Single, SimQ() As Single
    Dim i As Long, j As Integer, ii As Integer, jj As Integer, k As Long
    Dim i1 As Integer, j1 As Integer, m As Integer
    Dim DEMPrecision As Integer, ENo As Integer
    Dim IW As Single
    Dim Dis() As Single, Sumdis As Single, Sump As Single
    Dim SaP As Boolean, Kp As Integer
    Dim GLAI As Single, Scmax As Single, Cp As Single, Cvd As Single, Pcum() As Single
    Dim Icum() As Single
    Dim RFC() As Integer, Grfc As Integer
    Dim Pnet() As Single
    Dim GW As Single, GS As Single, GWM As Single, GSM As Single
    Dim HumousT() As Single, ThickoVZ() As Single
    Dim ThitaS As Single, ThitaF As Single, ThitaW As Single, Thita As Single
    Dim SumKgKi As Single, KgKi As Single
    Dim Kg As Single, Ki As Single, KKg As Single, KKi As Single
    Dim Pe As Single, Cg As Single, Ci As Single, Ct As Single, CCg As Single, CCci As Single
    Dim R As Single, Rs As Single, Ri As Single, Rg As Single, GQs As Single, GQi As Single, GQg As Single, Gqch As Single
    Dim DWD As Single
    Dim SS0 As Single, SSf As Single
    Dim DWU As Single
    Dim AlUpper As Single, AlLower As Single, AlDeeper As Single
    Dim ZUpper As Single, ZLower As Single, ZDeeper As Single
    Dim GWUM As Single, GWU As Single, GWLM As Single, GWL As Single, GWDM As Single, GWD As Single
    Dim GridWU() As Single, GridWL() As Single, GridWD() As Single
    Dim GridWU0() As Single, GridWL0() As Single, GridWD0() As Single
    Dim GE As Single, GEU As Single, GEL As Single, ged As Single
    Dim Ek As Single, KEpC As Single, DeeperC As Single, Div As Integer
    Dim GridQi() As Single, GridQg() As Single, GridQs() As Single, GGEE() As Single
    Dim GridFLC() As Single, Dp As Single
    Dim JYear As Integer, JMonth As Integer, JDay As Integer, nd As Integer
    Dim GR As Single, GRs As Single, GRi As Single, GRg As Single
    Dim MC1 As Single, MC2 As Single, MC3 As Single, MDt As Single
    Dim CCS As Single, LagTime As Integer
    Dim MKch As Single, MXch As Single, MKs As Single, MXs As Single, MKi As Single, MXi As Single, MKg As Single, MXg As Single
    Dim InflowQs() As Single, OutflowQs() As Single, InflowQi() As Single, OutflowQi() As Single
    Dim InflowQg() As Single, OutflowQg() As Single, InflowQch() As Single
    Dim SumQs() As Single, SumQi() As Single, SumQg() As Single, SumQch() As Single

    '结果统计
    Dim SumOQ As Single, SumSQ As Single, OPeak As Single, SPeak As Single, SumEE As Single, ET0out() As Single
    Dim OPeakTime As Integer, SPeakTime As Integer, SumPre As Single, DArea As Single
    Dim ORunoff As Single, SRunoff As Single, ONC As Single, SNC As Single, AvgOQ As Single
    Dim Qoutch As Single, Qouts As Single, Qouti As Single, Qoutg As Single, CCS1 As Single
'    Dim MPeakE As Single, MRunoffE As Single, MTimeE As Single, MNashC As Single
    Dim UPRow() As Integer, UPCol() As Integer, UPName() As String, UPSimQ() As Single, UPQSim() As Single
    Dim UPSimH() As Single, UPHSim() As Single
    Dim GridWM() As Single, GridSM() As Single, NoWM() As Single, NoSM() As Single
    Dim kk1 As Integer, kk2 As Integer
    
    Dim x As Integer, y As Integer
    Dim gridRg() As Single, gridQg2() As Single, baohe() As Integer
    ''''subgrid variability of runoff generation'''''''
    Dim aa As Single, GWMM As Single, b As Single, R1 As Single, R2 As Single, peds As Single
     ''''subgrid variability of runoff division'''''''
    Dim xx As Single, fr As Single  ''''initial value needed
    Dim au As Single, ff As Single, ex As Single, GSMM As Single, R3 As Single
    Dim gridfr() As Single
    '''''infiltration excess'''''''''''''''
    Dim STMIR() As Single, STSHC() As Single
    Dim fcd() As Single, fmd() As Single, fxd() As Single
    Dim fc() As Single, fm() As Single
    Dim GRe As Single, fd As Single, fdmax As Single, ex2 As Single, R4 As Single
    ''''infiltration excess2''''''''
    Dim uu As Single, kim As Single
    Dim iff As Single, ffc As Single, sfx As Single
    With Rd
        .Open "select * from [WholeCatchPara] where YesOrNo= Yes", ConnectSys, adOpenStatic, adLockReadOnly
        If IsNull(Rd("Shortening")) Or IsNull(Rd("Time Interval(h)")) Or IsNull(Rd("Time Interval(h)")) Then
            MsgBox "请核实研究流域相关参数，程序中断！"
            .Close
            Exit Sub
        End If
        If IsNull(Rd("Ratio of OC")) Or IsNull(Rd("Outflow Coefficients")) Or IsNull(Rd("CS-HM")) Or IsNull(Rd("Lag Time-HM")) Then
            MsgBox "请核实研究流域相关参数，程序中断！"
            .Close
            Exit Sub
        End If
        If IsNull(Rd("K")) Or IsNull(Rd("C")) Or IsNull(Rd("LUM")) Or IsNull(Rd("LLM")) Or IsNull(Rd("CG")) Or IsNull(Rd("CI")) Then
            MsgBox "请核实研究流域相关参数，程序中断！"
            .Close
            Exit Sub
        End If
        SSName = Rd("Shortening")
        DT = Rd("Time Interval(h)")
        OC = Rd("Outflow Coefficients")
        ROC = Rd("Ratio of OC")
        CSBeta = Rd("Beta")
        DeeperC = Rd("C")
        KEpC = Rd("K")
        AlUpper = Rd("LUM")
        AlLower = Rd("LLM")
        AlDeeper = 1 - AlUpper - AlLower
        CCS = Rd("CS-HM")
        LagTime = Rd("Lag Time-HM")
        CCg = Rd("CG")
        CCci = Rd("CI")
        MKch = Rd("Kech")
        MXch = Rd("Xech")
        MKs = Rd("Kes")
        MXs = Rd("Xes")
        MKi = Rd("Kei")
        MXi = Rd("Xei")
        MKg = Rd("Keg")
        MXg = Rd("Xeg")
        b = Rd("b")
        ex = Rd("ex")
        ex2 = Rd("ex2")
        uu = Rd("uu")
        kim = Rd("kim")
        .Close

        .Open "select * from [HFlood Events-" & SSName & "] where [Purpose]='" & Aimat & "' order by [FloodNo]", ConnectSys, adOpenStatic, adLockReadOnly
        i = 0
        If IsNull(Rd("Purpose")) Or IsNull(Rd("Start time")) Or IsNull(Rd("End time")) Then
            MsgBox "请核实率定洪水的相关参数，程序中断！"
            .Close
            Exit Sub
        End If
        NoEvents = .RecordCount
        If NoEvents = 0 Then
            MsgBox "数据表中没有率定的洪水，程序中断！"
            .Close
            Exit Sub
        End If
        .MoveFirst
        Do
            i = i + 1
            ReDim Preserve EStarttime(1 To i), EEndtime(1 To i), EStartday(1 To i), EEndday(1 To i)
            EStarttime(i) = Rd("Start time")
            EEndtime(i) = Rd("End time")
            EStartday(i) = Format(EStarttime(i), "YYYY-MM-DD")
            EEndday(i) = Format(EEndtime(i), "YYYY-MM-DD")
            .MoveNext
        Loop Until i = NoEvents
        .Close

        If NPStation = 0 Then
            MsgBox "P-Station缺乏站点信息，程序中断！"
            Exit Sub
        End If
        DEMPrecision = DDem * 3600
        .Open "select * from [P-Station] where [Watersheds]='" & StationName & "' order by [SubbasinNo]", ConnectSys, adOpenStatic, adLockReadOnly
        .MoveFirst
        ReDim SubbasinName(1 To NPStation), StrLon(1 To NPStation), StrLat(1 To NPStation), PSCol(1 To NPStation), PSRow(1 To NPStation)
        For i = 1 To NPStation
            SubbasinName(i) = Rd("PStationName")
            StrLat(i) = Rd("Latitude")
            StrLon(i) = Rd("Longitude")
            PSCol(i) = Int(((StrLon(i) - XllCorner) * 60 * (60 / DEMPrecision))) + 1
            PSRow(i) = Nx - Int(((StrLat(i) - YllCorner) * 60 * (60 / DEMPrecision)))
            .MoveNext
        Next i
        .Close

'        .Open "select * from [StationXY] where [Watersheds]='" & StationName & "' order by [SubbasinNo]", ConnectSys, adOpenStatic, adLockReadOnly
'        ReDim PSCol(1 To NPStation), PSRow(1 To NPStation)
'        .MoveFirst
'        For i = 1 To NPStation
'            PSCol(i) = Rd("Col")
'            PSRow(i) = Rd("Row")
'            .MoveNext
'        Next i
'        .Close

        If NUPoints > 0 Then
            ReDim UPRow(1 To NUPoints), UPCol(1 To NUPoints), UPName(1 To NUPoints)
            DEMPrecision = DDem * 3600
            .Open "select * from [Upstream outlets] where Watersheds='" & StationName & "'order by Points", ConnectSys, adOpenStatic, adLockReadOnly
            If .RecordCount <> NUPoints Then
                MsgBox "表Upstream outlets中上游出口点个数与表WholeCatchPara中Upstream points值不一致，请核实！"
                .Close
                Exit Sub
            End If
            .MoveFirst
            For i = 1 To NUPoints
                UPName(i) = Rd("Name")
                StrLat(i) = Rd("Latitude")
                StrLon(i) = Rd("Longitude")
                UPRow(i) = Nx - Int(((StrLat(i) - YllCorner) * 60 * (60 / DEMPrecision)))
                UPCol(i) = Int(((StrLon(i) - XllCorner) * 60 * (60 / DEMPrecision))) + 1
                .MoveNext
            Next i
            .Close
        End If
    End With

    ReDim STSWC(0 To 12), STFC(0 To 12), STWP(0 To 12), STSHC(0 To 12), STMIR(0 To 12)
    Rd.Open "select * from  [Soil Types] order by Category", ConnectSys, adOpenStatic, adLockReadOnly
    If Rd.RecordCount = 0 Then
        MsgBox "请在Soil Types表中输入土壤类型相关参数！"
        Rd.Close
        Exit Sub
    End If
    Rd.MoveFirst
    i = 0
    Do
        STSWC(i) = Rd("SWC")
        STFC(i) = Rd("FC")
        STWP(i) = Rd("WP")
        STSHC(i) = Rd("SHC (cm/h)")           'SHC为稳定下渗率'
        STMIR(i) = Rd("MIR(cm/h)")            'MIR为最大下渗率'
        i = i + 1
        Rd.MoveNext
    Loop Until Rd.EOF
    Rd.Close

    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "DEM.asc")
    If CheckFile = "" Then
        MsgBox "缺少" & StationName & "流域DEM高程ASC文件！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "累积汇水面积.asc")
    If CheckFile = "" Then
        MsgBox "缺少" & StationName & "流域累积汇水面积ASC文件！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "水系.asc")
    If CheckFile = "" Then
        MsgBox "缺少" & StationName & "流域水系ASC文件！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "栅格流向.asc")
    If CheckFile = "" Then
        MsgBox "缺少" & StationName & "流域栅格流向ASC文件！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "栅格演算次序.asc")
    If CheckFile = "" Then
        MsgBox "请先进行栅格间汇流演算次序计算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "张力水蓄水容量.asc")
    If CheckFile = "" Then
        MsgBox "请先进行包气带厚度估算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "自由水蓄水容量.asc")
    If CheckFile = "" Then
        MsgBox "请先进行包气带厚度估算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "河道汇流糙率.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "河道宽度指数.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "坡面汇流糙率.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "径流分配比例.asc")
    If CheckFile = "" Then
        MsgBox "请先进行栅格间汇流演算次序计算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Input\" & StationName & "\" & StationName & "最陡坡度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行最陡坡度提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "河道最大宽度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\input\" & StationName & "\" & StationName & "植被类型.asc")
    If CheckFile = "" Then
        MsgBox "请先给定植被类型数据！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\input\" & StationName & "\" & StationName & "0-30cm土壤类型.asc")
    If CheckFile = "" Then
        MsgBox "请先给定0-30cm土壤类型数据！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\input\" & StationName & "\" & StationName & "30-100cm土壤类型.asc")
    If CheckFile = "" Then
        MsgBox "请先给定30-100cm土壤类型数据！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "腐殖质土厚度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行包气带厚度估算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "栅格河道坡度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行糙率及宽度指数提取！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If
    CheckFile = Dir(App.Path & "\Output\" & StationName & "\" & StationName & "包气带厚度.asc")
    If CheckFile = "" Then
        MsgBox "请先进行包气带厚度估算！程序中断！", vbExclamation + vbInformation, "警告："
        Exit Sub
    End If

    ReDim E(1 To Nx, 1 To Ny), WaterArea(1 To Nx, 1 To Ny), RiverPoint(1 To Nx, 1 To Ny), FlowDirection(1 To Nx, 1 To Ny)
    ReDim DRMOrderNo(1 To Nx, 1 To Ny), DRMWM(1 To Nx, 1 To Ny), DRMSM(1 To Nx, 1 To Ny)
    ReDim DRMMNfC(1 To Nx, 1 To Ny), DRMAlpha(1 To Nx, 1 To Ny), DRMMNfS(1 To Nx, 1 To Ny)
    ReDim NextNoIJ(1 To Nx, 1 To Ny), DRMBmax(1 To Nx, 1 To Ny)
    ReDim DRMSSlope(1 To Nx, 1 To Ny), DRMCSlope(1 To Nx, 1 To Ny), DRMfc(1 To Nx, 1 To Ny)
    ReDim SType030(1 To Nx, 1 To Ny), SType30100(1 To Nx, 1 To Ny), TopIndex(1 To Nx, 1 To Ny), LCover(1 To Nx, 1 To Ny)
    ReDim DRMKg(1 To Nx, 1 To Ny), DRMKi(1 To Nx, 1 To Ny), HumousT(1 To Nx, 1 To Ny), ThickoVZ(1 To Nx, 1 To Ny)
    ReDim GridWM(1 To Nx, 1 To Ny), GridSM(1 To Nx, 1 To Ny)
    ReDim gridRg(1 To Nx, 1 To Ny), gridQg2(1 To Nx, 1 To Ny), baohe(1 To Nx, 1 To Ny)
    ReDim gridfr(1 To Nx, 1 To Ny)
    ReDim fc(1 To Nx, 1 To Ny), fm(1 To Nx, 1 To Ny)
    
    Open App.Path & "\Input\" & StationName & "\" & StationName & "DEM.asc" For Input As #1
    Input #1, Str
    Input #1, Str
    Input #1, Str
    Input #1, Str
    Input #1, Str
    Input #1, Str
    Open App.Path & "\Input\" & StationName & "\" & StationName & "累积汇水面积.asc" For Input As #2
    Input #2, Str
    Input #2, Str
    Input #2, Str
    Input #2, Str
    Input #2, Str
    Input #2, Str
    Open App.Path & "\Input\" & StationName & "\" & StationName & "水系.asc" For Input As #3
    Input #3, Str
    Input #3, Str
    Input #3, Str
    Input #3, Str
    Input #3, Str
    Input #3, Str
    Open App.Path & "\Input\" & StationName & "\" & StationName & "栅格流向.asc" For Input As #4
    Input #4, Str
    Input #4, Str
    Input #4, Str
    Input #4, Str
    Input #4, Str
    Input #4, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "栅格演算次序.asc" For Input As #5
    Input #5, Str
    Input #5, Str
    Input #5, Str
    Input #5, Str
    Input #5, Str
    Input #5, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "张力水蓄水容量.asc" For Input As #6
    Input #6, Str
    Input #6, Str
    Input #6, Str
    Input #6, Str
    Input #6, Str
    Input #6, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "自由水蓄水容量.asc" For Input As #7
    Input #7, Str
    Input #7, Str
    Input #7, Str
    Input #7, Str
    Input #7, Str
    Input #7, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "河道汇流糙率.asc" For Input As #8
    Input #8, Str
    Input #8, Str
    Input #8, Str
    Input #8, Str
    Input #8, Str
    Input #8, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "河道宽度指数.asc" For Input As #9
    Input #9, Str
    Input #9, Str
    Input #9, Str
    Input #9, Str
    Input #9, Str
    Input #9, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "坡面汇流糙率.asc" For Input As #10
    Input #10, Str
    Input #10, Str
    Input #10, Str
    Input #10, Str
    Input #10, Str
    Input #10, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "径流分配比例.asc" For Input As #11
    Input #11, Str
    Input #11, Str
    Input #11, Str
    Input #11, Str
    Input #11, Str
    Input #11, Str
    Open App.Path & "\Input\" & StationName & "\" & StationName & "最陡坡度.asc" For Input As #12
    Input #12, Str
    Input #12, Str
    Input #12, Str
    Input #12, Str
    Input #12, Str
    Input #12, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "河道最大宽度.asc" For Input As #13
    Input #13, Str
    Input #13, Str
    Input #13, Str
    Input #13, Str
    Input #13, Str
    Input #13, Str
    Open App.Path & "\input\" & StationName & "\" & StationName & "0-30cm土壤类型.asc" For Input As #14
    Input #14, Str
    Input #14, Str
    Input #14, Str
    Input #14, Str
    Input #14, Str
    Input #14, Str
    Open App.Path & "\input\" & StationName & "\" & StationName & "30-100cm土壤类型.asc" For Input As #15
    Input #15, Str
    Input #15, Str
    Input #15, Str
    Input #15, Str
    Input #15, Str
    Input #15, Str
    Open App.Path & "\input\" & StationName & "\" & StationName & "植被类型.asc" For Input As #16
    Input #16, Str
    Input #16, Str
    Input #16, Str
    Input #16, Str
    Input #16, Str
    Input #16, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "腐殖质土厚度.asc" For Input As #17
    Input #17, Str
    Input #17, Str
    Input #17, Str
    Input #17, Str
    Input #17, Str
    Input #17, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "栅格河道坡度.asc" For Input As #18
    Input #18, Str
    Input #18, Str
    Input #18, Str
    Input #18, Str
    Input #18, Str
    Input #18, Str
    Open App.Path & "\Output\" & StationName & "\" & StationName & "包气带厚度.asc" For Input As #20
    Input #20, Str
    Input #20, Str
    Input #20, Str
    Input #20, Str
    Input #20, Str
    Input #20, Str
    DRMGCNo = 0
    SumKgKi = 0
    Cellsize = Int(DDem / (3 / 3600) * 90)
    For i = 1 To Nx
        For j = 1 To Ny
            Input #1, E(i, j)
            Input #2, WaterArea(i, j)
            Input #3, RiverPoint(i, j)
            Input #4, FlowDirection(i, j)
            Input #5, DRMOrderNo(i, j)
            Input #6, DRMWM(i, j)
            Input #7, DRMSM(i, j)
            Input #8, DRMMNfC(i, j)
            Input #9, DRMAlpha(i, j)
            Input #10, DRMMNfS(i, j)
            Input #11, DRMfc(i, j)
            Input #12, DRMSSlope(i, j)
            Input #13, DRMBmax(i, j)
            Input #14, SType030(i, j)
            Input #15, SType30100(i, j)
            Input #16, LCover(i, j)
            Input #17, HumousT(i, j)
            Input #18, DRMCSlope(i, j)
            Input #20, ThickoVZ(i, j)
            If E(i, j) <> Nodata Then
                DRMGCNo = DRMGCNo + 1
                NextNoIJ(i, j) = DRMGCNo
                If HumousT(i, j) <= 300 Then
                    ThitaS = STSWC(SType030(i, j))
                    ThitaF = STFC(SType030(i, j))
                    ThitaW = STWP(SType030(i, j))
                Else
                    ThitaS = STSWC(SType030(i, j)) * (300 / HumousT(i, j)) + STSWC(SType30100(i, j)) * (1 - 300 / HumousT(i, j))
                    ThitaF = STFC(SType030(i, j)) * (300 / HumousT(i, j)) + STFC(SType30100(i, j)) * (1 - 300 / HumousT(i, j))
                    ThitaW = STWP(SType030(i, j)) * (300 / HumousT(i, j)) + STWP(SType30100(i, j)) * (1 - 300 / HumousT(i, j))
                End If
                DRMKi(i, j) = ((ThitaF / ThitaS) ^ OC) / (1 + ROC / (1 + 2 * (1 - ThitaW)))
                DRMKg(i, j) = (ThitaF / ThitaS) ^ OC - DRMKi(i, j)
                SumKgKi = SumKgKi + (ThitaF / ThitaS) ^ OC
                fc(i, j) = STSHC(SType030(i, j)) * 10
                fm(i, j) = STMIR(SType030(i, j)) * 10
            End If
        Next j
    Next i
    KgKi = SumKgKi / DRMGCNo
    Close #1
    Close #2
    Close #3
    Close #4
    Close #5
    Close #6
    Close #7
    Close #8
    Close #9
    Close #10
    Close #11
    Close #12
    Close #13
    Close #14
    Close #15
    Close #16
    Close #17
    Close #18
    Close #20
    DArea = Garea * DRMGCNo

    ReDim SortingOrder(1 To DRMGCNo), SortingRow(1 To DRMGCNo), SortingCol(1 To DRMGCNo)
    With Rd
        .Open "select * from [CalSorting] where [Watersheds]='" & StationName & "' order by [CalOrder],[Row],[Col] asc", ConnectSys, adOpenStatic, adLockReadOnly
        If .RecordCount <> DRMGCNo Then
            Rd.Close
            MsgBox "请确认栅格演算次序表！程序中断！"
            Exit Sub
        End If
        .MoveFirst
        i = 1
        Do
            SortingOrder(i) = Rd("CalOrder")
            SortingRow(i) = Rd("Row")
            SortingCol(i) = Rd("Col")
            .MoveNext
            i = i + 1
        Loop Until .EOF
        .Close
    End With
    IW = 2

    For ENo = 1 To NoEvents
        Rd.Open "select * from [HObserved-" & SSName & "] where [时间] between  #" & Format(EStarttime(ENo), "YYYY-MM-DD HH:NN:SS") & "# and #" & Format(EEndtime(ENo), "YYYY-MM-DD HH:NN:SS") & "# order by [时间]", ConnectSys, adOpenStatic, adLockReadOnly
        TSteps = Rd.RecordCount
        If TSteps <> (DateDiff("H", EStarttime(ENo), EEndtime(ENo)) / DT + 1) Then
            MsgBox "第" & ENo & "场洪水资料有误，程序中断！"
            Rd.Close
            Exit Sub
        End If
        ReDim PObs(1 To TSteps, 1 To NPStation), EObs(1 To TSteps), TimeSeries(1 To TSteps), AvgP(1 To TSteps)
        ReDim QObs(1 To TSteps), QSim(0 To TSteps + LagTime), HSim(1 To TSteps), GEE(1 To TSteps)
        If NUPoints > 0 Then
            ReDim UPQSim(1 To TSteps, 1 To NUPoints), UPHSim(1 To TSteps, 1 To NUPoints)
        End If
        ReDim NoWM(1 To TSteps), NoSM(1 To TSteps)
        Rd.MoveFirst
        For i = 1 To TSteps
            NoWM(i) = 0
            NoSM(i) = 0
            For j = 1 To NPStation
                If IsNull(Rd(SubbasinName(j))) Then
                    PObs(i, j) = 0
                    AvgP(i) = AvgP(i)
                Else
                    PObs(i, j) = Rd(SubbasinName(j))
                    AvgP(i) = AvgP(i) + Rd(SubbasinName(j))
                End If
            Next j
            AvgP(i) = AvgP(i) / NPStation
            TimeSeries(i) = Rd("时间")
            If IsNull(Rd("实测流量")) Then
                QObs(i) = 0
            Else
                QObs(i) = Rd("实测流量")
            End If
            Rd.MoveNext
        Next i
        Rd.Close

        StartDay = DateAdd("D", -1, EStartday(ENo))
        EndDay = DateAdd("D", 1, EEndday(ENo))
        DTS = DateDiff("D", StartDay, EndDay) + 1
        Rd.Open "select * from [DObserved-" & SSName & "] where [时间] between  #" & Format(StartDay, "YYYY-MM-DD") & "# and #" & Format(EndDay, "YYYY-MM-DD") & "# order by [时间]", ConnectSys, adOpenStatic, adLockReadOnly
        If Rd.RecordCount = 0 Then
            MsgBox "第" & ENo & "场洪水资料缺失，程序中断！"
            Rd.Close
            Exit Sub
        End If
        ReDim GGEE(1 To DTS)
        Rd.MoveFirst
        For i = 1 To DTS
            If IsNull(Rd("蒸发")) Then
                GGEE(i) = 0
            Else
                GGEE(i) = Rd("蒸发")
            End If
            Rd.MoveNext
        Next i
        HTS = (DTS - 1) * 24 + 8
        ReDim GGE(1 To HTS)
        For i = 1 To 8
            GGE(i) = GGEE(1) / 24
        Next i
        For i = 1 To DTS - 1
            For k = 1 To 24
                GGE((i - 1) * 24 + 8 + k) = GGEE(i + 1) / 24
            Next k
        Next i
        HTS = DateDiff("H", StartDay, EStarttime(ENo)) - 1
        For i = 1 To TSteps
            EObs(i) = GGE(HTS + i)
        Next i
        Rd.Close

        ReDim LAI(0 To 13, 1 To 12)
        Rd.Open "select * from [Land Cover] order by [Category]", ConnectSys, adOpenStatic, adLockReadOnly
        Rd.MoveFirst
        For i = 0 To 13
            LAI(i, 1) = Rd("LAI-Jan")
            LAI(i, 2) = Rd("LAI-Feb")
            LAI(i, 3) = Rd("LAI-Mar")
            LAI(i, 4) = Rd("LAI-Apr")
            LAI(i, 5) = Rd("LAI-May")
            LAI(i, 6) = Rd("LAI-Jun")
            LAI(i, 7) = Rd("LAI-Jul")
            LAI(i, 8) = Rd("LAI-Aug")
            LAI(i, 9) = Rd("LAI-Sep")
            LAI(i, 10) = Rd("LAI-Oct")
            LAI(i, 11) = Rd("LAI-Nov")
            LAI(i, 12) = Rd("LAI-Dec")
            Rd.MoveNext
        Next i
        Rd.Close

        ReDim GridQi(1 To Nx, 1 To Ny), GridQg(1 To Nx, 1 To Ny), GridQs(1 To Nx, 1 To Ny)
        ReDim GridW(1 To Nx, 1 To Ny), GridS(1 To Nx, 1 To Ny), GridFLC(1 To Nx, 1 To Ny)
        ReDim GridWU(1 To Nx, 1 To Ny), GridWL(1 To Nx, 1 To Ny), GridWD(1 To Nx, 1 To Ny)

        LongFDTime = DatePart("YYYY", StartDay) * 10000 + DatePart("M", StartDay) * 100 + DatePart("D", StartDay)
        JYear = DatePart("YYYY", StartDay)
        JMonth = DatePart("M", StartDay)
        CheckFile = Dir(App.Path & "\output\" & StationName & "\日模张力水容量\" & JYear & "\" & LongFDTime & ".asc")
        If CheckFile = "" Then
            MsgBox "请先进行第" & ENo & "次洪水的日洪模拟！程序中断！", vbExclamation + vbInformation, "警告："
            Exit Sub
        End If
        Open App.Path & "\output\" & StationName & "\日模张力水容量\" & JYear & "\" & LongFDTime & ".asc" For Input As #90
        Open App.Path & "\output\" & StationName & "\日模自由水容量\" & JYear & "\" & LongFDTime & ".asc" For Input As #91
        Open App.Path & "\output\" & StationName & "\日模上层张力水容量\" & JYear & "\" & LongFDTime & ".asc" For Input As #92
        Open App.Path & "\output\" & StationName & "\日模下层张力水容量\" & JYear & "\" & LongFDTime & ".asc" For Input As #93
        Open App.Path & "\output\" & StationName & "\日模植被覆盖率\" & JYear & "\" & LongFDTime & ".asc" For Input As #94
        Input #90, Str
        Input #90, Str
        Input #90, Str
        Input #90, Str
        Input #90, Str
        Input #90, Str
        Input #91, Str
        Input #91, Str
        Input #91, Str
        Input #91, Str
        Input #91, Str
        Input #91, Str
        Input #92, Str
        Input #92, Str
        Input #92, Str
        Input #92, Str
        Input #92, Str
        Input #92, Str
        Input #93, Str
        Input #93, Str
        Input #93, Str
        Input #93, Str
        Input #93, Str
        Input #93, Str
        Input #94, Str
        Input #94, Str
        Input #94, Str
        Input #94, Str
        Input #94, Str
        Input #94, Str
        For ii = 1 To Nx
            For jj = 1 To Ny
                GridWM(ii, jj) = Nodata
                GridSM(ii, jj) = Nodata
                Input #90, GridW(ii, jj)
                Input #91, GridS(ii, jj)
                Input #92, GridWU(ii, jj)
                Input #93, GridWL(ii, jj)
                Input #94, GridFLC(ii, jj)
                If GridW(ii, jj) <> Nodata Then
                    GridQi(ii, jj) = QObs(1) / DRMGCNo / 2
                    GridQg(ii, jj) = QObs(1) / DRMGCNo / 2
                    gridfr(ii, jj) = 0.01
                End If
            Next
        Next
        Close #90
        Close #91
        Close #92
        Close #93
        Close #94

        SumDRMIca = 0
        Div = 5
        Ct = Garea / DT / 3.6
        Cg = CCg ^ (DT / 24)
        Ci = CCci ^ (DT / 24)
        For i = 1 To LagTime
            QSim(i) = QObs(i)
        Next i

        ReDim DRMPO(1 To DRMGCNo), DRMET(1 To DRMGCNo)
        ReDim Dis(1 To NPStation), Pcum(1 To Nx, 1 To Ny)
        ReDim Icum(0 To 1, 1 To DRMGCNo), DRMIca(1 To DRMGCNo)
        ReDim DRMIch(1 To DRMGCNo), SimQ(0 To TSteps)
        ReDim RFC(1 To Nx, 1 To Ny), Pnet(1 To DRMGCNo), ETime(0 To TSteps)
        ReDim DRMW0(1 To Nx, 1 To Ny), DRMS0(1 To Nx, 1 To Ny), RFC0(1 To Nx, 1 To Ny)
        If NUPoints > 0 Then
            ReDim UPSimQ(0 To TSteps, 1 To NUPoints)
        End If
        ReDim InflowQs(0 To 1, 1 To DRMGCNo), OutflowQs(0 To 1, 1 To DRMGCNo), InflowQi(0 To 1, 1 To DRMGCNo), OutflowQi(0 To 1, 1 To DRMGCNo)
        ReDim InflowQg(0 To 1, 1 To DRMGCNo), OutflowQg(0 To 1, 1 To DRMGCNo), InflowQch(0 To 1, 1 To DRMGCNo), OutflowQch(0 To 1, 1 To DRMGCNo)

        For i = 1 To TSteps
            If i = 1 Then
                ETime(1) = EStarttime(ENo)
            Else
                ETime(i) = DateAdd("H", DT, ETime(i - 1))
            End If
            JDay = DatePart("M", ETime(i))
            If JDay <> JMonth Then
                LongFDTime = DatePart("YYYY", ETime(i)) * 10000 + DatePart("M", ETime(i)) * 100 + DatePart("D", ETime(i))
                Open App.Path & "\output\" & StationName & "\日模植被覆盖率\" & JYear & "\" & LongFDTime & ".asc" For Input As #94
                Input #94, Str
                Input #94, Str
                Input #94, Str
                Input #94, Str
                Input #94, Str
                Input #94, Str
                For ii = 1 To Nx
                    For jj = 1 To Ny
                        Input #94, GridFLC(ii, jj)
                    Next
                Next
                Close #94
                JMonth = JDay
            End If

            ReDim SumQs(1 To DRMGCNo), SumQi(1 To DRMGCNo), SumQg(1 To DRMGCNo), SumQch(1 To DRMGCNo)
            For k = 1 To DRMGCNo
                ii = SortingRow(k)
                jj = SortingCol(k)
                Sumdis = 0
                Sump = 0
                SaP = False
                Kp = 1
                For j = 1 To NPStation
                    Dis(j) = ((ii - PSRow(j)) ^ 2 + (jj - PSCol(j)) ^ 2) ^ 0.5
                    If Dis(j) = 0 Then
                        Kp = j
                        SaP = True
                        Exit For
                    End If
                    Sump = Sump + PObs(i, j) * Dis(j) ^ (-IW)
                    Sumdis = Sumdis + Dis(j) ^ (-IW)
                Next j
                If SaP = True Then
                    DRMPO(NextNoIJ(ii, jj)) = PObs(i, Kp)
                Else
                    DRMPO(NextNoIJ(ii, jj)) = Sump / Sumdis
                End If

                DRMET(NextNoIJ(ii, jj)) = KEpC * EObs(i)
                GLAI = LAI(LCover(ii, jj), JMonth)
                Scmax = 0.935 + 0.498 * GLAI - 0.00575 * GLAI ^ 2
                Cp = GridFLC(ii, jj)
                Cvd = 0.046 * GLAI
                Pcum(ii, jj) = Pcum(ii, jj) + DRMPO(NextNoIJ(ii, jj))
                Icum(1, NextNoIJ(ii, jj)) = Cp * Scmax * (1 - Exp((-Cvd) * Pcum(ii, jj) / Scmax))
                DRMIca(NextNoIJ(ii, jj)) = Icum(1, NextNoIJ(ii, jj)) - Icum(0, NextNoIJ(ii, jj))
                Icum(0, NextNoIJ(ii, jj)) = Icum(1, NextNoIJ(ii, jj))
                If DRMET(NextNoIJ(ii, jj)) >= DRMIca(NextNoIJ(ii, jj)) Then
                    DRMET(NextNoIJ(ii, jj)) = DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj))
                    DRMIca(NextNoIJ(ii, jj)) = 0
                    If DRMET(NextNoIJ(ii, jj)) >= SumDRMIca Then
                        DRMET(NextNoIJ(ii, jj)) = DRMET(NextNoIJ(ii, jj)) - SumDRMIca
                        SumDRMIca = 0
                    Else
                        SumDRMIca = SumDRMIca - DRMET(NextNoIJ(ii, jj))
                        DRMET(NextNoIJ(ii, jj)) = 0
                    End If
                Else
                    DRMIca(NextNoIJ(ii, jj)) = DRMIca(NextNoIJ(ii, jj)) - DRMET(NextNoIJ(ii, jj))
                    SumDRMIca = SumDRMIca + DRMIca(NextNoIJ(ii, jj))
                    DRMET(NextNoIJ(ii, jj)) = 0
                End If

                If RiverPoint(ii, jj) = 1 Then
                    If FlowDirection(ii, jj) = 0 Or FlowDirection(ii, jj) = 2 Or FlowDirection(ii, jj) = 4 Or FlowDirection(ii, jj) = 6 Then
                        DRMIch(NextNoIJ(ii, jj)) = (DRMPO(NextNoIJ(ii, jj)) - DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj))) * ((DRMBmax(ii, jj) * Cellsize ^ 0.5 * 10 ^ (-6)) / Garea)
                    Else
                        DRMIch(NextNoIJ(ii, jj)) = (DRMPO(NextNoIJ(ii, jj)) - DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj))) * ((DRMBmax(ii, jj) * Cellsize * 10 ^ (-6)) / Garea)
                    End If
                Else
                    DRMIch(NextNoIJ(ii, jj)) = 0
                End If

                If RFC(ii, jj) = 1 Or FlowDirection(ii, jj) = 8 Then
                    Pnet(NextNoIJ(ii, jj)) = DRMPO(NextNoIJ(ii, jj)) - DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj)) '+ SumQs(NextNoIJ(ii, jj)) / Ct
                Else
                    Pnet(NextNoIJ(ii, jj)) = DRMPO(NextNoIJ(ii, jj)) - DRMET(NextNoIJ(ii, jj)) - DRMIca(NextNoIJ(ii, jj)) '+ SumQs(NextNoIJ(ii, jj)) / Ct
                    If Pnet(NextNoIJ(ii, jj)) > 500 Then
                        Pnet(NextNoIJ(ii, jj)) = Pnet(NextNoIJ(ii, jj))
                    End If
                End If

                Pe = Pnet(NextNoIJ(ii, jj))
                Ek = DRMET(NextNoIJ(ii, jj))
                GWM = DRMWM(ii, jj)
                GWMM = GWM * (1# + b)
                If GWM = 0.000015 Then
                    GWM = GWM
                End If
                GSM = DRMSM(ii, jj)
                GSMM = GSM * (1 + ex)
                GW = GridW(ii, jj)
                GS = GridS(ii, jj)
                fr = gridfr(ii, jj)
'                GS = GS / fr
                iff = kim * GW
                GWU = GridWU(ii, jj)
                GWL = GridWL(ii, jj)
                GWD = GW - GWU - GWL
                Kg = DRMKg(ii, jj)
                Ki = DRMKi(ii, jj)
                KKi = (1 - ((1 - (Kg + Ki)) ^ (DT / 24))) / (1 + Kg / Ki)
                KKg = KKi * Kg / Ki
                Kg = KKg
                Ki = KKi
                Grfc = RFC(ii, jj)
                Dp = DRMfc(ii, jj)
                ZUpper = AlUpper * ThickoVZ(ii, jj)
                ZLower = AlLower * ThickoVZ(ii, jj)
                ZDeeper = AlDeeper * ThickoVZ(ii, jj)
                If ZUpper > 300 Then
                    GWUM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * 300 + (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * (ZUpper - 300)
                    GWLM = (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * ZLower
                Else
                    GWUM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * ZUpper
                    If ZUpper + ZLower > 300 Then
                        GWLM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * (300 - ZUpper) + (STFC(SType30100(ii, jj)) - STWP(SType30100(ii, jj))) * (ZLower - 300 + ZUpper)
                    Else
                        GWLM = (STFC(SType030(ii, jj)) - STWP(SType030(ii, jj))) * ZLower
                    End If
                End If
                GWDM = GWM - GWUM - GWLM
                If GWM <> 0 Then
                    If Pe <= 0 Then
                        Grfc = 0
                        If GWU + Pe < 0# Then
                            GEU = GWU + Ek + Pe
                            GWU = 0
                            GEL = (Ek - GEU) * GWL / GWLM
                            If GWL < DeeperC * GWLM Then
                                GEL = DeeperC * (Ek - GEU)
                            End If
                            If (GWL - GEL) < 0# Then
                                ged = GEL - GWL
                                GEL = GWL
                                GWL = 0
                                GWD = GWD - ged
                            Else
                                ged = 0
                                GWL = GWL - GEL
                            End If
                        Else
                            GEU = Ek
                            GEL = 0
                            ged = 0
                            GWU = GWU + Pe
                        End If
                        GW = GWU + GWL + GWD
                        If GW < 0 Then
                            GEU = 0
                            GEL = 0
                            ged = GridW(ii, jj)
                            GE = ged
                            GWU = 0
                            GWL = 0
                            GWD = 0
                            GW = 0
                        End If
                        GE = GEU + GEL + ged
                        SumDRMIca = Ek - GE
                        SumDRMIca = IIf(SumDRMIca < 0, 0, SumDRMIca)
                        GRs = 0
                        GRe = 0
                        GRg = GS * Kg
                        GRi = GS * Ki
                        GS = GS * (1 - Kg - Ki)
                    Else
                        GEU = Ek
                        GEL = 0
                        ged = 0
                        GE = Ek
                        nd = Int(Pe / Div) + 1
                        If Pe Mod Div = 0 Then
                            If Pe - Int(Pe) = 0 Then
                                nd = nd - 1
                            End If
                        End If
                        ReDim PPe(1 To nd), fcd(1 To nd), fmd(1 To nd), fxd(1 To nd)
                        For m = 1 To nd - 1
                            PPe(m) = Div
                        Next m
                        PPe(nd) = Pe - (nd - 1) * Div
                        For m = 1 To nd
                            fcd(m) = PPe(m) / Pe * fc(ii, jj)
                            fmd(m) = PPe(m) / Pe * fm(ii, jj)
                        Next m
                        GRe = 0
                        GRs = 0
                        GRg = 0
                        GRi = 0
                        KKi = (1# - (1# - (Kg + Ki)) ^ (1# / nd)) / (Kg + Ki)
                        KKg = KKi * Kg
                        KKi = KKi * Ki
                        aa = 1# - (1 - GW / (GWM + 0.0001)) ^ (1# / (1# + b))
                        aa = GWMM * aa
                        R1 = 0#
                        peds = 0#
                        sfx = 0#
                        For m = 1 To nd
'                            fd = fmd(m) + (fcd(m) - fmd(m)) * GW / (GWM + 0.000001)
'                            fdmax = fd * (1 + ex2)
'                            If PPe(m) < fdmax Then
'                                R4 = PPe(m) - fd + fd * (1 - PPe(m) / fdmax) ^ (1 + ex2)
'                                GRe = GRe + R4
'                                fxd(m) = PPe(m) - R4
'                            Else
'                                R4 = PPe(m) - fd
'                                GRe = GRe + R4
'                                fxd(m) = fd
'                            End If

'                            ffc = sfx + iff
'                            fd = fcd(m) + (fmd(m) - fcd(m)) * Exp(-uu * ffc)
'                            fdmax = fd * (1 + ex2)
'                            If PPe(m) < fdmax Then
'                                R4 = PPe(m) - fd + fd * (1 - PPe(m) / fdmax) ^ (1 + ex2)
'                                GRe = GRe + R4
'                                fxd(m) = PPe(m) - R4
'                            Else
'                                R4 = PPe(m) - fd
'                                GRe = GRe + R4
'                                fxd(m) = fd
'                            End If
                            
                            ffc = sfx + iff
                            fd = fmd(m)
                            fdmax = fd * (1 + ex2)
                            If PPe(m) < fdmax Then
                                fxd(m) = fd - fd * (1 - PPe(m) / fdmax) ^ (1 + ex2)
                                If fxd(m) > fcd(m) Then
                                    fxd(m) = fcd(m) + (fxd(m) - fcd(m)) * Exp(-uu * ffc)
                                Else
                                    fxd(m) = fxd(m)
                                End If
                                R4 = PPe(m) - fxd(m)
                                GRe = GRe + R4
                            Else
                                fxd(m) = fd
                                If fxd(m) > fcd(m) Then
                                    fxd(m) = fcd(m) + (fxd(m) - fcd(m)) * Exp(-uu * ffc)
                                Else
                                    fxd(m) = fxd(m)
                                End If
                                R4 = PPe(m) - fxd(m)
                                GRe = GRe + R4
                            End If

                                
                            aa = aa + fxd(m)
                            R2 = R1
                            peds = peds + fxd(m)
                            If aa < GWMM Then
                                R1 = peds - GWM + GW + GWM * (1# - aa / GWMM) ^ (1 + b)
                                GR = R1 - R2
                                GRg = GRg + GR
                            Else
                                R1 = peds - GWM + GW
                                GR = R1 - R2
                                If GR > fcd(m) Then
                                    GRe = GRe + GR - fcd(m)
                                    GRg = GRg + fcd(m)
                                    fxd(m) = fxd(m) - (GR - fcd(m))
                                Else
                                    GRg = GRg + GR
                                End If
                            End If
                            sfx = sfx + fxd(m)
                            If sfx + iff > kim * GWMM Then
                                sfx = kim * GWMM - iff
                            End If
                            
'                            xx = fr
'                            fr = GR / (fxd(m) + 0.00001)
'                            GS = GS * xx / fr
'                            If GS >= GSM Then
'                                R3 = (fxd(m) + GS - GSM) * fr
'                                GRs = GRs + R3
'                                GS = GSM
'                                GRi = GRi + GS * KKi * fr
'                                GRg = GRg + GS * KKg * fr
'                                GS = GS * (1# - KKi - KKg)
'                            End If
'                            au = GSMM * (1# - (1# - GS / GSM) ^ (1# / (1# + ex)))
'                            ff = au + fxd(m)
'                            If ff < GSMM Then
'                                R3 = (fxd(m) - GSM + GS + GSM * (1# - ff / GSMM) ^ (1 + ex)) * fr
'                                GRs = GRs + R3
'                                GS = fxd(m) + GS - R3 / fr
'                                GRi = GRi + GS * KKi * fr
'                                GRg = GRg + GS * KKg * fr
'                                GS = GS * (1# - KKi - KKg)
'                            Else
'                                R3 = (fxd(m) + GS - GSM) * fr
'                                GRs = GRs + R3
'                                GS = GSM
'                                GRi = GRi + GS * KKi * fr
'                                GRg = GRg + GS * KKg * fr
'                                GS = GS * (1# - KKi - KKg)
'                            End If
                        Next m
                        If GWU + Pe - GRg - GRe < GWUM Then
                           GWU = GWU + Pe - GRg - GRe
                           GW = GWU + GWL + GWD
                        ElseIf GWU + GWL + Pe - GRg - GRe - GWUM >= GWLM Then
                            GWU = GWUM
                            GWL = GWLM
                            GWD = GW + Pe - GRg - GRe - GWU - GWL
                            If GWD > GWDM Then
                                GWD = GWDM
                            End If
                            GW = GWU + GWL + GWD
                        Else
                            GWL = GWU + GWL + Pe - GRg - GRe - GWUM
                            GWU = GWUM
                            GW = GWU + GWL + GWD
                        End If
                        If GW > GWM Then
                            GW = GWM
                        End If
                    End If
                    If GRs = 0 Then
                        Grfc = 0
                    Else
                        Grfc = 1
                    End If
                Else
                    If Pe <= 0 Then
                        GEU = DRMPO(NextNoIJ(ii, jj)): GEL = 0: ged = 0
                        GRs = 0: GRe = 0: GRg = 0: GRi = 0
                    Else
                        GEU = Ek: GEL = 0: ged = 0
                        GRs = Pe: GRe = 0: GRg = 0: GRi = 0
                    End If
                End If
                GRs = GRs + GRe
'                If Grfc = 1 Then
'                    baohe(ii, jj) = 2
'                Else
'                    baohe(ii, jj) = 3
'                End If
                GQi = GridQi(ii, jj)
                GQg = GridQg(ii, jj)
                GQs = GridQs(ii, jj)
                CCS1 = 0
                GQs = GQs * CCS1 + GRs * Ct * (1 - CCS1) * (1 - Dp)
                Gqch = GRs * Ct * Dp
                GQi = GQi * Ci + GRi * Ct * (1 - Ci)
                GQg = GQg * Cg + GRg * Ct * (1 - Cg)
                GridQi(ii, jj) = GQi
                GridQg(ii, jj) = GQg
                GridQs(ii, jj) = GQs
                gridRg(ii, jj) = GRg
                GW = GWU + GWL + GWD
                GridW(ii, jj) = GW
                GridS(ii, jj) = GS
                gridfr(ii, jj) = fr
                If GW >= GWM Then
                    GridWM(ii, jj) = 1
                    NoWM(i) = NoWM(i) + 1
                Else
                    GridWM(ii, jj) = 0
                End If
                If GS / (1 - KKg - KKi) >= GSM Then
                    GridSM(ii, jj) = 1
                    NoSM(i) = NoSM(i) + 1
                Else
                    GridSM(ii, jj) = 0
                End If
                GridWU(ii, jj) = GWU
                GridWL(ii, jj) = GWL
                GridWD(ii, jj) = GWD
                RFC(ii, jj) = Grfc
                DRMET(NextNoIJ(ii, jj)) = GE

                Select Case FlowDirection(ii, jj)
                Case 0
                    i1 = ii - 1
                    j1 = jj + 1
                Case 1
                    i1 = ii
                    j1 = jj + 1
                Case 2
                    i1 = ii + 1
                    j1 = jj + 1
                Case 3
                    i1 = ii + 1
                    j1 = jj
                Case 4
                    i1 = ii + 1
                    j1 = jj - 1
                Case 5
                    i1 = ii
                    j1 = jj - 1
                Case 6
                    i1 = ii - 1
                    j1 = jj - 1
                Case 7
                    i1 = ii - 1
                    j1 = jj
                Case 8
                    i1 = ii
                    j1 = jj
                Case Else
                End Select

                InflowQs(1, NextNoIJ(ii, jj)) = GQs + SumQs(NextNoIJ(ii, jj))
                If i = 1 Then
                    OutflowQs(1, NextNoIJ(ii, jj)) = InflowQs(1, NextNoIJ(ii, jj))
                End If
                If i > 1 Then
                    MDt = DT
                    MC1 = (MXs * MKs + 0.5 * MDt) / ((1 - MXs) * MKs + 0.5 * MDt)
                    MC2 = (0.5 * MDt - MXs * MKs) / ((1 - MXs) * MKs + 0.5 * MDt)
                    MC3 = ((1 - MXs) * MKs - 0.5 * MDt) / ((1 - MXs) * MKs + 0.5 * MDt)
                    OutflowQs(1, NextNoIJ(ii, jj)) = MC1 * InflowQs(0, NextNoIJ(ii, jj)) + MC2 * InflowQs(1, NextNoIJ(ii, jj)) + MC3 * OutflowQs(0, NextNoIJ(ii, jj))
                    OutflowQs(1, NextNoIJ(ii, jj)) = IIf(OutflowQs(1, NextNoIJ(ii, jj)) > 0, OutflowQs(1, NextNoIJ(ii, jj)), 0)
                    InflowQs(0, NextNoIJ(ii, jj)) = InflowQs(1, NextNoIJ(ii, jj))
                    OutflowQs(0, NextNoIJ(ii, jj)) = OutflowQs(1, NextNoIJ(ii, jj))
                End If
                If FlowDirection(ii, jj) = 8 Then
                    SumQs(NextNoIJ(i1, j1)) = OutflowQs(1, NextNoIJ(ii, jj))
                Else
                    SumQs(NextNoIJ(i1, j1)) = SumQs(NextNoIJ(i1, j1)) + OutflowQs(1, NextNoIJ(ii, jj))
                End If
                InflowQi(1, NextNoIJ(ii, jj)) = GQi + SumQi(NextNoIJ(ii, jj))
                If i = 1 Then
                    OutflowQi(1, NextNoIJ(ii, jj)) = InflowQi(1, NextNoIJ(ii, jj))
                End If
                If i > 1 Then
                    MDt = DT
                    MC1 = (MXi * MKi + 0.5 * MDt) / ((1 - MXi) * MKi + 0.5 * MDt)
                    MC2 = (0.5 * MDt - MXi * MKi) / ((1 - MXi) * MKi + 0.5 * MDt)
                    MC3 = ((1 - MXi) * MKi - 0.5 * MDt) / ((1 - MXi) * MKi + 0.5 * MDt)
                    OutflowQi(1, NextNoIJ(ii, jj)) = MC1 * InflowQi(0, NextNoIJ(ii, jj)) + MC2 * InflowQi(1, NextNoIJ(ii, jj)) + MC3 * OutflowQi(0, NextNoIJ(ii, jj))
                    OutflowQi(1, NextNoIJ(ii, jj)) = IIf(OutflowQi(1, NextNoIJ(ii, jj)) > 0, OutflowQi(1, NextNoIJ(ii, jj)), 0)
                    InflowQi(0, NextNoIJ(ii, jj)) = InflowQi(1, NextNoIJ(ii, jj))
                    OutflowQi(0, NextNoIJ(ii, jj)) = OutflowQi(1, NextNoIJ(ii, jj))
                End If
                If FlowDirection(ii, jj) = 8 Then
                    SumQi(NextNoIJ(i1, j1)) = OutflowQi(1, NextNoIJ(ii, jj))
                Else
                    SumQi(NextNoIJ(i1, j1)) = SumQi(NextNoIJ(i1, j1)) + OutflowQi(1, NextNoIJ(ii, jj))
                End If
                InflowQg(1, NextNoIJ(ii, jj)) = GQg + SumQg(NextNoIJ(ii, jj))
                If i = 1 Then
                    OutflowQg(1, NextNoIJ(ii, jj)) = InflowQg(1, NextNoIJ(ii, jj))
                End If
                If i > 1 Then
                    MDt = DT
                    MC1 = (MXg * MKg + 0.5 * MDt) / ((1 - MXg) * MKg + 0.5 * MDt)
                    MC2 = (0.5 * MDt - MXg * MKg) / ((1 - MXg) * MKg + 0.5 * MDt)
                    MC3 = ((1 - MXg) * MKg - 0.5 * MDt) / ((1 - MXg) * MKg + 0.5 * MDt)
                    OutflowQg(1, NextNoIJ(ii, jj)) = MC1 * InflowQg(0, NextNoIJ(ii, jj)) + MC2 * InflowQg(1, NextNoIJ(ii, jj)) + MC3 * OutflowQg(0, NextNoIJ(ii, jj))
                    OutflowQg(1, NextNoIJ(ii, jj)) = IIf(OutflowQg(1, NextNoIJ(ii, jj)) > 0, OutflowQg(1, NextNoIJ(ii, jj)), 0)
                    InflowQg(0, NextNoIJ(ii, jj)) = InflowQg(1, NextNoIJ(ii, jj))
                    OutflowQg(0, NextNoIJ(ii, jj)) = OutflowQg(1, NextNoIJ(ii, jj))
                End If
                If FlowDirection(ii, jj) = 8 Then
                    SumQg(NextNoIJ(i1, j1)) = OutflowQg(1, NextNoIJ(ii, jj))
                Else
                    SumQg(NextNoIJ(i1, j1)) = SumQg(NextNoIJ(i1, j1)) + OutflowQg(1, NextNoIJ(ii, jj))
                End If
                InflowQch(1, NextNoIJ(ii, jj)) = Gqch + SumQch(NextNoIJ(ii, jj))
                If i = 1 Then
                    OutflowQch(1, NextNoIJ(ii, jj)) = InflowQch(1, NextNoIJ(ii, jj))
                End If
                If i > 1 Then
                    MDt = DT
                    MC1 = (MXch * MKch + 0.5 * MDt) / ((1 - MXch) * MKch + 0.5 * MDt)
                    MC2 = (0.5 * MDt - MXch * MKch) / ((1 - MXch) * MKch + 0.5 * MDt)
                    MC3 = ((1 - MXch) * MKch - 0.5 * MDt) / ((1 - MXch) * MKch + 0.5 * MDt)
                    OutflowQch(1, NextNoIJ(ii, jj)) = MC1 * InflowQch(0, NextNoIJ(ii, jj)) + MC2 * InflowQch(1, NextNoIJ(ii, jj)) + MC3 * OutflowQch(0, NextNoIJ(ii, jj))
                    OutflowQch(1, NextNoIJ(ii, jj)) = IIf(OutflowQch(1, NextNoIJ(ii, jj)) > 0, OutflowQch(1, NextNoIJ(ii, jj)), 0)
                    InflowQch(0, NextNoIJ(ii, jj)) = InflowQch(1, NextNoIJ(ii, jj))
                    OutflowQch(0, NextNoIJ(ii, jj)) = OutflowQch(1, NextNoIJ(ii, jj))
                End If
                If RiverPoint(i1, j1) = 1 Then
                    If FlowDirection(ii, jj) = 8 Then
                        SumQch(NextNoIJ(i1, j1)) = OutflowQch(1, NextNoIJ(ii, jj)) + SumQi(NextNoIJ(i1, j1)) + SumQg(NextNoIJ(i1, j1))
                    Else
                        SumQch(NextNoIJ(i1, j1)) = SumQch(NextNoIJ(i1, j1)) + OutflowQch(1, NextNoIJ(ii, jj)) + SumQi(NextNoIJ(i1, j1)) + SumQg(NextNoIJ(i1, j1))
                    End If
                    SumQi(NextNoIJ(i1, j1)) = 0
                    SumQg(NextNoIJ(i1, j1)) = 0
                End If
                If FlowDirection(ii, jj) = 8 Then
                    Qoutch = SumQch(NextNoIJ(i1, j1))
                    Qouts = SumQs(NextNoIJ(i1, j1))
                    Qouti = SumQi(NextNoIJ(i1, j1))
                    Qoutg = SumQg(NextNoIJ(i1, j1))
                    QSim(i + LagTime) = QSim(i + LagTime - 1) * CCS + (Qoutch + Qouts + Qouti + Qoutg) * (1 - CCS)
                End If
                gridQg2(ii, jj) = OutflowQg(1, NextNoIJ(ii, jj)) + OutflowQs(1, NextNoIJ(ii, jj)) + OutflowQi(1, NextNoIJ(ii, jj)) + OutflowQch(1, NextNoIJ(ii, jj))
            Next k

'            If ENo = 13 Then
'                Open App.Path & "\Output\" & StationName & "\" & "时段径流" & "\" & StationName & "Q" & i & ".asc" For Output As #100
'                Print #100, "ncols      ", Ny
'                Print #100, "nrows      ", Nx
'                Print #100, "xllcorner  ", XllCorner
'                Print #100, "yllcorner  ", YllCorner
'                Print #100, "cellsize   ", DDem
'                Print #100, "NODATA_value ", 0
'                For x = 1 To Nx
'                        For y = 1 To Ny
'                                Print #100, gridQg2(x, y);
'                        Next y
'                        Print #100,
'                Next x
'                Close #100
'            End If
'             Open App.Path & "\Output\" & StationName & "\" & "蓄超分布" & "\" & StationName & "第" & ENo & "场" & "蓄超" & i & ".asc" For Output As #101
'                Print #101, "ncols      ", Ny
'                Print #101, "nrows      ", Nx
'                Print #101, "xllcorner  ", XllCorner
'                Print #101, "yllcorner  ", YllCorner
'                Print #101, "cellsize   ", DDem
'                Print #101, "NODATA_value ", 0
'                For x = 1 To Nx
'                    For y = 1 To Ny
'                            Print #101, baohe(x, y);
'                    Next y
'                    Print #101,
'                Next x
'                Close #101
            If NUPoints > 0 Then
                For m = 1 To NUPoints
                    ii = UPRow(m)
                    jj = UPCol(m)
                    If i = 1 Then
                        UPSimQ(1, m) = SumQch(NextNoIJ(ii, jj)) + SumQs(NextNoIJ(ii, jj)) + SumQi(NextNoIJ(ii, jj)) + SumQg(NextNoIJ(ii, jj))
                    Else
                        UPSimQ(i, m) = UPSimQ(i - 1, m) * CCS + (SumQch(NextNoIJ(ii, jj)) + SumQs(NextNoIJ(ii, jj)) + SumQi(NextNoIJ(ii, jj)) + SumQg(NextNoIJ(ii, jj))) * (1 - CCS)
                    End If
                Next m
            End If
        Next i

        If NUPoints > 0 Then
            On Error GoTo AccessError
            ConnectSys.Execute "delete * from [UpHResultsMK-" & SSName & "] where [时间] between  #" & Format(EStarttime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and #" & Format(EEndtime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and [目的]= '" & Aimat & " '"
            Rd.Open "select * from [UpHResultsMK-" & SSName & "] ", ConnectSys, adOpenDynamic, adLockOptimistic
            For i = 1 To TSteps
                Rd.AddNew
                Rd("洪号") = ENo
                Rd("目的") = Aimat
                Rd("时间") = TimeSeries(i)
                For j = 1 To NUPoints
                    Rd(UPName(j)) = UPSimQ(i, j)
                Next j
                Rd.Update
            Next i
            Rd.Close
            GoTo NoAccessError
AccessError:
            MsgBox "请在表UpHResultsMK-" & SSName & "中添加上游入流点[" & UPName(j) & "]字段！"
            Rd.Update
            Rd.Close
        End If
NoAccessError:
        ConnectSys.Execute "delete * from [HResultsMK-" & SSName & "] where [Time] between  #" & Format(EStarttime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and #" & Format(EEndtime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and [Purpose]= '" & Aimat & " '"
        ConnectSys.Execute "delete * from [HCResultsMK-" & SSName & "] where [Start Time] = #" & Format(EStarttime(ENo), "YYYY-MM-DD HH:MM:SS") & "# and [Purpose]= '" & Aimat & " '"

        With Rd
            SumOQ = 0
            SumSQ = 0
            OPeak = 0
            SPeak = 0
            SumPre = 0
            SumEE = 0
            ONC = 0
            SNC = 0

            For i = 1 To TSteps
                SumOQ = SumOQ + QObs(i)
                SumSQ = SumSQ + QSim(i)
                If OPeak < QObs(i) Then
                    OPeak = QObs(i)
                    OPeakTime = i
                End If
                If SPeak < QSim(i) Then
                    SPeak = QSim(i)
                    SPeakTime = i
                End If
                SumPre = SumPre + AvgP(i)
            Next i

            AvgOQ = SumOQ / TSteps
            .Open "select * from [HResultsMK-" & SSName & "] ", ConnectSys, adOpenDynamic, adLockOptimistic
            For i = 1 To TSteps
                ONC = ONC + (QObs(i) - AvgOQ) ^ 2
                SNC = SNC + (QSim(i) - QObs(i)) ^ 2
                .AddNew
                Rd("FloodNo") = ENo
                Rd("Purpose") = Aimat
                Rd("Time") = TimeSeries(i)
                Rd("SimulatedQ") = QSim(i)
                Rd("ObservedQ") = QObs(i)
                Rd("AverageP") = AvgP(i)
                .Update
            Next i
            .Close
            ORunoff = SumOQ * 3.6 * DT / DArea
            SRunoff = SumSQ * 3.6 * DT / DArea
            .Open "select * from [HCResultsMK-" & SSName & "] ", ConnectSys, adOpenDynamic, adLockOptimistic
            .AddNew
            Rd("FloodNo") = ENo
            Rd("Purpose") = Aimat
            Rd("Start Time") = TimeSeries(1)
            Rd("Precipitation") = SumPre
            Rd("ObservedRO") = ORunoff
            Rd("SimulatedRO") = SRunoff
            Rd("RO Error(%)") = (SRunoff - ORunoff) / ORunoff * 100
            Rd("ObservedPeak") = OPeak
            Rd("SimulatedPeak") = SPeak
            Rd("Peak Error(%)") = (SPeak - OPeak) / OPeak * 100
            Rd("Time Error") = SPeakTime - OPeakTime
            Rd("NC") = 1 - SNC / ONC
'            If Abs(Rd("Peak Error(%)")) > 20 Then
'                MPeakE = 0
'            Else
'                MPeakE = 2.5 * (1 - (Abs(Rd("Peak Error(%)")) / 20))
'            End If
'            If Abs(Rd("RO Error(%)")) > 20 Then
'                MRunoffE = 0
'            Else
'                MRunoffE = 2.5 * (1 - (Abs(Rd("RO Error(%)")) / 20))
'            End If
'            If Abs(Rd("Time Error")) > 3 Then
'                MTimeE = 0
'            Else
'                MTimeE = 2.5 * (1 - (Abs(Rd("Time Error")) / 4))
'            End If
'            If Rd("NC") < 0.5 Then
'                MNashC = 0
'            Else
'                MNashC = 2.5 * Rd("NC")
'            End If
'                Rd("Grading-marks") = MPeakE + MRunoffE + MTimeE + MNashC
                .Update
            .Close
        End With
    Next ENo
    HourlyScale = True
End Sub

