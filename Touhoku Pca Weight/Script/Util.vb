Imports Microsoft.Office.Interop.Excel

Friend Module Util
    ''' <summary>
    ''' 運賃(2トン車).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Fare(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            DctVal(xlApp, "BA271", choosen)
            DctVal(xlApp, "BA158", 3) ' D16
            DctVal(xlApp, "BA159", 2) ' D13
            DctVal(xlApp, "BA160", 4) ' D10
        Else
            DctVal(xlApp, "BA161", 2) ' D16
            DctVal(xlApp, "BA162", 1) ' D13
            DctVal(xlApp, "BA163", 3) ' D10
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-300.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA17", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA18", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA19", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA20", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA21", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA22", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA23", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA24", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周/内周GL-150.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Unit150(xlApp As Application)
        PubDVal(xlApp, "BA27", DtlDInp(vbTab & "  4G: "))
        PubDVal(xlApp, "BA28", DtlDInp(vbTab & "3.5G: "))
        PubDVal(xlApp, "BA29", DtlDInp(vbTab & "  3G: "))
        PubDVal(xlApp, "BA30", DtlDInp(vbTab & "2.5G: "))
        PubDVal(xlApp, "BA31", DtlDInp(vbTab & "  2G: "))
        PubDVal(xlApp, "BA32", DtlDInp(vbTab & "1.5G: "))
        PubDVal(xlApp, "BA33", DtlDInp(vbTab & "  1G: "))
        PubDVal(xlApp, "BA34", DtlDInp(vbTab & "0.5G: "))
    End Sub

    ''' <summary>
    ''' 外周深GL-300/+30.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300Cut(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA36", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA37", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA38", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA39", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA40", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA41", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA42", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA43", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-400.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit400(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA81", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA82", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA83", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA84", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA85", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA86", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA87", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA88", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-500.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit500(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA45", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA46", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA47", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA48", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA49", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA50", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA51", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA52", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-500/+30.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit500Cut(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA54", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA55", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA56", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA57", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA58", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA59", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA60", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA61", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-600.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit600(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA63", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA64", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA65", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA66", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA67", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA68", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA69", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA70", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' 外周深GL-700.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit700(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA72", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA73", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA74", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA75", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA76", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA77", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA78", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA79", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' ガレージ外周GL-300.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Unit300Gar(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA90", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA91", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA92", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA93", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA94", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA95", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA96", DtlDInp(vbTab & "  1G: "))
            PubDVal(xlApp, "BA97", DtlDInp(vbTab & "0.5G: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブユニット.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub SlabUnit(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA99", DtlDInp(vbTab & "  4G: "))
            PubDVal(xlApp, "BA100", DtlDInp(vbTab & "3.5G: "))
            PubDVal(xlApp, "BA101", DtlDInp(vbTab & "  3G: "))
            PubDVal(xlApp, "BA102", DtlDInp(vbTab & "2.5G: "))
            PubDVal(xlApp, "BA103", DtlDInp(vbTab & "  2G: "))
            PubDVal(xlApp, "BA104", DtlDInp(vbTab & "1.5G: "))
            PubDVal(xlApp, "BA105", DtlDInp(vbTab & "  1G: "))
        End If
    End Sub

    ''' <summary>
    ''' 電気温水器.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="value">Value.</param>
    Friend Sub ElecWtrHtr(xlApp As Application, value As Double)
        If value > 0 Then
            DctVal(xlApp, "BA107", value)
        Else
            ClrVal(xlApp, "BA107")
        End If
    End Sub

    ''' <summary>
    ''' スリーブ.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="branch">Branch GL-300.</param>
    ''' <param name="unitSlab">Slab unit.</param>
    Friend Sub Sleeve(xlApp As Application, branch As Double, unitSlab As Double)
        If branch = 1 Then
            Dim value = HdrDInp(vbTab & vbTab & "スリーブ: ")
            If value > 0 Then
                SleeveMain(xlApp, unitSlab, value)
                DctVal(xlApp, "BA222", value)
            End If
        Else
            Dim value = HdrDInpDesc(vbTab & vbTab & "スリーブ ", "(フック ※)")
            If value > 0 Then
                SleeveMain(xlApp, unitSlab, value)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Sleeve main.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="unitSlab">Slab unit.</param>
    ''' <param name="value">Value.</param>
    Private Sub SleeveMain(xlApp As Application, unitSlab As Double, value As Double)
        If unitSlab = 1 Then
            DctVal(xlApp, "BA219", value)
        Else
            DctVal(xlApp, "BA219", value + 2)
        End If
        DctVal(xlApp, "BA220", value)
        DctVal(xlApp, "BA221", value)
        DctVal(xlApp, "BA224", value)
    End Sub

    ''' <summary>
    ''' ストレート.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub JtStr(xlApp As Application)
        PubDVal(xlApp, "BA166", DtlDInp(vbTab & "D16: "))
        PubDVal(xlApp, "BA165", DtlDInp(vbTab & "D13: "))
        PubDVal(xlApp, "BA164", DtlDInp(vbTab & "D10: "))
    End Sub

    ''' <summary>
    ''' コーナー.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub JtCor(xlApp As Application)
        PubDVal(xlApp, "BA169", DtlDInp(vbTab & "D16: "))
        PubDVal(xlApp, "BA168", DtlDInp(vbTab & "D13: "))
        PubDVal(xlApp, "BA167", DtlDInp(vbTab & "D10: "))
    End Sub

    ''' <summary>
    ''' ロングコーナー.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub LongCor(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA173", DtlDInp(vbTab & "D16 ( 750×2250): "))
            PubDVal(xlApp, "BA171", DtlDInp(vbTab & "D16 ( 750×1750): "))
            PubDVal(xlApp, "BA172", DtlDInp(vbTab & "D16 ( 750×1250): "))
            PubDVal(xlApp, "BA174", DtlDInp(vbTab & "D16 (1250×1750): "))
            PubDVal(xlApp, "BA176", DtlDInp(vbTab & "D10 ( 500×1500): "))
            PubDVal(xlApp, "BA175", DtlDInp(vbTab & "D10 ( 500×1000): "))
        End If
    End Sub

    ''' <summary>
    ''' 端部 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Edge(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA178", DtlDInp(vbTab & "670×780: "))
            PubDVal(xlApp, "BA177", DtlDInp(vbTab & "570×780: "))
            PubDVal(xlApp, "BA170", DtlDInp(vbTab & "390×780: "))
        End If
    End Sub

    ''' <summary>
    ''' クランク.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Crank(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA184", DtlDInp(vbTab & "D16 (750×920×750): "))
            PubDVal(xlApp, "BA185", DtlDInp(vbTab & "D10 (500×910×500): "))
            PubDVal(xlApp, "BA186", DtlDInp(vbTab & "D16 (750×465×750): "))
            PubDVal(xlApp, "BA187", DtlDInp(vbTab & "D10 (500×460×500): "))
        End If
    End Sub

    ''' <summary>
    ''' 島 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Island(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA180", DtlDInp(vbTab & "390×910×390" & vbTab & vbTab & ": "))
            PubDModVal(xlApp, "212", "390×690×390", 2.4, DtlDInpDesc(vbTab & "390×690×390 ", "[2.4]" & vbTab))
            PubDVal(xlApp, "BA179", DtlDInp(vbTab & "390×455×390" & vbTab & vbTab & ": "))
        End If
    End Sub

    ''' <summary>
    ''' ストレート (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Straight(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA199", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "BA200", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA201", DtlDInp(vbTab & "2000: "))
        End If
    End Sub

    ''' <summary>
    ''' キャップタイヤ (320).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub CapTire(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA181", DtlDInp(vbTab & "D16: "))
            PubDVal(xlApp, "BA183", DtlDInp(vbTab & "D10: "))
        End If
    End Sub

    ''' <summary>
    ''' ハンチ (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Haunch(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA208", DtlDInp(vbTab & "H200" & vbTab & vbTab & ": "))
            PubDModVal(xlApp, "207", "660×曲（H300）×660", 2.9, DtlDInpDesc(vbTab & "H300 ", "[2.9]" & vbTab))
        End If
    End Sub

    ''' <summary>
    ''' コーナー(曲 135°).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Corner135deg(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA217", DtlDInp(vbTab & "D16: "))
            PubDVal(xlApp, "BA216", DtlDInp(vbTab & "D13: "))
            PubDVal(xlApp, "BA215", DtlDInp(vbTab & "D10: "))
        End If
    End Sub

    ''' <summary>
    ''' フック (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Hook(xlApp As Application)
        PubDVal(xlApp, "BA192", DtlDInp(vbTab & "795×160" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA204", DtlDInp(vbTab & "695×160" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA189", DtlDInp(vbTab & "595×160" & vbTab & vbTab & ": "))
        PubDModVal(xlApp, "205", "555×160　　フック付", 0.5, DtlDInpDesc(vbTab & "555×160 ", "[0.5]" & vbTab))
        PubDVal(xlApp, "BA191", DtlDInp(vbTab & "455×160" & vbTab & vbTab & ": "))
        PubDModVal(xlApp, "214", "360×160　　フック付", 0.4, DtlDInpDesc(vbTab & "360×160 ", "[0.4]" & vbTab))
        PubDVal(xlApp, "BA206", DtlDInp(vbTab & "260×160" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA190", DtlDInp(vbTab & "160×160" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA188", DtlDInp(vbTab & "435×250" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA218", DtlDInp(vbTab & "溶接 (220×85)" & vbTab & ": "))
    End Sub

    ''' <summary>
    ''' コーナー3 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Corner3d(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "202", "750×690×390", 3, DtlDInpDesc(vbTab & "右 (750×690×390) ", "[3.0]" & vbTab))
            PubDModVal(xlApp, "203", "750×690×390", 3, DtlDInpDesc(vbTab & "左 (750×690×390) ", "[3.0]" & vbTab))
            PubDVal(xlApp, "BA193", DtlDInp(vbTab & "右 (750×460×390)" & vbTab & ": "))
            PubDVal(xlApp, "BA194", DtlDInp(vbTab & "左 (750×460×390)" & vbTab & ": "))
            PubDModVal(xlApp, "209", "750×240×390", 2.3, DtlDInpDesc(vbTab & "右 (750×240×390) ", "[2.3]" & vbTab))
        End If
    End Sub

    ''' <summary>
    ''' クランク3 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Crank3d(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "196", "（クランク３右）", "750×920×460×390", 4.1, DtlDInpDesc(vbTab & "右 (750×920×460×390) ", "[4.1]"))
            PubDModVal(xlApp, "197", "（クランク３左）", "750×920×460×390", 4.1, DtlDInpDesc(vbTab & "左 (750×920×460×390) ", "[4.1]"))
        End If
    End Sub

    ''' <summary>
    ''' M型 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub MType(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "195", "（Ｍ型コーナー）", "390×460×460×390", 2.8, DtlDInpDesc(vbTab & "390×460×460×390 ", "[2.8]"))
            PubDModVal(xlApp, "210", "（Ｍ型コーナー）", "390×460×690×390", 3.2, DtlDInpDesc(vbTab & "390×460×690×390 ", "[3.2]"))
            PubDModVal(xlApp, "211", "（Ｍ型コーナー）", "390×690×690×390", 3.5, DtlDInpDesc(vbTab & "390×690×690×390 ", "[3.5]"))
        End If
    End Sub

    ''' <summary>
    ''' 主筋補強 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub MainReinf(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDVal(xlApp, "BA226", DtlDInp(vbTab & "    2000: "))
            PubDVal(xlApp, "BA225", DtlDInp(vbTab & "    1500: "))
            PubDVal(xlApp, "BA227", DtlDInp(vbTab & "500×1500: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ曲 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="unitSlab">Slab unit.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabBndg(xlApp As Application, unitSlab As Double, truck2Ton As Double)
        If unitSlab = 1 Then
            If HdrYNQ(vbTab & vbTab & "スラブ曲 (D13): ") = 1 Then
                SlabBndgMain(xlApp, truck2Ton)
            End If
        Else
            PrefWarn(vbTab & vbTab & "スラブ曲 (D13)")
            SlabBndgMain(xlApp, truck2Ton)
        End If
    End Sub

    ''' <summary>
    ''' Slab bending main.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Private Sub SlabBndgMain(xlApp As Application, truck2Ton As Double)
        If Not truck2Ton = 1 Then
            PubDVal(xlApp, "BA115", DtlDInp(vbTab & "250×5250" & vbTab & vbTab & ": "))
            PubDVal(xlApp, "BA116", DtlDInp(vbTab & "250×4750" & vbTab & vbTab & ": "))
        End If
        PubDVal(xlApp, "BA117", DtlDInp(vbTab & "250×4250" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA118", DtlDInp(vbTab & "250×3750" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA119", DtlDInp(vbTab & "250×3250" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA120", DtlDInp(vbTab & "250×2750" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA121", DtlDInp(vbTab & "250×2250" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA122", DtlDInp(vbTab & "250×1750" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA123", DtlDInp(vbTab & "250×1250" & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA124", DtlDInp(vbTab & "250× 750" & vbTab & vbTab & ": "))
        PubDModVal(xlApp, "125", "250×490×250", 1.1, DtlDInpDesc(vbTab & "250×490×250 ", "[1.1]" & vbTab))
    End Sub

    ''' <summary>
    ''' スラブ直 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="unitSlab">Slab unit.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabStr(xlApp As Application, unitSlab As Double, truck2Ton As Double)
        If unitSlab = 1 Then
            If HdrYNQ(vbTab & vbTab & "スラブ直 (D13): ") = 1 Then
                SlabStrMain(xlApp, truck2Ton)
            End If
        Else
            PrefWarn(vbTab & vbTab & "スラブ直 (D13)")
            SlabStrMain(xlApp, truck2Ton)
        End If
    End Sub

    ''' <summary>
    ''' Slab straight main.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabStrMain(xlApp As Application, truck2Ton As Double)
        If Not truck2Ton = 1 Then
            PubDVal(xlApp, "BA126", DtlDInp(vbTab & "5500: "))
            PubDVal(xlApp, "BA127", DtlDInp(vbTab & "5000: "))
        End If
        PubDVal(xlApp, "BA128", DtlDInp(vbTab & "4500: "))
        PubDVal(xlApp, "BA129", DtlDInp(vbTab & "4000: "))
        PubDVal(xlApp, "BA130", DtlDInp(vbTab & "3500: "))
        PubDVal(xlApp, "BA131", DtlDInp(vbTab & "3000: "))
        PubDVal(xlApp, "BA132", DtlDInp(vbTab & "2500: "))
        PubDVal(xlApp, "BA133", DtlDInp(vbTab & "2000: "))
        PubDVal(xlApp, "BA134", DtlDInp(vbTab & "1500: "))
        PubDVal(xlApp, "BA135", DtlDInp(vbTab & "1200: "))
        PubDVal(xlApp, "BA136", DtlDInp(vbTab & "1000: "))
        PubDVal(xlApp, "BA137", DtlDInp(vbTab & " 900: "))
    End Sub

    ''' <summary>
    ''' スラブ補強曲 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabReinfBndg(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDVal(xlApp, "BA138", DtlDInp(vbTab & "250×5250: "))
                PubDVal(xlApp, "BA139", DtlDInp(vbTab & "250×4750: "))
            End If
            PubDVal(xlApp, "BA140", DtlDInp(vbTab & "250×4250: "))
            PubDVal(xlApp, "BA141", DtlDInp(vbTab & "250×3750: "))
            PubDVal(xlApp, "BA142", DtlDInp(vbTab & "250×3250: "))
            PubDVal(xlApp, "BA143", DtlDInp(vbTab & "250×2750: "))
            PubDVal(xlApp, "BA144", DtlDInp(vbTab & "250×2250: "))
            PubDVal(xlApp, "BA145", DtlDInp(vbTab & "250×1750: "))
            PubDVal(xlApp, "BA146", DtlDInp(vbTab & "250×1250: "))
            PubDVal(xlApp, "BA147", DtlDInp(vbTab & "250× 750: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブ補強曲 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabReinfStr(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDVal(xlApp, "BA148", DtlDInp(vbTab & "5500: "))
                PubDVal(xlApp, "BA149", DtlDInp(vbTab & "5000: "))
            End If
            PubDVal(xlApp, "BA150", DtlDInp(vbTab & "4500: "))
            PubDVal(xlApp, "BA151", DtlDInp(vbTab & "4000: "))
            PubDVal(xlApp, "BA152", DtlDInp(vbTab & "3500: "))
            PubDVal(xlApp, "BA153", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "BA154", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA155", DtlDInp(vbTab & "2000: "))
            PubDVal(xlApp, "BA156", DtlDInp(vbTab & "1500: "))
            PubDVal(xlApp, "BA157", DtlDInp(vbTab & "1000: "))
        End If
    End Sub

    ''' <summary>
    ''' 副資材リスト.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Parts(xlApp As Application)
        Dim name = $"{DtlSInp(vbTab & "邸名" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")}様邸"
        DctVal(xlApp, "BJ12", name)
        CType(xlApp.ActiveSheet, Worksheet).Name = name
        PubSVal(xlApp, vbTab & "邸名コード" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "AD5")
        PubSVal(xlApp, vbTab & "納品日" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BO2")
        PubSVal(xlApp, vbTab & "住所" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BJ13")
        Dim ipp = DtlYNQ(vbTab & "運賃 (分納)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")
        If ipp = 1 Then
            DctVal(xlApp, "BA272", ipp)
        End If
        Dim curingShRingTree = DtlDInpDesc(vbTab & "養生シート輪木 (ｾｯﾄ)", vbTab & vbTab & vbTab & "[3.6×5.4]" & vbTab & vbTab)
        If curingShRingTree > 0 Then
            DctVal(xlApp, "BA249", curingShRingTree)
        Else
            DctVal(xlApp, "BA249", 1)
            ClrVal(xlApp, "BF249")
            ClrVal(xlApp, "CB249")
        End If
        PubDVal(xlApp, "BA237", DtlDInpDesc(vbTab & "フラットアンカーボルト (本)", vbTab & vbTab & "[M12×350]" & vbTab & vbTab))
        PubDVal(xlApp, "BA239", DtlDInpDesc(vbTab & "ホールダウンアンカーボルト (本)", vbTab & vbTab & "[M12×700]" & vbTab & vbTab))
        PubDVal(xlApp, "BA240", DtlDInpDesc(vbTab & "ホールダウンアンカーボルト (本)", vbTab & vbTab & "[M12×498]" & vbTab & vbTab))
        PubDVal(xlApp, "BA254", DtlDInpDesc(vbTab & "アンカーボルト M16 (本)", vbTab & vbTab & vbTab & "[M16×415]" & vbTab & vbTab))
        PubDVal(xlApp, "BA253", DtlDInpDesc(vbTab & "マグネット差筋アンカー D13 (ｾｯﾄ)", vbTab & "[直]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA252", DtlDInpDesc(vbTab & "マグネット差筋アンカー D13 (ｾｯﾄ)", vbTab & "[曲]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA248", DtlDInpDesc(vbTab & "給水用スリーブホルダー・D10用 (個)", vbTab & "[50ﾊﾟｲ]:" & vbTab & vbTab))
        PubDVal(xlApp, "BA247", DtlDInpDesc(vbTab & "排水用スリーブホルダー・D10用 (個)", vbTab & "[50ﾊﾟｲ 75ﾊﾟｲﾖｳ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA238", DtlDInpDesc(vbTab & "カットスクリュー・Ⅱ (袋)", vbTab & vbTab & "[M12ﾖｳ 50ｺｲﾘ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA257", DtlDInp(vbTab & "カットスクリュー・Ⅱ専用ピット (個)" & vbTab & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA242", DtlDInpDesc(vbTab & "グリッパーM12アンカー用 (箱)", vbTab & vbTab & "[TG1210D 50ｺ/ﾊｺ]" & vbTab))
        PubDVal(xlApp, "BA243", DtlDInpDesc(vbTab & "グリッパーM12アンカー用 (箱)", vbTab & vbTab & "[TG1213D 50ｺ/ﾊｺ]" & vbTab))
        PubDVal(xlApp, "BA244", DtlDInpDesc(vbTab & "グリッパーM12アンカー用 (箱)", vbTab & vbTab & "[TG1216D 50ｺ/ﾊｺ]" & vbTab))
        PubDVal(xlApp, "BA245", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[H60・70・80]" & vbTab & vbTab))
        PubDVal(xlApp, "BA258", DtlDInpDesc(vbTab & "グリッパーM16アンカー用 (箱)", vbTab & vbTab & "[TG1610D 100ｺ/ﾊｺ]" & vbTab))
        PubDVal(xlApp, "BA259", DtlDInpDesc(vbTab & "グリッパーM16アンカー用 (箱)", vbTab & vbTab & "[TG1616D 100ｺ/ﾊｺ]" & vbTab))
        PubDVal(xlApp, "BA250", DtlDInp(vbTab & "Ｍ型鉄筋ベース (個)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA265", DtlDInpDesc(vbTab & "偏心用鉄筋ベース (個)", vbTab & vbTab & vbTab & "[280×160×60]" & vbTab & vbTab))
        PubDVal(xlApp, "BA256", DtlDInpDesc(vbTab & "防錆巾止め金具 W160用 (本)", vbTab & vbTab & "[Fﾊﾟﾈﾙ]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA267", DtlDInpDesc(vbTab & "防錆巾止め金具 W160用 (本)", vbTab & vbTab & "[200本入]" & vbTab & vbTab))
        PubDVal(xlApp, "BA264", DtlDInpDesc(vbTab & "アンカーボルトセット (ｾｯﾄ)", vbTab & vbTab & "[M18×380]" & vbTab & vbTab))
        PubDVal(xlApp, "BA266", DtlDInpDesc(vbTab & "止水板 (巻)", vbTab & vbTab & vbTab & vbTab & "[100×20M]" & vbTab & vbTab))
        PubDVal(xlApp, "BA268", DtlDInpDesc(vbTab & "らくらく天端ビス (箱)", vbTab & vbTab & vbTab & "[500コ入]:" & vbTab & vbTab))
        PubDVal(xlApp, "BA269", DtlDInpDesc(vbTab & "結束線メッキ450 (ｹｰｽ)", vbTab & vbTab & vbTab & "[20kg]:" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA270", DtlDInpDesc(vbTab & "Ｕボルト (ｾｯﾄ)", vbTab & vbTab & vbTab & vbTab & "[M8]" & vbTab & vbTab & vbTab))
        ' Extend
        PubDVal(xlApp, "BA241", DtlDInpDesc(vbTab & "アンカーボルトセット (本)", vbTab & vbTab & "[M12×498]" & vbTab & vbTab))
        PubDVal(xlApp, "BA246", DtlDInpDesc(vbTab & "樹脂スペーサー (個)", vbTab & vbTab & vbTab & "[70×80 50ｺ/ﾊｺ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA251", DtlDInpDesc(vbTab & "樹脂スペーサー改 (ｹｰｽ)", vbTab & vbTab & vbTab & "[300ｺｲﾘ/ｹｰｽ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA255", DtlDInpDesc(vbTab & "アンカーボルトセット (本)", vbTab & vbTab & "[M16×417]" & vbTab & vbTab))
        PubDVal(xlApp, "BA260", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[H60]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA261", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & "[40・50・60]" & vbTab & vbTab))
        PubDVal(xlApp, "BA262", DtlDInpDesc(vbTab & "鉄筋スペーサー (個)", vbTab & vbTab & vbTab & "[60ﾖｳ]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA263", DtlDInpDesc(vbTab & "鉄筋スペーサー (個)", vbTab & vbTab & vbTab & "[80ﾖｳ]" & vbTab & vbTab & vbTab))
    End Sub
End Module
