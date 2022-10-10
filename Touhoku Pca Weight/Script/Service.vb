Imports Microsoft.Office.Interop.Excel

Friend Module Service
    ''' <summary>
    ''' Weight Touhoku Pca.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub WtTouhokuPca(xlApp As Application)
        ' Fare
        Dim truck2Ton = HdrYNQ(vbTab & vbTab & "運賃 (2トン車): ")
        Fare(xlApp, truck2Ton)
        ' Unit GL-300
        Dim gl300 = HdrYNQ(vbTab & vbTab & "外周深GL-300: ")
        Unit300(xlApp, gl300)
        ' Unit GL-150
        PrefWarn(vbTab & vbTab & "外周/内周GL-150")
        Unit150(xlApp)
        ' Unit GL-300/+30
        Unit300Cut(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-300/+30: "))
        ' Unit GL-400
        Unit400(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-400: "))
        ' Unit GL-500
        Unit500(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-500: "))
        ' Unit GL-500/+30
        Unit500Cut(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-500/+30: "))
        ' Unit GL-600
        Unit600(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-600: "))
        ' Unit GL-700
        Unit700(xlApp, HdrYNQ(vbTab & vbTab & "外周深GL-700: "))
        ' Unit Garage GL-300
        Unit300Gar(xlApp, HdrYNQ(vbTab & vbTab & "ガレージ外周GL-300: "))
        ' Slab unit
        Dim unitSlab = HdrYNQ(vbTab & vbTab & "スラブユニット: ")
        SlabUnit(xlApp, unitSlab)
        ' Electric water heater
        ElecWtrHtr(xlApp, HdrDInp(vbTab & vbTab & "電気温水器: "))
        ' Sleeve
        Sleeve(xlApp, gl300, unitSlab)
        ' Straight joint
        PrefWarn(vbTab & vbTab & "ストレート")
        JtStr(xlApp)
        ' Corner joint
        PrefWarn(vbTab & vbTab & "コーナー")
        JtCor(xlApp)
        ' Long corner
        LongCor(xlApp, HdrYNQ(vbTab & vbTab & "ロングコーナー: "))
        ' Edge
        Edge(xlApp, HdrYNQ(vbTab & vbTab & "端部 (D16): "))
        ' Crank
        Crank(xlApp, HdrYNQ(vbTab & vbTab & "クランク: "))
        ' U type
        PubDVal(xlApp, "BA182", HdrDInp(vbTab & vbTab & "コ型 (D16[750×920×750]): "))
        ' Island
        Island(xlApp, HdrYNQ(vbTab & vbTab & "島 (D16): "))
        ' Straight
        Straight(xlApp, HdrYNQ(vbTab & vbTab & "ストレート (D16): "))
        ' Cap tire
        CapTire(xlApp, HdrYNQ(vbTab & vbTab & "キャップタイヤ (320): "))
        ' Haunch
        Haunch(xlApp, HdrYNQ(vbTab & vbTab & "ハンチ (D16): "))
        ' Corner 135 degree
        Corner135deg(xlApp, HdrYNQ(vbTab & vbTab & "コーナー(曲 135°): "))
        ' Hook
        PrefWarn(vbTab & vbTab & "フック (D10)")
        Hook(xlApp)
        ' Corner 3D
        Corner3d(xlApp, HdrYNQ(vbTab & vbTab & "コーナー3 (D16): "))
        ' Crank 3D
        Crank3d(xlApp, HdrYNQ(vbTab & vbTab & "クランク3 (D16): "))
        ' U type 3D
        PubDModVal(xlApp, "198", "（コノ字３右）", "750×920×460×390", 4.1, HdrDInpDesc(vbTab & vbTab & "コ型3右 (D16[750×920×460×390]) ", "[4.1]"))
        ' M type
        MType(xlApp, HdrYNQ(vbTab & vbTab & "M型 (D16): "))
        ' Bending 135 degree
        PubDModVal(xlApp, "213", "750×曲（H80）×460×750", 3.4, HdrDInpDesc(vbTab & vbTab & "曲 (135°[D16{750×H80×460×750}]) ", "[3.4]"))
        ' Main reinforcement
        MainReinf(xlApp, HdrYNQ(vbTab & vbTab & "主筋補強 (D10): "))
        ' Slab bending
        SlabBndg(xlApp, unitSlab, truck2Ton)
        ' Slab straight
        SlabStr(xlApp, unitSlab, truck2Ton)
        ' Slab reinforcement bending
        SlabReinfBndg(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強曲 (D10): "), truck2Ton)
        ' Slab reinforcement straight
        SlabReinfStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強直 (D10): "), truck2Ton)
        ' Parts
        PrefWarn(vbTab & vbTab & "副資材リスト")
        Parts(xlApp)
    End Sub
End Module
