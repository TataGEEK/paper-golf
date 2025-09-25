Option Explicit

'================= Nastavení =================
'Velikost hřiště
Const ROWS_COUNT As Long = 25 'Počet řádků
Const COLS_COUNT As Long = 16 'Počet sloupců

' Plocha, kterou zabírají stromy
Const TREES As Long = 4             'Počet bloků lesa
Const TREES_MIN_SIZE As Long = 10   'Počet stromů od
Const TREES_MAX_SIZE As Long = 30   'Počet stromů do

' Start a odpaliště
Const ODPALISTE As Long = 3         'Počet řádků od spodu, kde se může objevit odpaliště
Const ODPALISTE_MIN_SIZE As Long = 15 'Velikost plochy odpaliště od - do
Const ODPALISTE_MAX_SIZE As Long = 25

' Plocha, kterou zabírá písek
Const SAND As Long = 2              'Počet ploch s pískem
Const SAND_MIN_SIZE As Long = 15    'Velikost plochy písku od - do
Const SAND_MAX_SIZE As Long = 35

' Jamka a green (plocha kolem)
Const GREEN As Long = 3             'Počet řádků shora, kde se může objevit cílová jamka
Const TOP_MIN_SIZE As Long = 15     'Velikost plochy kolem jamky (greenu) od - do
Const TOP_MAX_SIZE As Long = 25

' Plocha, kterou zabírá voda
Const WATER As Long = 1             'Počet ploch s vodou
Const WATER_MIN_SIZE As Long = 15   'Velikost vodní plochy od - do
Const WATER_MAX_SIZE As Long = 35

' Sousedství pro růst (True = do 4 směrů; False = do 8 směrů, tedy i šikmo)
Const USE_FOUR_NEIGHBOR As Boolean = True

' Povinná mezera mezi různými oblastmi (Čebyševův poloměr)
Const GAP_RADIUS As Long = 2

' Počáteční zahuštění (po položení seedu) pro generování kompaktních ploch (menší číslo = větší šance na protáhlé tvary)
Const SEED_BOOST_STEPS As Long = 3
'============================================

Private Type CellRC
    r As Long
    c As Long
End Type


'================= Generování 4 hracích polí na jednu stránku =================
Public Sub GenerateFourOnOnePage()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim starts As Variant: starts = Array("A1", "R1", "A28", "R28") 'hrací pole začínají v těchto buňkách
    Dim i As Long
    
    'Zmenším okraje papíru
    With ActiveSheet.PageSetup
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With

    'Nastavím šířku sloupců a výšku řádků na 20px
    Columns("A:AG").ColumnWidth = 2.14
    Rows("1:54").RowHeight = 15

    'Pro všechny 4 části papíru vygeneruji hrací pole
    For i = LBound(starts) To UBound(starts)
        GenerateAllBlobs CStr(starts(i)), ws
    Next i

End Sub

'================= Generování hracího pole =================
Public Sub GenerateAllBlobs(Optional ByVal startAddr As String = "A1", Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = ActiveSheet
    Dim topLeft As Range: Set topLeft = ws.Range(startAddr)

    ' ID pro vytvoření blobů
    ' 0 = prázdné pole, normální hrací bod
    ' 1 = les
    ' 2 = start v odpališti
    ' 3 = celá plocha odpaliště
    ' 4 = písek
    ' 5 = cílová jamka
    ' 6 = plocha kolem cílové jamky (green)
    ' 7 = voda

    ' gridVal = co se zapíše do listu; gridGroup = ID konkrétního blobu
    Dim gridVal() As Long, gridGroup() As Long
    ReDim gridVal(1 To ROWS_COUNT, 1 To COLS_COUNT)
    ReDim gridGroup(1 To ROWS_COUNT, 1 To COLS_COUNT)

    'vyčištění obsahu listu před generováním nového obsahu hřiště
    With ws.Range(topLeft, topLeft.Offset(ROWS_COUNT - 1, COLS_COUNT - 1))
        .Value = 0
        .ClearFormats
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
    End With
    
    Randomize Timer

    '================= Generování jednitlivých ploch hřiště =================
    ' 1) Jamka a green (plocha kolem) - jamku značí číslo 5, green (plocha kolem jamky) značí číslo 6
    Call GrowSingleBlob( _
        gridVal, gridGroup, _
        rMin:=1, rMax:=Application.Min(GREEN, ROWS_COUNT), _
        cMin:=1, cMax:=COLS_COUNT, _
        seedValue:=5, growValue:=6, _
        minSize:=TOP_MIN_SIZE, maxSize:=TOP_MAX_SIZE, _
        preferCompact:=True, _
        blobId:=500)

    ' 2) Start a odpaliště - start značí číslo 2, odpaliště značí číslo 3
    Call GrowSingleBlob( _
        gridVal, gridGroup, _
        rMin:=Application.Max(1, ROWS_COUNT - ODPALISTE + 1), rMax:=ROWS_COUNT, _
        cMin:=1, cMax:=COLS_COUNT, _
        seedValue:=2, growValue:=3, _
        minSize:=ODPALISTE_MIN_SIZE, maxSize:=ODPALISTE_MAX_SIZE, _
        preferCompact:=True, _
        blobId:=200)

    ' 3) Zalesněné plochy
    Dim t As Long
    For t = 1 To TREES
        Call GrowSingleBlob( _
            gridVal, gridGroup, _
            rMin:=1, rMax:=ROWS_COUNT, _
            cMin:=1, cMax:=COLS_COUNT, _
            seedValue:=1, growValue:=1, _
            minSize:=TREES_MIN_SIZE, maxSize:=TREES_MAX_SIZE, _
            preferCompact:=False, _
            blobId:=400 + t)
    Next t
    
    ' 4) Plocha písku
    Dim s As Long
    For s = 1 To SAND
        Call GrowSingleBlob( _
            gridVal, gridGroup, _
            rMin:=1, rMax:=ROWS_COUNT, _
            cMin:=1, cMax:=COLS_COUNT, _
            seedValue:=4, growValue:=4, _
            minSize:=SAND_MIN_SIZE, maxSize:=SAND_MAX_SIZE, _
            preferCompact:=False, _
            blobId:=400 + s)
    Next s

    ' 5) Plocha vody
    Dim w As Long
    For w = 1 To WATER
        Call GrowSingleBlob( _
            gridVal, gridGroup, _
            rMin:=1, rMax:=ROWS_COUNT, _
            cMin:=1, cMax:=COLS_COUNT, _
            seedValue:=7, growValue:=7, _
            minSize:=WATER_MIN_SIZE, maxSize:=WATER_MAX_SIZE, _
            preferCompact:=False, _
            blobId:=400 + w)
    Next w

    ' Zápis do listu
    Dim arrOut() As Variant
    ReDim arrOut(1 To ROWS_COUNT, 1 To COLS_COUNT)
    Dim r As Long, c As Long
    For r = 1 To ROWS_COUNT
        For c = 1 To COLS_COUNT
            arrOut(r, c) = gridVal(r, c)
        Next c
    Next r

    ws.Range(topLeft, topLeft.Offset(ROWS_COUNT - 1, COLS_COUNT - 1)).Value = arrOut
    
    'Vygenerujeme design hracího prostředí
    Call ApplySymbolMap(ws, topLeft, ROWS_COUNT, COLS_COUNT)
    
End Sub

'================ Logika růstu =================

' preferCompact=True › kompaktní výběr hraniční buňky + seed-boost pro zahuštění.
Private Sub GrowSingleBlob(ByRef gridVal() As Long, ByRef gridGroup() As Long, _
                           ByVal rMin As Long, ByVal rMax As Long, _
                           ByVal cMin As Long, ByVal cMax As Long, _
                           ByVal seedValue As Long, ByVal growValue As Long, _
                           ByVal minSize As Long, ByVal maxSize As Long, _
                           ByVal preferCompact As Boolean, _
                           ByVal blobId As Long)

    Dim s As CellRC
    s = RandomEmptyCellWithGap(gridVal, gridGroup, rMin, rMax, cMin, cMax, blobId)
    If s.r = 0 Then Exit Sub
    If Not IsCellOkWithGap(gridVal, gridGroup, s.r, s.c, blobId) Then Exit Sub

    gridVal(s.r, s.c) = seedValue
    gridGroup(s.r, s.c) = blobId

    Dim target As Long
    target = minSize + Int(Rnd * (maxSize - minSize + 1))

    Dim front As Object
    Set front = CreateObject("Scripting.Dictionary")
    AddNeighborsToFront front, gridVal, s.r, s.c

    Dim sizeNow As Long: sizeNow = 1
    Dim safety As Long: safety = ROWS_COUNT * COLS_COUNT * 50

    ' Seed-boost pro růst okolo semínka, aby se omezily dlouhá ramena do stran
    ' Se seed-boostem: blob má hned na začátku hustější jádro, takže je kompaktnější a souvislejší.
    ' Kolik kroků takto nastartuje růst určuje konstanta SEED_BOOST_STEPS (typicky 2–5).
    If preferCompact And SEED_BOOST_STEPS > 0 Then
        Dim boost As Long: boost = SEED_BOOST_STEPS
        Dim boostFront As Object: Set boostFront = CreateObject("Scripting.Dictionary")
        AddNeighborsToFront boostFront, gridVal, s.r, s.c

        Do While boost > 0 And boostFront.Count > 0 And sizeNow < target
            Dim kkey As Variant, br As Long, bc As Long
            kkey = PickBestFrontierKey(boostFront, gridVal, gridGroup, s.r, s.c, blobId)
            If IsEmpty(kkey) Then Exit Do
            ParseKey kkey, br, bc
            boostFront.Remove kkey

            If gridVal(br, bc) = 0 Then
                If IsCellOkWithGap(gridVal, gridGroup, br, bc, blobId) Then
                    gridVal(br, bc) = growValue
                    gridGroup(br, bc) = blobId
                    sizeNow = sizeNow + 1
                    AddNeighborsToFront boostFront, gridVal, br, bc
                    AddNeighborsToFront front, gridVal, br, bc
                    boost = boost - 1
                End If
            End If
        Loop
    End If

    ' Hlavní růst
    Do While sizeNow < target And front.Count > 0 And safety > 0
        safety = safety - 1

        Dim rr As Long, cc As Long, key As Variant
        If preferCompact Then
            key = PickBestFrontierKey(front, gridVal, gridGroup, s.r, s.c, blobId)
        Else
            key = PickRandomKey(front)
        End If
        If IsEmpty(key) Then Exit Do

        ParseKey key, rr, cc
        front.Remove key

        If gridVal(rr, cc) = 0 Then
            If IsCellOkWithGap(gridVal, gridGroup, rr, cc, blobId) Then
                gridVal(rr, cc) = growValue
                gridGroup(rr, cc) = blobId
                sizeNow = sizeNow + 1
                AddNeighborsToFront front, gridVal, rr, cc
            End If
        End If
    Loop
End Sub

'================ Kompaktní výběr hraniční buňky =================

' Vybere buňku, která má nejvíce sousedů stejného blobu; při shodě je nejblíž k semínku.
Private Function PickBestFrontierKey(ByRef front As Object, ByRef gridVal() As Long, ByRef gridGroup() As Long, _
                                     ByVal seedR As Long, ByVal seedC As Long, _
                                     ByVal blobId As Long) As Variant
    Dim bestKey As Variant
    Dim bestNeighbors As Long: bestNeighbors = -1
    Dim bestDist2 As Long: bestDist2 = 0

    If front Is Nothing Or front.Count = 0 Then Exit Function

    Dim keys As Variant: keys = front.keys
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim rr As Long, cc As Long
        ParseKey keys(i), rr, cc

        ' Kandidáty filtrujeme rovnou mezerou (urychlení)
        If IsCellOkWithGap(gridVal, gridGroup, rr, cc, blobId) Then
            Dim n As Long
            n = CountSameGroupNeighbors(gridGroup, rr, cc, blobId)

            Dim d2 As Long
            d2 = (rr - seedR) * (rr - seedR) + (cc - seedC) * (cc - seedC)

            If n > bestNeighbors Or (n = bestNeighbors And d2 < bestDist2) Or bestNeighbors = -1 Then
                bestNeighbors = n
                bestDist2 = d2
                bestKey = keys(i)
            End If
        End If
    Next i

    PickBestFrontierKey = bestKey
End Function

' Počet sousedů (4/8 dle nastavení USE_FOUR_NEIGHBOR) patřících do stejného blobu (stejný groupId)
Private Function CountSameGroupNeighbors(ByRef gridGroup() As Long, ByVal r As Long, ByVal c As Long, _
                                         ByVal blobId As Long) As Long
    Dim dirs4 As Variant, dirs8 As Variant, dirs As Variant
    Dim k As Long, rr As Long, cc As Long, cnt As Long

    dirs4 = Array(Array(-1, 0), Array(1, 0), Array(0, -1), Array(0, 1))
    dirs8 = Array(Array(-1, 0), Array(1, 0), Array(0, -1), Array(0, 1), _
                  Array(-1, -1), Array(-1, 1), Array(1, -1), Array(1, 1))

    If USE_FOUR_NEIGHBOR Then
        dirs = dirs4
    Else
        dirs = dirs8
    End If

    For k = LBound(dirs) To UBound(dirs)
        rr = r + dirs(k)(0): cc = c + dirs(k)(1)
        If rr >= 1 And rr <= ROWS_COUNT And cc >= 1 And cc <= COLS_COUNT Then
            If gridGroup(rr, cc) = blobId Then cnt = cnt + 1
        End If
    Next k

    CountSameGroupNeighbors = cnt
End Function

'================ Mezera mezi blobky (GAP) =================

' Povolit obsadit [r,c], jen když v okolí Čebyševovým poloměrem GAP_RADIUS není buňka jiného blobu.
Private Function IsCellOkWithGap(ByRef gridVal() As Long, ByRef gridGroup() As Long, ByVal r As Long, ByVal c As Long, ByVal blobId As Long) As Boolean
    Dim rr As Long, cc As Long, dr As Long, dc As Long
    For dr = -GAP_RADIUS To GAP_RADIUS
        For dc = -GAP_RADIUS To GAP_RADIUS
            rr = r + dr: cc = c + dc
            If rr >= 1 And rr <= ROWS_COUNT And cc >= 1 And cc <= COLS_COUNT Then
                If gridVal(rr, cc) <> 0 Then
                    If gridGroup(rr, cc) <> blobId Then
                        IsCellOkWithGap = False
                        Exit Function
                    End If
                End If
            End If
        Next dc
    Next dr
    IsCellOkWithGap = True
End Function

' Najde prázdnou buňku splňující GAP; pokud žádná neexistuje, vrátí (0,0)
Private Function RandomEmptyCellWithGap(ByRef gridVal() As Long, ByRef gridGroup() As Long, _
                                        ByVal rMin As Long, ByVal rMax As Long, _
                                        ByVal cMin As Long, ByVal cMax As Long, _
                                        ByVal blobId As Long) As CellRC
    Dim tries As Long, r As Long, c As Long
    Dim i As Long, j As Long

    ' 1) Náhodné pokusy
    For tries = 1 To 500
        r = rMin + Int(Rnd * (rMax - rMin + 1))
        c = cMin + Int(Rnd * (cMax - cMin + 1))
        If gridVal(r, c) = 0 Then
            If IsCellOkWithGap(gridVal, gridGroup, r, c, blobId) Then
                RandomEmptyCellWithGap.r = r
                RandomEmptyCellWithGap.c = c
                Exit Function
            End If
        End If
    Next tries

    ' 2) Deterministický průchod (zamíchané pořadí)
    Dim rowsCount As Long: rowsCount = rMax - rMin + 1
    Dim colsCount As Long: colsCount = cMax - cMin + 1

    Dim rr() As Long, cc() As Long
    ReDim rr(1 To rowsCount)
    ReDim cc(1 To colsCount)
    For i = 1 To rowsCount: rr(i) = rMin + (i - 1): Next i
    For j = 1 To colsCount: cc(j) = cMin + (j - 1): Next j
    ShuffleLong rr: ShuffleLong cc

    For i = 1 To rowsCount
        For j = 1 To colsCount
            r = rr(i): c = cc(j)
            If gridVal(r, c) = 0 Then
                If IsCellOkWithGap(gridVal, gridGroup, r, c, blobId) Then
                    RandomEmptyCellWithGap.r = r
                    RandomEmptyCellWithGap.c = c
                    Exit Function
                End If
            End If
        Next j
    Next i

    ' 3) Nic nenalezeno › vrať (0,0)
    RandomEmptyCellWithGap.r = 0
    RandomEmptyCellWithGap.c = 0
End Function

'================ Pomocné funkce =================

' Do fronty přidej prázdné sousedy (0) dle 4/8-sousedství
Private Sub AddNeighborsToFront(ByRef front As Object, ByRef gridVal() As Long, ByVal r As Long, ByVal c As Long)
    Dim dirs4 As Variant, dirs8 As Variant, dirs As Variant
    Dim k As Long, rr As Long, cc As Long

    dirs4 = Array(Array(-1, 0), Array(1, 0), Array(0, -1), Array(0, 1))
    dirs8 = Array(Array(-1, 0), Array(1, 0), Array(0, -1), Array(0, 1), _
                  Array(-1, -1), Array(-1, 1), Array(1, -1), Array(1, 1))

    If USE_FOUR_NEIGHBOR Then
        dirs = dirs4
    Else
        dirs = dirs8
    End If

    For k = LBound(dirs) To UBound(dirs)
        rr = r + dirs(k)(0): cc = c + dirs(k)(1)
        If rr >= 1 And rr <= ROWS_COUNT And cc >= 1 And cc <= COLS_COUNT Then
            If gridVal(rr, cc) = 0 Then
                Dim key As String: key = MakeKey(rr, cc)
                If Not front.Exists(key) Then front.Add key, True
            End If
        End If
    Next k
End Sub

Private Function PickRandomKey(ByRef dict As Object) As Variant
    If dict Is Nothing Or dict.Count = 0 Then Exit Function
    Dim idx As Long: idx = 1 + Int(Rnd * dict.Count)
    PickRandomKey = dict.keys()(idx - 1)
End Function

' Fisher–Yates zamíchání pole Long
Private Sub ShuffleLong(ByRef a() As Long)
    Dim i As Long, j As Long, t As Long
    For i = UBound(a) To LBound(a) + 1 Step -1
        j = LBound(a) + Int(Rnd * (i - LBound(a) + 1))
        t = a(i): a(i) = a(j): a(j) = t
    Next i
End Sub

Private Function MakeKey(ByVal r As Long, ByVal c As Long) As String
    MakeKey = CStr(r) & "," & CStr(c)
End Function

Private Sub ParseKey(ByVal key As String, ByRef r As Long, ByRef c As Long)
    Dim p As Long: p = InStr(1, key, ",")
    r = CLng(Left$(key, p - 1))
    c = CLng(Mid$(key, p + 1))
End Sub

'================ Funkce pro generování designu hracího prostředí =================

Sub ApplySymbolMap(ByVal ws As Worksheet, ByVal topLeft As Range, ByVal rowsCount As Long, ByVal colsCount As Long)
    Dim r As Long, c As Long, cell As Range
    Dim rng As Range
    Set rng = ws.Range(topLeft, topLeft.Offset(rowsCount - 1, colsCount - 1))

    Application.ScreenUpdating = False

    For r = 1 To rowsCount
        For c = 1 To colsCount
            Set cell = rng.Cells(r, c)
            Select Case cell.Value
                ' hrací body
                Case 0
                    cell.NumberFormat = ChrW(&H25E6)
                    cell.Font.Name = "Times New Roman"
                
                ' les
                Case 1
                    cell.NumberFormat = "u" 'písmenko určuje tvar stromu podle použitého fontu Tree Icons
                    cell.Font.Name = "Tree Icons"
                    cell.Font.Color = RGB(0, 128, 0) 'tmavě zelená
                    With cell.Interior
                        .Pattern = xlPatternGray8
                        .PatternColorIndex = xlAutomatic
                        .PatternColor = RGB(51, 204, 51) 'zelená
                    End With
                    
                ' startovací pozice
                Case 2
                    cell.NumberFormat = ChrW(&H25CF)
                    cell.Font.Name = "Times New Roman"
                    cell.Value = 2
                    With cell.Interior
                        .Pattern = xlPatternLightDown
                        .PatternColorIndex = xlAutomatic
                        .Color = RGB(255, 255, 255)
                        .PatternColor = RGB(51, 204, 51) 'zelená
                    End With
                    
                ' plocha kolem startovací pozice
                Case 3
                    cell.NumberFormat = ChrW(&H25E6)
                    cell.Font.Name = "Times New Roman"
                    
                    With cell.Interior
                        .Pattern = xlPatternLightDown
                        .PatternColorIndex = xlAutomatic
                        .Color = RGB(255, 255, 255)
                        .PatternColor = RGB(51, 204, 51) 'zelená
                    End With
                    
                ' písek
                Case 4
                    cell.NumberFormat = ChrW(&H25E6)
                    cell.Font.Name = "Times New Roman"
                    With cell.Interior
                        .Pattern = xlPatternGray25
                        .PatternColor = RGB(255, 153, 0) 'oranžová
                    End With
            
                ' cílová jamka
                Case 5
                    cell.NumberFormat = ChrW(&H25CB)
                    cell.Font.Name = "Times New Roman"
                    
                    With cell.Interior
                        .Pattern = xlPatternLightUp
                        .PatternColorIndex = xlAutomatic
                        .Color = RGB(255, 255, 255)
                        .PatternColor = RGB(51, 204, 51) 'zelená
                    End With
            
                ' plocha kolem cílové jamky
                Case 6
                    cell.NumberFormat = ChrW(&H25E6)
                    cell.Font.Name = "Times New Roman"
                    
                    With cell.Interior
                        .Pattern = xlPatternLightUp
                        .PatternColorIndex = xlAutomatic
                        .Color = RGB(255, 255, 255)
                        .PatternColor = RGB(51, 204, 51) 'zelená
                    End With
            
                ' voda
                Case 7
                    cell.NumberFormat = ChrW(&H25E6)
                    cell.Font.Name = "Times New Roman"
                    With cell.Interior
                        .Pattern = xlPatternGray25
                        .PatternColor = RGB(51, 204, 255) 'modrá
                    End With
            
                Case Else
                    cell.NumberFormat = "General"
                    cell.Font.Name = "Calibri"
            End Select
        Next c
    Next r

    Application.ScreenUpdating = True
End Sub
