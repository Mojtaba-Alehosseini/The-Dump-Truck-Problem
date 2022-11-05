Attribute VB_Name = "Module1"
Sub RunSimulation()

    Dim tekrar As Integer
    tekrar = Cells(7, "j")
    
    ' dota dastgah bargir
    ' ye dastgah baskool
    
    Dim Loader(1 To 6) As Integer
    'Dim LoaderTime(1 To 6) As Double
    
    
    Dim LoaderOne As Integer
    Dim LoaderOneTime As Double
    
    Dim LoaderTwo As Integer
    Dim LoaderTwoTime As Double
    
    Dim BaskoolSaf(1 To 6) As Integer
    Dim Baskool As Integer
    Dim BaskoolTime As Double
    
    
    Dim Travel(1 To 6) As Integer
    Dim TravelTime As Double
    
    Dim randomdigit As Double
    
    Dim LQ As Integer
    Dim L As Integer
    Dim WQ As Integer
    Dim W As Integer
    'Counters
    Dim L1C As Integer
    Dim L2C As Integer
    Dim BC As Integer
    L1C = 0
    L2C = 0
    BC = 0
    
    
    'soorate soal
    Loader(1) = 1
    Loader(2) = 2
    Loader(3) = 3
    Loader(4) = 4
    Loader(5) = 5
    BaskoolSaf(1) = 6
    Cells(22, "e") = "LQ"
    Cells(23, "e") = "LQ"
    Cells(24, "e") = "LQ"
    Cells(25, "e") = "LQ"
    Cells(26, "e") = "LQ"
    Cells(27, "e") = "WQ"
    
    
    For clock = 0 To tekrar
    Cells(30 + clock, "c") = clock
    Cells(30 + clock, "h") = 0
    Cells(30 + clock, "i") = 0
    Cells(30 + clock, "j") = 0
        
        'az akhar be avval check mikonim yani aval travel, bad W, bad WQ, L,LQ
        
        'Teravel
        For x = 1 To 6
        If Cells(21 + x, "e") = "T" Then
                If Cells(21 + x, "h") = 1 Then 'yani teravelesh tamum shode
                    Cells(21 + x, "e") = "LQ" 'state ro avaz mikonim be safe loading
                    Cells(21 + x, "f") = "Travel Finish"
                    Cells(21 + x, "g") = "Travel Finish"
                    Cells(21 + x, "h") = "Travel Finish"
                'bayad too safe loader benevisim
                For i = 1 To 6
                If Loader(i) = 0 Then
                    Loader(i) = x
                    Exit For
                End If
                Next i
                           
                'age bishtar az 1 vahed zamani munde bud azash
                Else ' ye vahed az zamane baghimande kam mikonim o ye vahed be tey shode ezafe mikonim
                If Cells(21 + x, "h") > 1 Then
                    Cells(21 + x, "g") = Cells(21 + x, "g") + 1
                    Cells(21 + x, "h") = Cells(21 + x, "h") - 1
                End If
                End If
                
        End If
        Next x
        
        'W
        For x = 1 To 6
        If Cells(21 + x, "e") = "W" Then
                If Cells(21 + x, "h") = 1 Then 'yani tozinesh tamum shode
                    Cells(21 + x, "e") = "T" 'state ro avaz mikonim be Travel
' hala bayad w ro khali konim
                    Baskool = 0
                    Cells(30 + clock, "j") = Baskool
                    randomdigit = Rnd()
                        If randomdigit <= 0.4 Then
                            TravelTime = 40
                        Else
                        If randomdigit <= 0.7 Then
                            TravelTime = 60
                        Else
                        If randomdigit <= 0.9 Then
                            TravelTime = 80
                        Else
                        If randomdigit <= 1 Then
                            TravelTime = 100
                        End If
                        End If
                        End If
                        End If
                    Cells(21 + x, "f") = TravelTime
                    Cells(21 + x, "g") = 0
                    Cells(21 + x, "h") = TravelTime
                    Cells(21 + x, "e") = "T"
                    
                'age bishtar az 1 vahed zamani munde bud azash
                Else ' ye vahed az zamane baghimande kam mikonim o ye vahed be tey shode ezafe mikonim
                If Cells(21 + x, "h") > 1 Then
                    Cells(21 + x, "g") = Cells(21 + x, "g") + 1
                    Cells(21 + x, "h") = Cells(21 + x, "h") - 1
                    Cells(30 + clock, "j") = Baskool
                End If
                End If
            End If
        Next x
        
        'WQ
        For x = 1 To 6
        If Cells(21 + x, "e") = "WQ" Then
                If BaskoolSaf(1) = Cells(21 + x, "d") Then 'yani avale safe tozin bud
                    If Baskool = 0 Then 'yani baskool khalie
                        Baskool = x
                        Cells(30 + clock, "j") = Baskool
                        randomdigit = Rnd()
                        If randomdigit <= 0.7 Then
                            BaskoolTime = 12
                        Else
                        If randomdigit <= 1 Then
                            BaskoolTime = 16
                        End If
                        End If
                ' bayad az saf kharejesh konim va saf ro moratab konim
                For i = 1 To 5
                    BaskoolSaf(i) = BaskoolSaf(i + 1)
                Next i
                BaskoolSaf(6) = 0
                ' in che esmi bud man gozashtam khkhkh
                Cells(21 + x, "f") = BaskoolTime
                Cells(21 + x, "h") = BaskoolTime
                Cells(21 + x, "g") = 0
                Cells(21 + x, "e") = "W"
                    End If 'baskool=0
                End If ' baskoolsaf
        End If
        Next x
        
        'L
        For x = 1 To 6
        If Cells(21 + x, "e") = "L" Then
                If Cells(21 + x, "h") = 1 Then 'yani loadingesh tamum shode
                    Cells(21 + x, "e") = "WQ" 'state ro avaz mikonim
' age mizashtam =0 bad inja bayad mibordamesh too W va ye vahede zamani unja mibordamesh jolo
' hala bayad bere tahe safe WQ
                    For i = 1 To 6
                    If BaskoolSaf(i) = 0 Then
                        BaskoolSaf(i) = Cells(21 + x, "d")
                        Exit For
                    End If
                    Next i
'loader ham bayad khali beshe
                    If LoaderOne = Cells(21 + x, "d") Then
                        LoaderOne = 0
                        Cells(30 + clock, "h") = LoaderOne
                    Else
                    If LoaderTwo = Cells(21 + x, "d") Then
                        LoaderTwo = 0
                        Cells(30 + clock, "i") = LoaderTwo
                    End If
                    End If
                    
                Cells(21 + x, "f") = "In Queue"
                Cells(21 + x, "g") = "In Queue"
                Cells(21 + x, "h") = "In Queue"
                'age bishtar az 1 vahed zamani munde bud azash
                Else ' ye vahed az zamane baghimande kam mikonim o ye vahed be tey shode ezafe mikonim
                If Cells(21 + x, "h") > 1 Then
                    Cells(21 + x, "g") = Cells(21 + x, "g") + 1
                    Cells(21 + x, "h") = Cells(21 + x, "h") - 1
                    If LoaderOne = Cells(21 + x, "d") Then
                        Cells(30 + clock, "h") = LoaderOne
                    Else
                    If LoaderTwo = Cells(21 + x, "d") Then
                        Cells(30 + clock, "i") = LoaderTwo
                    End If
                    End If
                End If
                End If
        End If
        Next x
        
        'LQ
        For x = 1 To 6
        If Cells(21 + x, "e") = "LQ" Then 'age halate oon truck LQ bashe
                If Loader(1) = Cells(21 + x, "d") Then 'age oon truck avvale saf load bashe
                    If LoaderOne = 0 Then ' age loader e aval khali bashe
                        LoaderOne = x 'oon truck mire too loader e avali
                        Cells(30 + clock, "h") = LoaderOne
                        randomdigit = Rnd()
                        If randomdigit <= 0.3 Then
                            LoaderOneTime = 5
                        Else
                        If randomdigit <= 0.8 Then
                            LoaderOneTime = 10
                        Else
                        If randomdigit <= 1 Then
                            LoaderOneTime = 15
                
                        End If
                        End If
                        End If
                ' bayad az saf kharejesh konim va saf ro moratab konim
                For i = 1 To 5
                    Loader(i) = Loader(i + 1)
                Next i
                Loader(6) = 0
                Cells(21 + x, "f") = LoaderOneTime
                Cells(21 + x, "h") = LoaderOneTime
                Cells(21 + x, "g") = 0
                Cells(21 + x, "e") = "L"
                    Else
                    If LoaderTwo = 0 Then
                        LoaderTwo = x
                        Cells(30 + clock, "i") = LoaderTwo
                        randomdigit = Rnd()
                        If randomdigit <= 0.3 Then
                            LoaderTwoTime = 5
                        Else
                        If randomdigit <= 0.8 Then
                            LoaderTwoTime = 10
                        Else
                        If randomdigit <= 1 Then
                            LoaderTwoTime = 15
                
                        End If
                        End If
                        End If
                    For i = 1 To 5
                    Loader(i) = Loader(i + 1)
                    Next i
                Loader(6) = 0
                Cells(21 + x, "f") = LoaderTwoTime
                Cells(21 + x, "h") = LoaderTwoTime
                Cells(21 + x, "g") = 0
                Cells(21 + x, "e") = "L"
                    End If
                    End If
                End If
        End If
        Next x
        
        LQ = 0
        L = 0
        WQ = 0
        W = 0
        
        For x = 1 To 6
        If Cells(21 + x, "e") = "LQ" Then
            LQ = LQ + 1
        Else
        If Cells(21 + x, "e") = "L" Then
            L = L + 1
        Else
        If Cells(21 + x, "e") = "WQ" Then
            WQ = WQ + 1
        Else
        If Cells(21 + x, "e") = "W" Then
            W = W + 1
        If Cells(21 + x, "e") = "T" Then
            Travel(x) = 1
        End If
        End If
        End If
        End If
        End If
        Next x
        
        
        Cells(30 + clock, "d") = LQ
        Cells(30 + clock, "e") = L
        Cells(30 + clock, "f") = WQ
        Cells(30 + clock, "g") = W
        
        'safe bargiri o tozin o safar
        'Range(Cells(30 + clock, "k")) = Join(Loader, ",")
        'Cells(30 + clock, "k") = Join(Loader, ",")
        'Cells(30 + clock, "l") = Join(BaskoolSaf, ",")
        'Cells(30 + clock, "m") = Join(Travel, ",")
       
    If Cells(30 + clock, "h") <> 0 Then
        L1C = L1C + 1
    End If
    If Cells(30 + clock, "i") <> 0 Then
        L2C = L2C + 1
    End If
    If Cells(30 + clock, "j") <> 0 Then
        BC = BC + 1
    End If
        
        
    Next clock
    
    Cells(23, "j") = (L1C / tekrar) * 100
    Cells(23, "k") = (L2C / tekrar) * 100
    Cells(23, "l") = (BC / tekrar) * 100
        
    
    

End Sub
Sub Clear()
Range("e22", "h27").Clear
Range("c30", "j10000").Clear
End Sub
