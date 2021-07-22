
Imports VB = Microsoft.VisualBasic

Module Generale
    Private Sub WriteLog(ByVal PercorsoFile As String, ByVal NomeFile As String, ByVal TextLog As String)
        Dim NumfileLog As Integer

        Try
            NumfileLog = FreeFile()
            FileOpen(NumfileLog, (PercorsoFile & "LOG_" & NomeFile), OpenMode.Output)
            PrintLine(NumfileLog, Format(Now, "yyyy-MM-dd hh:mm:ss") & " - " & TextLog)
        Catch
        Finally
            'chiusura del file
            FileClose(NumfileLog)
        End Try
    End Sub

    Function Formatta(ByVal allign As Integer, ByVal tipocampo As Integer, ByVal lunghcampo As Integer, ByVal stringa As String) As String

        Dim piena As String

        stringa = Replace(stringa, Chr(13), "")
        stringa = Replace(stringa, Chr(10), "")

        'Vedo che tipo di riempimento devo gestire (tipocampo = 1 0, tipocampo = 0 " ")
        If (tipocampo = 1) Then
            Dim a As Char
            a = "0"c
            'piena = String$(lunghcampo, "0")   'numerico
            piena = New String(a, lunghcampo)
        Else
            Dim a As Char
            a = " "c
            'piena = String(lunghcampo, " ")   'stringa
            piena = New String(a, lunghcampo)
        End If

        'Vedo che tipo di allineamento devo gestire (allign = 1 DX, allign = 0 SX)
        If (allign = 1) Then
            Formatta = Right(piena + Trim$(stringa), lunghcampo)   'numerico
        Else
            Formatta = Left((Trim$(stringa) + piena), lunghcampo)   'stringa
        End If

    End Function
    Public Structure ObjFlussiScartati
        Dim sNomeFile As String
    End Structure

    Public Function VerificaFile(ByRef PathFile As String, ByRef NomeFile As String, ByRef ObjScartati() As ObjFlussiScartati, ByRef Errstring As String) As String
        Dim TrovatoG1 As Boolean
        Dim NumeroA1, NumeroZ1 As Integer
        Dim DateScarto As String
        Dim AcquisireICI As Boolean
        Dim NumFile As Integer
        Dim Text_Line As String
        Dim DataAccredito As String
        'Dim nfile, i As Integer
        Dim Reader As New LettoreFile

        'Dim ImportoDaFlusso As Double
        'Dim Importo_Versamento As Double
        'Dim TotaleImportiEuro As Double
        'Dim TotaleImportiEuroFlusso As Double
        'Dim Divisa As String
        Dim messaggio As String = ""


        Dim TotFlussiScartati As Integer

        Try
            WriteLog(PathFile, NomeFile, "inizio verifica")

            TotFlussiScartati = 0
            'nfile = CInt(Reader.ReadFileFromFolder(PathFile, -1))
            'If nfile = 0 Then

            'Else
            ''''se passo i>=0 prelevo i nomi dei file
            'i = 0
            'Do While i < nfile
            'NomeFile = Reader.ReadFileFromFolder(PathFile, i)
            AcquisireICI = False : DateScarto = ""
            NumeroA1 = 0 : TrovatoG1 = False : NumeroZ1 = 0

            'controllo che il file sia di un singolo ente (A1/G1/Z1)
            'e che abbia almeno un record di versamento per tenerne traccia
            NumFile = FreeFile()
            FileOpen(NumFile, PathFile & NomeFile, OpenMode.Input)
            Do While Not EOF(NumFile)
                Text_Line = LineInput(NumFile)
                If Left(Text_Line, 2) = "A1" Then
                    NumeroA1 = NumeroA1 + 1

                ElseIf Left(Text_Line, 2) = "G1" Then
                    TrovatoG1 = True

                    '*** Fabi 18022008
                    'If IsNumeric(Trim(Mid(Text_Line, 96, 15))) Then
                    '    Importo_Versamento = CDbl(Trim(Mid(Text_Line, 96, 15)))
                    '    'calcolo il totale dei versamenti
                    '    If UCase(Divisa) = "E" Then
                    '        TotaleImportiEuro += CDbl(CentInEuro(Importo_Versamento, 2))
                    '        TotaleImportiEuroFlusso += CDbl(CentInEuro(Importo_Versamento, 2))
                    '    Else
                    '        TotaleImportiEuro += CDbl(Importo_Versamento)
                    '        TotaleImportiEuroFlusso += CDbl(Importo_Versamento)
                    '    End If
                    '    If Importo_Versamento = 0 Then
                    '        'incremento la variabile dei record scartati
                    '        'NScarti = NScarti + 1
                    '    End If
                    'Else
                    '    'incremento la variabile dei record scartati
                    '    'NScarti = NScarti + 1
                    'End If
                    '*** fine Fabi 18022008

                ElseIf Left(Text_Line, 2) = "Z1" Then
                    NumeroZ1 = NumeroZ1 + 1
                ElseIf Left(Text_Line, 2) = "G2" Then

                    ''*** Fabi 18022008
                    'If IsNumeric(Trim(Mid(Text_Line, 46, 15))) Then
                    '    ImportoDaFlusso += CDbl(Trim(Mid(Text_Line, 46, 15)))
                    'End If
                    ''*** fine Fabi 18022008

                End If
            Loop
            FileClose(NumFile)
            WriteLog(PathFile, NomeFile, "fine conteggio")

            If NumeroA1 = 1 And NumeroZ1 = 1 And TrovatoG1 = True Then
                'il flusso è ok
                AcquisireICI = True
            Else
                AcquisireICI = False
            End If

            ''*** Fabi 18022008
            'If CStr(Format(CDbl(CentInEuro(ImportoDaFlusso, 2)), "#,##0.00")) <> CStr(Format(CDbl(CentInEuro(TotaleImportiEuroFlusso, 2)), "#,##0.00")) Then
            '    messaggio = "Nel Flusso analizzato i totali acquisiti non tornano con i totali del record G2"
            'End If

            If AcquisireICI = True Then
                'controllo che tutti i record abbiano lunghezza 300
                NumFile = FreeFile()
                FileOpen(NumFile, PathFile & NomeFile, OpenMode.Input)
                Do While Not EOF(NumFile)
                    Text_Line = LineInput(NumFile)
                    If Len(Text_Line) <> 300 Then
                        AcquisireICI = False
                        Exit Do
                    End If
                Loop
                FileClose(NumFile)
                WriteLog(PathFile, NomeFile, "fine controllo lunghezza")

                If AcquisireICI = True Then
                    NumFile = FreeFile()
                    FileOpen(NumFile, PathFile & NomeFile, OpenMode.Input)
                    Do While Not EOF(NumFile)
                        Text_Line = LineInput(NumFile)
                        If Left(Text_Line, 2) = "G1" Then
                            'prelevo la data di valuta
                            WriteLog(PathFile, NomeFile, "devo prelevare data")
                            DataAccredito = Trim(Mid(Text_Line, 23, 8))
                            WriteLog(PathFile, NomeFile, "passo a controlla_data dataaccredito::" & DataAccredito & "::" & DataAccredito.Substring(6, 2) & "/" & DataAccredito.Substring(4, 2) & "/" & DataAccredito.Substring(0, 4))

                            'DataAccredito = Right(DataAccredito, 2) & Mid(DataAccredito, 5, 2) & Left(DataAccredito, 4)
                            If controlla_data(DataAccredito.Substring(6, 2) & "/" & DataAccredito.Substring(4, 2) & "/" & DataAccredito.Substring(0, 4)) = True Then
                                DateScarto = DateScarto & " " & DataAccredito.Substring(6, 2) & "/" & DataAccredito.Substring(4, 2) & "/" & DataAccredito.Substring(0, 4)
                            Else
                                'If DateDiff("d", Format(Now, "MM/dd/yyyy"), Format(CDate(DataAccredito.Substring(4, 2) & "/" & DataAccredito.Substring(6, 2) & "/" & DataAccredito.Substring(0, 4)), "MM/dd/yyyy")) > 0 Then
                                '    'If DateDiff("d", Format(Now, "MM/dd/yyyy"), DataAccredito.Substring(4, 2) & "/" & DataAccredito.Substring(6, 2) & "/" & DataAccredito.Substring(0, 4)) > 0 Then
                                '    DateScarto = DateScarto & " " & Left(DataAccredito, 2) & "/" & Mid(DataAccredito, 3, 2) & "/" & Right(DataAccredito, 4)
                                'End If
                            End If
                        End If
                    Loop
                    FileClose(NumFile)

                    If DateScarto <> "" Then
                        messaggio = "Il flusso " & NomeFile & " che si sta acquisendo ha le seguenti Date di Accredito non valide: " & DateScarto & "\n Si vuole scartare la fornitura?"
                        ReDim Preserve ObjScartati(TotFlussiScartati)
                        ObjScartati(TotFlussiScartati).sNomeFile = NomeFile
                        TotFlussiScartati += 1
                    End If
                Else
                    ReDim Preserve ObjScartati(TotFlussiScartati)
                    ObjScartati(TotFlussiScartati).sNomeFile = NomeFile
                    TotFlussiScartati += 1

                End If
            Else
                ReDim Preserve ObjScartati(TotFlussiScartati)
                ObjScartati(TotFlussiScartati).sNomeFile = NomeFile
                TotFlussiScartati += 1
            End If
            'i += 1
            'Loop

            'End If
        Catch ex As Exception
            'MessageBox.Show("Errore durante la fase di verifica del file " & NomeFile, "VERIFICA FILE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Errstring = "Errore durante la fase di verifica del file " & NomeFile & "::" & ex.Message
            FileClose(NumFile)
            messaggio = Errstring
        End Try
        Return messaggio
    End Function

    Public Function controlla_data(ByRef datacontrollo As String) As Boolean
        'SUB CHE CONTROLLA SE E' STATA INSERITA UNA DATA CORRETTA
        Dim controllo_data As Object

        On Error GoTo controllodata

        Dim Mese, Giorno, Anno As Integer
        Dim bisestile As Integer

        controllo_data = False

        Giorno = CInt(Microsoft.VisualBasic.Left(datacontrollo, 2))
        Mese = CInt(Mid(datacontrollo, 4, 2))
        Anno = CInt(Mid(datacontrollo, 7, 4))
        If Len(Anno) = 4 Then
            bisestile = CInt(Anno) Mod 4
        Else
            GoTo controllodata
        End If

        'controllo del giorno
        If Mese = 2 And bisestile = 0 Then 'controllo giorni di feb. quando anno bisestile
            If Giorno < 1 Or Giorno > 29 Then
                GoTo controllodata
            End If
        ElseIf Mese = 2 And bisestile <> 0 Then  'controllo giorni di feb. quando anno non bisestile
            If Giorno < 1 Or Giorno > 28 Then
                GoTo controllodata
            End If
        ElseIf Mese = 11 Or Mese = 4 Or Mese = 6 Or Mese = 9 Then  'controllo giorni se il mese ne deve avere 30
            If Giorno < 1 Or Giorno > 30 Then
                GoTo controllodata
            End If
        ElseIf Mese <> 11 And Mese <> 4 And Mese <> 6 And Mese <> 9 Then  'altri mesi
            If Giorno < 1 Or Giorno > 31 Then
                GoTo controllodata
            End If
        End If

        'controllo mese
        If Mese < 1 Or Mese > 12 Then
            GoTo controllodata
        End If

        Exit Function

controllodata:
        controlla_data = True
        Exit Function

    End Function
    Public Function CentInEuro(ByVal Importo As String, ByVal Decimali As Integer) As String

        CentInEuro = Importo / CLng(1 & StrDup(Decimali, "0"))

        CentInEuro = FormattaValuta("E", CentInEuro)

    End Function

    Public Function FormattaValuta(ByVal TypeFormat As String, ByVal ValutaFormatta As String) As String
        'TypeFormat può valere E = euro, L = lire
        FormattaValuta = ""
        If TypeFormat = "E" Then
            FormattaValuta = CStr(Format(CDbl(ValutaFormatta), "#,##0.00"))
        ElseIf TypeFormat = "L" Then
            FormattaValuta = CStr(Format(CDbl(ValutaFormatta), "#,##0"))
        End If
    End Function

    Public Function LeggiPath(ByVal Chiave As String) As String
        Dim AppReader As New System.Configuration.AppSettingsReader
        Dim Valore As String

        Valore = CType(AppReader.GetValue(Chiave, GetType(String)), String)

        Return Valore
    End Function

    Public Sub CheckCreateDir(ByVal PathCartella As String, ByVal NomeCartella As String)
        Dim DirFound As String
        Dim ExistDir As Integer

        'controllo se la cartella esiste al percorso dichiarato altrimenti la(creo)

        ExistDir = 0
        DirFound = Dir(PathCartella, vbDirectory)   'Retrieve the firstentry.
        Do While DirFound <> ""   'Start the loop.
            'Use bitwise comparison to make sure NomeCartella is a directory.
            If (GetAttr(PathCartella & DirFound) And vbDirectory) = vbDirectory Then
                'If DirFound = Replace(NomeCartella, "\", "") Then
                If DirFound = NomeCartella Then
                    ExistDir = 1
                    Exit Do
                End If
            End If
            DirFound = Dir()   'Get next entry.
        Loop

        If ExistDir = 0 Then
            MkDir(PathCartella & NomeCartella & "\")
        End If
    End Sub

End Module
