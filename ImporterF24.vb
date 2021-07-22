Imports System.Data.SqlClient
Imports ImportazioneF24.Generale
Imports System.Configuration
Imports Utility

Public Class ImporterF24
    'Private NomeFile As String
    'Private PercorsoFile As String
    Private MsgCaption As String = "ANALISI FLUSSO"
    Private PercorsoFileEstrazione As String
    Private DBType As String = "SQL"

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

    Public Function AvviaImportazione(ByVal PercorsoFile As String, ByVal NomeFile As String, ByVal connessione As String, ByVal operatore As String, ByRef messaggio As String) As String
        Dim NrcFile As Integer
        Dim NrcScarti As Integer
        Dim NrcFlusso As Integer
        Dim TotaliImportiFile As Double
        Dim ErroreVerifica As String
        Dim ObjFileScartati() As ObjFlussiScartati
        Dim TotFileScartati As Integer
        Dim PathFileScartati As String
        Dim acquisizione As String = ""

        WriteLog(PercorsoFile, NomeFile, "inizio import")
        If InStr(PercorsoFile + NomeFile, ".") <> 0 Then
            ObjFileScartati = Nothing
            ErroreVerifica = ""
            VerificaFile(PercorsoFile, NomeFile, ObjFileScartati, ErroreVerifica)
            If ErroreVerifica = "" Then
                If Not ObjFileScartati Is Nothing Then
                    If ObjFileScartati.Length > 0 Then
                        CheckCreateDir(PercorsoFile, "FILE_SCARTATI")
                        PathFileScartati = PercorsoFile & "FILE_SCARTATI\"
                        For TotFileScartati = 0 To ObjFileScartati.Length - 1
                            FileCopy(PercorsoFile & ObjFileScartati(TotFileScartati).sNomeFile, PathFileScartati & ObjFileScartati(TotFileScartati).sNomeFile)
                            Kill(PercorsoFile & ObjFileScartati(TotFileScartati).sNomeFile)
                        Next
                        acquisizione = "s"
                        messaggio += "Alcuni file sono stati scartati al percorso: \n" & PathFileScartati & " e non verranno acquisiti dalla procedura"
                    Else
                        acquisizione = ""
                    End If
                End If

                If acquisizione <> "s" Then
                    WriteLog(PercorsoFile, NomeFile, "inizio acquisizione massiva")
                    If AcquisizioneFileF24massiva(PercorsoFile, NomeFile, connessione, NrcFile, NrcScarti, NrcFlusso, TotaliImportiFile, operatore, ErroreVerifica) = True Then
                        'sposto il file nella cartella acquisiti
                        CheckCreateDir(PercorsoFile, "ACQUISITI")
                        'controllo se il file esiste già
                        If System.IO.File.Exists(PercorsoFile & "ACQUISITI\" & NomeFile) Then
                            'se si lo rinomino
                            System.IO.File.Copy(PercorsoFile & "ACQUISITI\" & NomeFile, PercorsoFile & "ACQUISITI\" & NomeFile & "." & Format(Now, "yyyyMMdd_hhmmss"))
                            System.IO.File.Delete(PercorsoFile & "ACQUISITI\" & NomeFile)
                        End If
                        System.IO.File.Move(PercorsoFile & NomeFile, PercorsoFile & "ACQUISITI\" & NomeFile)
                        messaggio += "\nAcquisizione terminata con successo!"
                        If ErroreVerifica <> "" Then
                            messaggio += "\n" & ErroreVerifica
                        End If
                        acquisizione = "ok"
                        'DEVO AGGIORNARE I FLAG DATI ERRATI SU TUTTI I G1 DELLO STESSO CF_PIVA, ANNO, ACCONTO E\0 SALDO, DATA_ACCREDITO
                        AggiornaFlagDatiErrati(connessione)
                        'DEVO AGGIORNARE I FLAG ANNO ERRATO SU TUTTI I G1 DELLO STESSO CF_PIVA, ACCONTO E\0 SALDO, DATA_ACCREDITO
                        AggiornaFlagAnnoErrato(connessione)
                        WriteLog(PercorsoFile, NomeFile, "fine import OK")
                    Else
                        WriteLog(PercorsoFile, NomeFile, "Acquisizione da AcquisizioneFileF24massiva interrotta a causa di anomalia::" & ErroreVerifica)
                        messaggio += "\nAcquisizione interrotta a causa di anomalia!\n" & ErroreVerifica
                        acquisizione = "err"
                    End If
                End If
            Else
                WriteLog(PercorsoFile, NomeFile, "Verifica interrotta a causa di anomalia::" & ErroreVerifica)
                acquisizione = "err"
                messaggio += "\nVerifica interrotta a causa di anomalia!\n" & ErroreVerifica
            End If
        End If

        Return acquisizione
    End Function

    Public Function AcquisizioneFileF24massiva(ByRef PathFile As String, ByRef NomeFile As String, ByVal strconn As String, ByRef NRcFile As Integer, ByRef NScarti As Integer, ByRef NRcFlusso As Integer, ByRef TotImportiFile As Double, ByVal operatore As String, ByRef sMyErr As String) As Boolean
        Dim NumFileLOG As Integer
        Dim CodiceFiscale As String
        Dim Cognome, Nome As String
        Dim DataAccredito As String
        Dim DataVersamento As String
        Dim Tipo_Pagamento As String
        Dim Anno_Riferimento As String
        Dim Num_Fab As Integer
        Dim Importo_Versamento As Double
        Dim Imp_Ter_Agr As Double
        Dim Imp_Ter_Fab As Double
        Dim Imp_Altri_Fab As Double
        Dim Imp_Abi_Prin As Double
        Dim Detrazione As Double
        Dim CapEnte As String
        Dim Bollettino_EXRurale As String
        Dim Bollettino_ICI As String
        Dim Flag_Rav_Operoso As Integer
        Dim FlagCFerrato As Integer
        Dim FlagTributoErrato As Integer
        Dim FlagAnnoErrato As Integer
        Dim FlagAcconto As Integer
        Dim FlagSaldo As Integer
        Dim FlagAccontoSaldo As Integer
        Dim FlagImmVariati As Integer
        Dim FlagDatiICIerrati As Integer
        Dim DataNascita As String
        Dim ComuneNascita As String
        Dim PvNascita As String
        Dim NumFile As Short
        Dim Text_Line As String
        Dim Concessione As String
        Dim Ente As String
        Dim Data_Inizio As String
        Dim Data_Fine As String
        Dim TotPagamenti As Double
        Dim CodEnte As String
        Dim CodFlusso As Short
        Dim Divisa As String
        Dim TotaleImportiEuro As Double
        Dim TotaleImportiEuroFlusso As Double
        Dim ProgPagamento As Double
        Dim TipoRecord As String
        Dim SQLInsert As String
        Dim SQLValues As String
        Dim Provenienza As String
        Dim DataCreazioneFlusso As String
        Dim TipoSoggetto As String
        Dim Tipo_Bollettino As String
        Dim TotImpCodeLine As Double
        Dim RcFinaleAcquisito As Short
        Dim CodTributo As String
        Dim IdentificativoA1 As String
        Dim Reader As New LettoreFile
        Dim PathScartiAcquisizione As String = ""
        '*** 20140109 - import F24 altri tributi ***
        Dim myKey As String = ""
        Dim IdOperazione As String
        '*** ***
        Dim objconn As New SqlConnection(strconn)
        'Dim objtrans As SqlTransaction
        Dim CmdDati As SqlClient.SqlCommand
        Dim ImpAccreditoDispostoEnte, ImpAccreditoDispostoIFEL As Double
        Dim k As Integer = 0
        Dim contaG1 As Integer = 0
        Dim contaG1Acq As Integer = 0

        'Dim ObjTabF24() As ObjTabellaF24

        Try
            CheckCreateDir(PathFile, "SCARTI_ACQUISIZIONE")

            PathScartiAcquisizione += "SCARTI_ACQUISIZIONE" & "\"

            objconn.Open()

            'CmdDati = New SqlClient.SqlCommand("DELETE FROM TAB_ACQUISIZIONE_F24", objconn)
            'CmdDati.ExecuteNonQuery()
            'CmdDati.Dispose()
            'CmdDati = Nothing

            'nfile = CInt(Reader.ReadFileFromFolder(PathFile, -1))
            'If nfile = 0 Then
            'Else
            '    ''''se passo i>=0 prelevo i nomi dei file
            '    i = 0
            '    Do While i < nfile
            'NomeFile = Reader.ReadFileFromFolder(PathFile, i)
            AcquisizioneFileF24massiva = False

            'ripulisco le variabili
            Concessione = "" : Ente = "" : Data_Inizio = "" : Data_Fine = "" : TotPagamenti = 0
            Anno_Riferimento = "" : CodEnte = "" : CodFlusso = 0 : Divisa = ""

            'creazione del file di LOG relativo al file che si sta acquisendo nella cartella 'Acquisiti'
            NumFileLOG = FreeFile()
            FileOpen(NumFileLOG, (PathFile & PathScartiAcquisizione & "SCARTI_" & NomeFile), OpenMode.Output)
            'FileOpen(NumFileLOG, (PathScartiAcquisizione & "SCARTI_" & NomeFile), OpenMode.Output)

            NScarti = 0

            NRcFile = 0
            'calcolo il numero di record del file per animare la progress bar del file
            NumFile = FreeFile()
            FileOpen(NumFile, PathFile & NomeFile, OpenMode.Input)
            Do While Not EOF(NumFile)
                Text_Line = LineInput(NumFile)
                NRcFile = NRcFile + 1
            Loop
            FileClose(NumFile)
            Text_Line = ""

            'apro il file
            FileOpen(NumFile, PathFile & NomeFile, OpenMode.Input)
            Dim CodBelfiore As String

            Do While Not EOF(NumFile)
                'leggo il tracciato
                Text_Line = LineInput(NumFile)
                'prelevo il primo carattere
                TipoRecord = Mid(Text_Line, 1, 2)
                'controllo su che tipo di record sono
                Select Case TipoRecord
                    'sono sul record di testa dell'Ente
                    Case "A1"
                        'controllo se devo registrare il record di DESCRIZIONE_FLUSSI_ICI
                        'setto la variabile di controllo acquisizione record di coda a false
                        RcFinaleAcquisito = 0
                        'valorizzo le variabili di totalizzazione importi e record dei flussi
                        NRcFlusso = NRcFlusso + TotPagamenti
                        TotImportiFile = TotImportiFile + TotaleImportiEuro

                        'ripulisco le variabili
                        Concessione = "" : Ente = "" : Data_Inizio = "" : Data_Fine = "" : DataAccredito = "" : TotPagamenti = 0
                        Anno_Riferimento = "" : CodEnte = "" : CodFlusso = 0 : Divisa = "" : TotaleImportiEuro = 0 : TotImpCodeLine = 0
                        'prelevo il N° di conto corrente del concessionario
                        CodBelfiore = Trim(Mid(Text_Line, 44, 4))
                        'prelevo l'identificativo del file
                        IdentificativoA1 = Trim(Mid(Text_Line, 56, 24))

                        'prelevo la data di fornitura
                        DataCreazioneFlusso = Trim(Mid(Text_Line, 3, 8))
                        If controlla_data(CStr(DataCreazioneFlusso).Substring(6, 2) & "/" & CStr(DataCreazioneFlusso).Substring(4, 2) & "/" & CStr(DataCreazioneFlusso).Substring(0, 4)) = True Then
                            PrintLine(NumFileLOG, Text_Line)
                            'incremento la variabile dei record scartati
                            NScarti = NScarti + 1
                            GoTo RecordSuccessivo
                        End If

                        'prelevo la provenienza del flusso
                        Provenienza = "F24"
                        'prelevo la divisa del tracciato
                        Divisa = Trim(Mid(Text_Line, 41, 1))

                        'azzero il progressivo del pagamento
                        ProgPagamento = 0

                        'sono sul un record di versamento
                    Case "G1"
                        contaG1 = contaG1 + 1
                        'incremento il progressivo pagamento
                        ProgPagamento = ProgPagamento + 1
                        'ripulisco le variabili
                        CodiceFiscale = "" : Cognome = "" : Nome = "" : DataAccredito = "" : DataVersamento = "" : Tipo_Pagamento = "" : Anno_Riferimento = "" : TipoSoggetto = ""
                        Num_Fab = 0 : Importo_Versamento = 0 : Imp_Ter_Agr = 0 : Imp_Ter_Fab = 0 : Imp_Altri_Fab = 0 : Imp_Abi_Prin = 0 : Detrazione = 0 : CapEnte = ""
                        Divisa = "" : Flag_Rav_Operoso = 0 : FlagAcconto = 0 : FlagSaldo = 0 : FlagAccontoSaldo = 0 : FlagCFerrato = 0 : FlagImmVariati = 0 : FlagAnnoErrato = 0 : FlagDatiICIerrati = 0 : Tipo_Bollettino = "" : Bollettino_ICI = "" : Divisa = ""
                        '*** 20140109 - import F24 altri tributi ***
                        myKey = ""
                        IdOperazione = ""
                        '*** ***
                        'prelevo la provenienza del flusso
                        Provenienza = "F24"
                        myKey = Text_Line.Substring(12, 31)
                        'prelevo la data di accredito
                        DataAccredito = Trim(Mid(Text_Line, 13, 8)) '***20081216 - cambiato data in data di ripartizione***Trim(Mid(Text_Line, 23, 8))
                        'prelevo il CF o partita IVA
                        CodiceFiscale = Trim(Mid(Text_Line, 50, 16))
                        'prelevo il flag_cf_errato
                        FlagCFerrato = Trim(Mid(Text_Line, 66, 1))
                        'prelevo la data di pagamento
                        DataVersamento = Trim(Mid(Text_Line, 67, 8))
                        'DataVersamento = VB.Right(DataVersamento, 2) & Mid(DataVersamento, 5, 2) & VB.Left(DataVersamento, 4)
                        'prelevo il codice tributo
                        CodTributo = Trim(Mid(Text_Line, 79, 4))
                        'prelevo il flag_cf_errato
                        FlagTributoErrato = Trim(Mid(Text_Line, 83, 1))
                        'prelevo l'anno di riferimento del versamento
                        Anno_Riferimento = Trim(Mid(Text_Line, 88, 4))
                        'prelevo il flag_cf_errato
                        FlagAnnoErrato = Trim(Mid(Text_Line, 92, 1))

                        'prelevo la divisa del versamento
                        Divisa = Trim(Mid(Text_Line, 93, 1))
                        'prelevo l'importo totale del versamento
                        If IsNumeric(Trim(Mid(Text_Line, 96, 15))) Then
                            Importo_Versamento = CDbl(Trim(Mid(Text_Line, 96, 15)))
                            ''''Importo_Code_Line = CDbl(Trim(Mid(Text_Line, 96, 15)))
                            'calcolo il totale dei versamenti
                            If UCase(Divisa) = "E" Then
                                TotaleImportiEuro += CDbl(CentInEuro(Importo_Versamento, 2))
                                TotaleImportiEuroFlusso += CDbl(CentInEuro(Importo_Versamento, 2))
                                'Importo_Code_Line = DecimalEuro(CCur(Importo_Code_Line))
                            Else
                                TotaleImportiEuro += CDbl(Importo_Versamento)
                                TotaleImportiEuroFlusso += CDbl(Importo_Versamento)
                                'Importo_Code_Line = CCur(Importo_Code_Line)
                            End If
                            If Importo_Versamento = 0 Then
                                PrintLine(NumFileLOG, Text_Line)
                                'incremento la variabile dei record scartati
                                NScarti = NScarti + 1
                                GoTo RecordSuccessivo
                            End If
                        Else
                            PrintLine(NumFileLOG, Text_Line)
                            'incremento la variabile dei record scartati
                            NScarti = NScarti + 1
                            GoTo RecordSuccessivo
                        End If
                        'prelevo flag ravvedimento operoso
                        Flag_Rav_Operoso = CInt(Trim(Mid(Text_Line, 126, 1)))
                        'prelevo flag immobili variati
                        FlagImmVariati = CInt(Trim(Mid(Text_Line, 127, 1)))

                        FlagAcconto = 0 : FlagSaldo = 0 : FlagAccontoSaldo = 0
                        'prelevo il tipo di pagamento FLAG_ACCONTO_SALDO
                        If Trim(Mid(Text_Line, 128, 1)) = "1" And Trim(Mid(Text_Line, 129, 1)) = "0" Then
                            FlagAcconto = 1
                        ElseIf Trim(Mid(Text_Line, 128, 1)) = "0" And Trim(Mid(Text_Line, 129, 1)) = "1" Then
                            FlagSaldo = 1
                        ElseIf Trim(Mid(Text_Line, 128, 1)) = "1" And Trim(Mid(Text_Line, 129, 1)) = "1" Then
                            FlagAccontoSaldo = 1
                        End If

                        'prelevo il numero di fabbricati
                        If IsNumeric(Trim(Mid(Text_Line, 130, 3))) Then
                            Num_Fab = CInt(Mid(Text_Line, 130, 3))
                        End If

                        'prelevo flag dati ICI errati
                        FlagDatiICIerrati = CInt(Trim(Mid(Text_Line, 133, 1)))

                        'prelevo l'importo della detrazione
                        If IsNumeric(Trim(Mid(Text_Line, 134, 15))) Then
                            Detrazione = CDbl(Trim(Mid(Text_Line, 134, 15)))
                            'If UCase(Divisa) = "E" Then
                            '    Detrazione = CDbl(CentInEuro(Detrazione, 2))
                            'Else
                            '    Detrazione = CDbl(Detrazione)
                            'End If
                        End If
                        'prelevo il nominativo eliminando i doppi spazi fra una parola e l'altra
                        Cognome = Trim(Mid(Text_Line, 149, 55))
                        Nome = Trim(Mid(Text_Line, 204, 20))
                        'prelevo il Tipo Soggetto, se persona fisica o giuridica
                        TipoSoggetto = Trim(Mid(Text_Line, 224, 1))
                        'prelevo la data nascita
                        DataNascita = Trim(Mid(Text_Line, 225, 8))
                        'prelevo la comune nascita
                        ComuneNascita = Trim(Mid(Text_Line, 233, 25))
                        'prelevo la provincia nascita
                        PvNascita = Trim(Mid(Text_Line, 258, 2))

                        'prelevo il codice che dice se è un bollettino ICI o NO
                        Bollettino_ICI = Trim(Mid(Text_Line, 260, 1))
                        '*** 20140109 - import F24 altri tributi ***
                        'If Bollettino_ICI <> "I" Then
                        '    PrintLine(NumFileLOG, Text_Line)
                        '    'incremento la variabile dei record scartati
                        '    NScarti = NScarti + 1
                        '    GoTo RecordSuccessivo
                        'End If
                        IdOperazione = Trim(Mid(Text_Line, 279, 18))
                        '*** ***
                        'se manca la data di versamento o è sbagliata il record viene scartato su un file di scarti
                        If Trim(DataVersamento) <> "" Then
                            If controlla_data(DataVersamento.Substring(6, 2) & "/" & DataVersamento.Substring(4, 2) & "/" & DataVersamento.Substring(0, 4)) = True Then
                                PrintLine(NumFileLOG, Text_Line)
                                'incremento la variabile dei record scartati
                                NScarti = NScarti + 1
                                GoTo RecordSuccessivo
                            Else
                                'se la data di pagamento è posteriore alla data odierna scarto il pagamento sul file di scarti
                                'If DateDiff("d", Format(Now, "MM/dd/yyyy"), Format(CDate(DataVersamento.Substring(4, 2) & "/" & DataVersamento.Substring(6, 2) & "/" & DataVersamento.Substring(0, 4)), "MM/dd/yyyy")) > 0 Then
                                '    'If DateDiff(Microsoft.VisualBasic.DateInterval.Day, Today, CDate(DataVersamento.Substring(6, 2) & "/" & DataVersamento.Substring(4, 2) & "/" & DataVersamento.Substring(0, 4))) > 0 Then
                                '    PrintLine(NumFileLOG, Text_Line)
                                '    'incremento la variabile dei record scartati
                                '    NScarti = NScarti + 1
                                '    GoTo RecordSuccessivo
                                'End If
                            End If
                        End If

                        'se manca la data di versamento o è sbagliata il record viene scartato su un file di scarti
                        If Trim(DataAccredito) <> "" Then
                            If controlla_data(DataAccredito.Substring(6, 2) & "/" & DataAccredito.Substring(4, 2) & "/" & DataAccredito.Substring(0, 4)) = True Then
                                PrintLine(NumFileLOG, Text_Line)
                                'incremento la variabile dei record scartati
                                NScarti = NScarti + 1
                                GoTo RecordSuccessivo
                            Else
                                ''se la data di pagamento è posteriore alla data odierna scarto il pagamento sul file di scarti
                                'If DateDiff("d", Format(Now, "MM/dd/yyyy"), Format(CDate(DataAccredito.Substring(4, 2) & "/" & DataAccredito.Substring(6, 2) & "/" & DataAccredito.Substring(0, 4)), "MM/dd/yyyy")) > 0 Then '
                                '    'If DateDiff(Microsoft.VisualBasic.DateInterval.Day, Today, CDate(DataAccredito.Substring(6, 2) & "/" & DataAccredito.Substring(4, 2) & "/" & DataAccredito.Substring(0, 4))) > 0 Then
                                '    PrintLine(NumFileLOG, Text_Line)
                                '    'incremento la variabile dei record scartati
                                '    NScarti = NScarti + 1
                                '    GoTo RecordSuccessivo
                                'End If
                            End If
                        End If

                        'ReDim Preserve ObjTabF24(k)
                        'ObjTabF24(k).CodTributo = CodTributo
                        'ObjTabF24(k).IdentificativoA1 = IdentificativoA1
                        'ObjTabF24(k).DataCreazioneFlusso = DataCreazioneFlusso
                        'ObjTabF24(k).Provenienza = Provenienza
                        'ObjTabF24(k).CodiceFiscale = CodiceFiscale
                        'ObjTabF24(k).Cognome = Cognome
                        'ObjTabF24(k).Nome = Nome
                        'ObjTabF24(k).TipoSoggetto = TipoSoggetto
                        'ObjTabF24(k).DataNascita = DataNascita
                        'ObjTabF24(k).ComuneNascita = ComuneNascita
                        'ObjTabF24(k).PvNascita = PvNascita
                        'ObjTabF24(k).Ente = Ente
                        'ObjTabF24(k).CapEnte = CapEnte
                        'ObjTabF24(k).DataVersamento = DataVersamento
                        'ObjTabF24(k).DataAccredito = DataAccredito
                        'ObjTabF24(k).FlagAccontoSaldo = FlagAccontoSaldo
                        'ObjTabF24(k).FlagAcconto = FlagAcconto
                        'ObjTabF24(k).FlagSaldo = FlagSaldo
                        'ObjTabF24(k).Anno_Riferimento = Anno_Riferimento
                        'ObjTabF24(k).FlagAnnoErrato = FlagAnnoErrato
                        'ObjTabF24(k).FlagDatiICIerrati = FlagDatiICIerrati
                        'ObjTabF24(k).FlagTributoErrato = FlagTributoErrato
                        'ObjTabF24(k).FlagCFerrato = FlagCFerrato
                        'ObjTabF24(k).FlagImmVariati = FlagImmVariati
                        'ObjTabF24(k).CodBelfiore = CodBelfiore
                        'ObjTabF24(k).Divisa = Divisa
                        'ObjTabF24(k).Num_Fab = Num_Fab
                        'ObjTabF24(k).Bollettino_EXRurale = Bollettino_EXRurale
                        'ObjTabF24(k).Flag_Rav_Operoso = Flag_Rav_Operoso
                        'ObjTabF24(k).Importo_Versamento = CDbl(CentInEuro(Importo_Versamento, 2))
                        'ObjTabF24(k).Detrazione = CDbl(CentInEuro(Detrazione, 2))
                        'ObjTabF24(k).Bollettino_ICI = Bollettino_ICI
                        'ObjTabF24(k).IdOperazione = IdOperazione

                        SQLInsert = "INSERT INTO TAB_ACQUISIZIONE_F24 ( COD_TRIBUTO,"
                        SQLValues = "VALUES ('" & CodTributo & "',"

                        SQLInsert = SQLInsert & "IDENTIFICATIVO,"
                        SQLValues = SQLValues & "'" & IdentificativoA1 & "',"

                        If Trim(DataCreazioneFlusso) <> "" Then
                            SQLInsert = SQLInsert & "DATA_CREAZIONE,"
                            SQLValues = SQLValues & "'" & DataCreazioneFlusso & "',"
                        Else
                            SQLInsert = SQLInsert & "DATA_CREAZIONE,"
                            SQLValues = SQLValues & "Null , "
                        End If

                        If Trim(Provenienza) <> "" Then
                            SQLInsert = SQLInsert & "PROVENIENZA,"
                            SQLValues = SQLValues & "'" & Provenienza & "',"
                        End If
                        If Trim(CodiceFiscale) <> "" Then
                            SQLInsert = SQLInsert & "CF_PIVA,"
                            SQLValues = SQLValues & "'" & CodiceFiscale & "',"
                        End If
                        ''SQLInsert = SQLInsert & "COGNOME_NOME,"
                        ''SQLValues = SQLValues & "'" & Replace(Nominativo, "'", "''") & "',"

                        'If Trim(Cognome) <> "" Then
                        SQLInsert = SQLInsert & "COGNOME,"
                        SQLValues = SQLValues & "'" & Replace(Cognome, "'", "''") & "',"
                        'End If
                        'If Trim(Nome) <> "" Then
                        SQLInsert = SQLInsert & "NOME,"
                        SQLValues = SQLValues & "'" & Replace(Nome, "'", "''") & "',"
                        'End If
                        SQLInsert = SQLInsert & "SESSO,"
                        SQLValues = SQLValues & "'" & Replace(TipoSoggetto, "'", "''") & "',"

                        SQLInsert = SQLInsert & "DATA_NASCITA,"
                        SQLValues = SQLValues & "'" & Replace(DataNascita, "'", "''") & "',"

                        SQLInsert = SQLInsert & "COMUNE_NASCITA,"
                        SQLValues = SQLValues & "'" & Replace(ComuneNascita, "'", "''") & "',"

                        SQLInsert = SQLInsert & "PV_NASCITA,"
                        SQLValues = SQLValues & "'" & Replace(PvNascita, "'", "''") & "',"

                        If Trim(Ente) <> "" Then
                            SQLInsert = SQLInsert & "DESCRIZIONE_ENTE,"
                            SQLValues = SQLValues & "'" & Replace(Ente, "'", "''") & "',"
                        End If
                        If Trim(CapEnte) <> "" Then
                            SQLInsert = SQLInsert & "CAP_ENTE,"
                            SQLValues = SQLValues & "'" & Replace(CapEnte, "'", "''") & "',"
                        End If
                        If Trim(DataVersamento) <> "" Then
                            SQLInsert = SQLInsert & "DATA_VERSAMENTO,"
                            SQLValues = SQLValues & "'" & DataVersamento & "',"
                        Else
                            SQLInsert = SQLInsert & "DATA_VERSAMENTO,"
                            SQLValues = SQLValues & "Null , "
                        End If
                        If Trim(DataAccredito) <> "" Then
                            SQLInsert = SQLInsert & "DATA_ACCREDITO,"
                            SQLValues = SQLValues & "'" & DataAccredito & "',"
                        Else
                            SQLInsert = SQLInsert & "DATA_ACCREDITO,"
                            SQLValues = SQLValues & "Null , "
                        End If

                        SQLInsert = SQLInsert & " FLAG_ACCONTO_SALDO, FLAG_ACCONTO, FLAG_SALDO,"
                        SQLValues = SQLValues & "" & FlagAccontoSaldo & "," & FlagAcconto & "," & FlagSaldo & ","

                        If Trim(Anno_Riferimento) <> "" Then
                            SQLInsert = SQLInsert & "ANNO,"
                            SQLValues = SQLValues & "'" & Anno_Riferimento & "',"
                        End If

                        SQLInsert = SQLInsert & "FLAG_ANNO_ERRATO,"
                        SQLValues = SQLValues & "" & FlagAnnoErrato & ","

                        SQLInsert = SQLInsert & "FLAG_DATI_ERRATI,"
                        SQLValues = SQLValues & "" & FlagDatiICIerrati & ","

                        SQLInsert = SQLInsert & "FLAG_COD_TRIBUTO_ERRATO,"
                        SQLValues = SQLValues & "" & FlagTributoErrato & ","

                        SQLInsert = SQLInsert & "FLAG_CF_ERRATO,"
                        SQLValues = SQLValues & "" & FlagCFerrato & ","

                        SQLInsert = SQLInsert & "FLAG_IMMOBILI_VARIATI,"
                        SQLValues = SQLValues & "" & FlagImmVariati & ","

                        If CodBelfiore <> "" Then
                            SQLInsert = SQLInsert & "COD_BELFIORE,"
                            SQLValues = SQLValues & "'" & CodBelfiore & "',"
                        End If

                        If Divisa <> "" Then
                            SQLInsert = SQLInsert & "DIVISA,"
                            SQLValues = SQLValues & "'" & Divisa & "',"
                        End If

                        If Trim(CStr(Num_Fab)) <> "" Then
                            SQLInsert = SQLInsert & "N_FAB,"
                            SQLValues = SQLValues & "" & Num_Fab & ","
                        End If

                        If Trim(Bollettino_EXRurale) <> "" Then
                            SQLInsert = SQLInsert & "BOLLETTINO_EX_RURALE,"
                            SQLValues = SQLValues & "'" & Bollettino_EXRurale & "',"
                        End If
                        If Trim(Flag_Rav_Operoso) <> "" Then
                            SQLInsert = SQLInsert & "FLAG_RAVVEDIMENTO_OPEROSO,"
                            SQLValues = SQLValues & "'" & Flag_Rav_Operoso & "',"
                        End If
                        If Trim(CStr(Importo_Versamento)) <> "" Then
                            SQLInsert = SQLInsert & "IMPORTO,"
                            SQLValues = SQLValues & Replace(CDbl(CentInEuro(Importo_Versamento, 2)), ",", ".") & ","
                        End If

                        If Trim(CStr(Detrazione)) <> "" Then
                            SQLInsert = SQLInsert & "DETRAZIONE,"
                            SQLValues = SQLValues & Replace(CDbl(CentInEuro(Detrazione, 2)), ",", ".") & ","
                        End If
                        '*** 20140109 - import F24 altri tributi ***
                        'If Trim(CStr(Bollettino_ICI)) <> "" Then
                        SQLInsert = SQLInsert & "TIPO_IMPOSTA"
                        SQLValues = SQLValues & "'" & Bollettino_ICI.Trim & "'"
                        'End If
                        SQLInsert += ",IDOPERAZIONE"
                        SQLValues += ",'" & IdOperazione.Trim & "'"
                        '*** ***
                        SQLInsert = SQLInsert & ",KEYF24)"
                        SQLValues = SQLValues & ",'" & myKey & "')"
                        WriteLog(PathFile, NomeFile, "query::" & SQLInsert & SQLValues)
                        Try
                            CmdDati = New SqlClient.SqlCommand(SQLInsert & SQLValues, objconn)
                            CmdDati.ExecuteNonQuery()
                            contaG1Acq = contaG1Acq + 1
                        Catch ex As Exception
                            PrintLine(NumFileLOG, Text_Line)
                            'incremento la variabile dei record scartati
                            NScarti = NScarti + 1
                        End Try

                        k = k + 1

                    Case "G2"
                        '*** 20120828 - IMU adeguamento per importi statali ***
                        If Text_Line.Substring(38, 4) = "IFEL" Then
                            If IsNumeric(Trim(Mid(Text_Line, 46, 15))) Then
                                ImpAccreditoDispostoIFEL += CDbl(Trim(Mid(Text_Line, 46, 15)))
                            End If
                        Else
                            If IsNumeric(Trim(Mid(Text_Line, 46, 15))) Then
                                ImpAccreditoDispostoEnte += CDbl(Trim(Mid(Text_Line, 46, 15)))
                            End If
                        End If
                        '*** ***
                        'If Trim(Mid(Text_Line, 61, 1)) = "I" Then
                        'End If

                    Case "Z1"
                        'setto la variabile di controllo acquisizione record di coda a true
                        RcFinaleAcquisito = 1
                        '*** 20120828 - IMU adeguamento per importi statali ***
                        'If CStr(Format(CDbl(CentInEuro(ImpAccreditoDispostoEnte, 2)), "#,##0.00")) <> CStr(Format(TotaleImportiEuroFlusso, "#,##0.00")) Then
                        '    sMyErr = "Nel Flusso analizzato i totali acquisiti non tornano con i totali del record G2"
                        'End If
                        sMyErr += "\nQuota Versamenti: " & CStr(Format(TotaleImportiEuroFlusso, "#,##0.00"))
                        sMyErr += "\nQuota Accredito Ente: " & CStr(Format(CDbl(CentInEuro(ImpAccreditoDispostoEnte, 2)), "#,##0.00"))
                        sMyErr += "\nQuota Accredito IFEL: " & CStr(Format(CDbl(CentInEuro(ImpAccreditoDispostoIFEL, 2)), "#,##0.00"))
                        '*** ***
                        If contaG1 <> (contaG1Acq + NScarti) Then
                            sMyErr += "\nAttenzione!\nNON sono stati importati tutti i record di versamenti"
                            Return False
                        End If
                        ImpAccreditoDispostoEnte = 0 : ImpAccreditoDispostoIFEL = 0
                        TotaleImportiEuroFlusso = 0
                End Select
RecordSuccessivo:
            Loop
            'controllo se sono stati scartati dei record
            If NScarti = 0 Then
                PrintLine(NumFileLOG, "NON SONO PRESENTI SCARTI")
                PrintLine(NumFileLOG, New String("-", 50))
            Else
                PrintLine(NumFileLOG, "TOTALE RECORD SCARTATI N° " & NScarti)
            End If
            'FileClose(NumFileLOG)

            'controllo prima di uscire se è passato dal record di coda altrimenti do errore e non salvo nulla
            If RcFinaleAcquisito = 0 Then
                'ErrAcqICI = "Attenzione! Il tracciato che si sta acquisendo ha un record corrotto." & vbCrLf & "La procedura verrà terminata!"
                Return False
            End If
            'i += 1
            '    Loop
            'End If
            Return True
        Catch ex As Exception
            WriteLog(PathFile, NomeFile, "AcquisizioneFileF24massiva::Si è verificato il seguente errore::" & ex.Message)
            sMyErr = "Si è verificato il seguente errore::" & ex.Message
            Return False
        Finally
            'chiusura del file
            FileClose(NumFile)
            FileClose(NumFileLOG)
            objconn.Close()
        End Try
    End Function

    Public Sub SalvaDatiImportazione(ByVal nomeFile As String, ByVal contaG1 As Integer, ByVal contaG1Acq As Integer, ByVal strconn As String, ByVal operatore As String, ByVal codEnte As String)

        Dim objconn As New SqlConnection(strconn)
        Dim CmdDati As SqlClient.SqlCommand
        Dim sql As String
        Dim DrImport As SqlClient.SqlDataReader
        Dim objDatiImportazione As ObjImportazioni
        Dim lingua_date As String

        Try
            lingua_date = ConfigurationSettings.AppSettings("lingua_date")
            objDatiImportazione.Codice = codEnte
            objDatiImportazione.Data = DateTime.Now.Date
            objDatiImportazione.FileName = nomeFile
            objDatiImportazione.Importato = 1
            objDatiImportazione.Operatore = operatore
            objDatiImportazione.Tipologia = "V"
            objDatiImportazione.toTRecord = contaG1
            objDatiImportazione.toTRecordImportati = contaG1Acq
            objDatiImportazione.Tributo = "8852"

            objconn.Open()

            sql = "INSERT INTO tblImportazioni"
            sql += " (Codice, Tributo, Tipologia, Data, Operatore, FileName, toTRecord, toTRecordImportati)"
            sql += " VALUES ('" + objDatiImportazione.Codice + "', '" + objDatiImportazione.Tributo + "',"
            sql += "'" + objDatiImportazione.Tipologia + "', "
            sql += "'" + CDate(objDatiImportazione.Data).ToString(lingua_date).Replace(".", ":") + "', "
            sql += "'" + objDatiImportazione.Operatore + "', '" + objDatiImportazione.FileName + "',"
            sql += "'" + objDatiImportazione.toTRecord.ToString() + "', "
            sql += "'" + objDatiImportazione.toTRecordImportati.ToString() + "',"
            sql += "'1'"
            sql += ")"

            CmdDati = New SqlClient.SqlCommand(sql, objconn)
            CmdDati.ExecuteNonQuery()

            CmdDati.Dispose()
            CmdDati = Nothing
            objconn.Close()
        Catch ex As Exception

        End Try


    End Sub

    Public Sub SelezionaDatiPerVersamenti(ByVal strconn As String)

        Dim sql, SqlTotalizzatori, SqlRcTesta, sqlInsert As String
        Dim objconn As New SqlConnection(strconn)
        Dim CmdDati As SqlClient.SqlCommand
        Dim DrRcTesta, DrVersamenti As SqlClient.SqlDataReader
        Dim TotXprogress, TotRcFlussi, TotErrati, I As Integer
        Dim ObjFlussi() As ObjFlussi

        objconn.Open()

        SqlTotalizzatori = "SELECT count(*) as totF24"
        SqlTotalizzatori += " FROM TAB_ACQUISIZIONE_F24"
        CmdDati = New SqlClient.SqlCommand(SqlTotalizzatori, objconn)

        TotXprogress = 0
        DrRcTesta = CmdDati.ExecuteReader
        If DrRcTesta.HasRows = True Then
            Do While DrRcTesta.Read
                TotXprogress = CInt(DrRcTesta("totf24"))
            Loop
        End If
        DrRcTesta.Close()
        CmdDati.Dispose()
        CmdDati = Nothing

        SqlRcTesta = "SELECT COD_BELFIORE, DATA_CREAZIONE, PROVENIENZA, IDENTIFICATIVO"
        SqlRcTesta += " FROM TAB_ACQUISIZIONE_F24"
        SqlRcTesta += " GROUP BY COD_BELFIORE, DATA_CREAZIONE, PROVENIENZA, IDENTIFICATIVO"
        CmdDati = New SqlClient.SqlCommand(SqlRcTesta, objconn)

        DrRcTesta = CmdDati.ExecuteReader
        If DrRcTesta.HasRows = True Then
            Do While DrRcTesta.Read
                ReDim Preserve ObjFlussi(TotRcFlussi)
                ObjFlussi(TotRcFlussi).sCOD_BELFIORE = DrRcTesta("COD_BELFIORE")
                ObjFlussi(TotRcFlussi).sDATA_CREAZIONE = DrRcTesta("DATA_CREAZIONE")
                ObjFlussi(TotRcFlussi).sPROVENIENZA = DrRcTesta("PROVENIENZA")
                ObjFlussi(TotRcFlussi).sIDENTIFICATIVO = DrRcTesta("IDENTIFICATIVO")
                TotRcFlussi += 1
            Loop
        End If
        DrRcTesta.Close()
        CmdDati.Dispose()
        CmdDati = Nothing

        If Not ObjFlussi Is Nothing Then
            For I = 0 To ObjFlussi.Length - 1
                sql = "SELECT COD_BELFIORE, DATA_CREAZIONE, PROVENIENZA, CF_PIVA, COGNOME, NOME, ANNO, FLAG_RAVVEDIMENTO_OPEROSO, FLAG_ACCONTO, FLAG_SALDO, "
                sql += "FLAG_ACCONTO_SALDO, SUM(CASE WHEN COD_TRIBUTO='3901' OR COD_TRIBUTO='3904' THEN N_FAB ELSE 0 END) AS NFAB, "
                sql += "DATA_ACCREDITO, DATA_VERSAMENTO, FLAG_ANNO_ERRATO = CASE WHEN SUM(FLAG_ANNO_ERRATO) > 0 THEN  1 ELSE 0 END, "
                sql += "FLAG_DATI_ERRATI = CASE WHEN SUM(FLAG_DATI_ERRATI) > 0 THEN  1 ELSE 0 END, "
                sql += "SUM(CASE WHEN COD_TRIBUTO='3901' THEN IMPORTO ELSE 0 END) AS ABIPRIN, "
                sql += "SUM(CASE WHEN COD_TRIBUTO='3902' THEN IMPORTO ELSE 0 END) AS TERAGR, "
                sql += "SUM(CASE WHEN COD_TRIBUTO='3903' THEN IMPORTO ELSE 0 END) AS AREEFAB, "
                sql += "SUM(CASE WHEN COD_TRIBUTO='3904' THEN IMPORTO ELSE 0 END) AS ALTRIFAB, "
                sql += "SUM(CASE WHEN COD_TRIBUTO='3906' THEN IMPORTO ELSE 0 END) AS INTERESSI, "
                sql += "SUM(CASE WHEN COD_TRIBUTO='3907' THEN IMPORTO ELSE 0 END) AS SANZIONI, SUM(IMPORTO) AS TOTVERSATO, SUM(DETRAZIONE) AS DETRAZIONE "
                sql += "FROM(TAB_ACQUISIZIONE_F24) "
                sql += "WHERE COD_BELFIORE='" + ObjFlussi(I).sCOD_BELFIORE + "' AND PROVENIENZA='" + ObjFlussi(I).sPROVENIENZA + "' AND DATA_CREAZIONE='" + ObjFlussi(I).sDATA_CREAZIONE + "' "
                sql += "AND IDENTIFICATIVO='" + ObjFlussi(I).sIDENTIFICATIVO + "' "
                sql += "GROUP BY COD_BELFIORE, DATA_CREAZIONE, PROVENIENZA, CF_PIVA, COGNOME, NOME, ANNO, FLAG_RAVVEDIMENTO_OPEROSO, FLAG_ACCONTO, FLAG_SALDO, "
                sql += "FLAG_ACCONTO_SALDO, DATA_ACCREDITO, DATA_VERSAMENTO "
                sql += "ORDER BY COD_BELFIORE, DATA_CREAZIONE, PROVENIENZA, DATA_ACCREDITO, DATA_VERSAMENTO, CF_PIVA, ANNO"

                CmdDati = New SqlClient.SqlCommand(sql, objconn)

                DrVersamenti = CmdDati.ExecuteReader
                If DrVersamenti.HasRows = True Then
                    'SCORRO DATAREADER
                    Do While DrVersamenti.Read
                        'sqlInsert = "INSERT INTO TblVersamenti (Ente, IdAnagrafico, AnnoRiferimento, "
                        'sqlInsert += "CodiceFiscale, PartitaIva, ImportoPagato, DataPagamento, NumeroBollettino, NumeroFabbricatiPosseduti, Acconto, "
                        'sqlInsert += "Saldo, RavvedimentoOperoso, ImpoTerreni, ImportoAreeFabbric, ImportoAbitazPrincipale, "
                        'sqlInsert += "ImportoAltriFabbric, DetrazioneAbitazPrincipale, ContoCorrente, ComuneUbicazioneImmobile, "
                        'sqlInsert += "ComuneIntestatario, Bonificato, DataInizioValidità, DataFineValidità, Operatore, Annullato, "
                        'sqlInsert += "ImportoSoprattassa, ImportoPenaPecuniaria, Interessi, Violazione, IDProvenienza, NumeroAttoAccertamento, DataProvvedimentoViolazione, ImportoPagatoArrotondamento) "
                        'sqlInsert += "VALUES ("
                        'sqlInsert += "ente, @idAnagrafico, @annoRiferimento, @codiceFiscale, @partitaIva, @importoPagato, "
                        'sqlInsert += "@dataPagamento, @numeroBollettino, @numeroFabbricatiPosseduti, @acconto, @saldo, @ravvedimentoOperoso, "
                        'sqlInsert += "@impoTerreni, @importoAreeFabbric, @importoAbitazPrincipale, @importoAltriFabbric, @detrazioneAbitazPrincipale, "
                        'sqlInsert += "@contoCorrente, @comuneUbicazioneImmobile, @comuneIntestatario, @bonificato, @dataInizioValidità, "
                        'sqlInsert += "@dataFineValidità, @operatore, @annullato, @importoSoprattassa, @importoPenaPecuniaria, @interessi, @violazione, @idProvenienza, @numeroAttoAccertamento, @dataProvvedimentoViolazione, @ImportoPagatoArrotondamento)"
                    Loop
                End If
                DrVersamenti.Close()
            Next
        End If

        objconn.Close()

    End Sub

    Public Sub AggiornaFlagDatiErrati(ByVal strconn As String)

        Dim objconn As New SqlConnection(strconn)
        'Dim objtrans As SqlTransaction
        Dim CmdDati As SqlClient.SqlCommand
        Dim sql As String
        Dim DrErrati As SqlClient.SqlDataReader
        Dim TotErrati As Integer
        Dim ObjG1errati() As ObjVersamentiErrati

        objconn.Open()

        sql = "SELECT CF_PIVA, ANNO, FLAG_ACCONTO, FLAG_SALDO, FLAG_ACCONTO_SALDO, DATA_ACCREDITO, DATA_VERSAMENTO"
        sql += " FROM TAB_ACQUISIZIONE_F24"
        sql += " WHERE FLAG_DATI_ERRATI=1"
        CmdDati = New SqlClient.SqlCommand(sql, objconn)

        TotErrati = 0

        DrErrati = CmdDati.ExecuteReader
        If DrErrati.HasRows = True Then
            Do While DrErrati.Read
                ReDim Preserve ObjG1errati(TotErrati)
                If Not IsDBNull(DrErrati("CF_PIVA")) Then
                    ObjG1errati(TotErrati).sCF_PIVA = CStr(DrErrati("CF_PIVA"))
                Else
                    ObjG1errati(TotErrati).sCF_PIVA = ""
                End If
                If Not IsDBNull(DrErrati("ANNO")) Then
                    ObjG1errati(TotErrati).sANNO = CStr(DrErrati("ANNO"))
                Else
                    ObjG1errati(TotErrati).sANNO = ""
                End If
                If Not IsDBNull(DrErrati("FLAG_ACCONTO")) Then
                    ObjG1errati(TotErrati).sFLAG_ACCONTO = CInt(DrErrati("FLAG_ACCONTO"))
                Else
                    ObjG1errati(TotErrati).sFLAG_ACCONTO = 0
                End If
                If Not IsDBNull(DrErrati("FLAG_SALDO")) Then
                    ObjG1errati(TotErrati).sFLAG_SALDO = CInt(DrErrati("FLAG_SALDO"))
                Else
                    ObjG1errati(TotErrati).sFLAG_SALDO = 0
                End If
                If Not IsDBNull(DrErrati("FLAG_ACCONTO_SALDO")) Then
                    ObjG1errati(TotErrati).sFLAG_ACCONTO_SALDO = CStr(DrErrati("FLAG_ACCONTO_SALDO"))
                Else
                    ObjG1errati(TotErrati).sFLAG_ACCONTO_SALDO = 0
                End If
                If Not IsDBNull(DrErrati("DATA_ACCREDITO")) Then
                    ObjG1errati(TotErrati).sDATA_ACCREDITO = CStr(DrErrati("DATA_ACCREDITO"))
                Else
                    ObjG1errati(TotErrati).sDATA_ACCREDITO = ""
                End If
                If Not IsDBNull(DrErrati("DATA_VERSAMENTO")) Then
                    ObjG1errati(TotErrati).sDATA_VERSAMENTO = CStr(DrErrati("DATA_VERSAMENTO"))
                Else
                    ObjG1errati(TotErrati).sDATA_VERSAMENTO = ""
                End If

                TotErrati += 1

            Loop
        End If
        DrErrati.Close()
        CmdDati.Dispose()
        CmdDati = Nothing

        If Not ObjG1errati Is Nothing Then
            If ObjG1errati.Length > 0 Then
                For TotErrati = 0 To ObjG1errati.Length - 1

                    sql = "UPDATE TAB_ACQUISIZIONE_F24 SET FLAG_DATI_ERRATI=1"
                    sql += " WHERE "

                    sql += " CF_PIVA='" & ObjG1errati(TotErrati).sCF_PIVA & "' AND"

                    'If Not IsDBNull(DrErrati("ANNO")) Then
                    sql += " ANNO='" & ObjG1errati(TotErrati).sANNO & "' AND"
                    'End If
                    'If Not IsDBNull(DrErrati("FLAG_ACCONTO")) Then
                    sql += " FLAG_ACCONTO=" & ObjG1errati(TotErrati).sFLAG_ACCONTO & " AND"
                    'End If
                    'If Not IsDBNull(DrErrati("FLAG_SALDO")) Then
                    sql += " FLAG_SALDO='" & ObjG1errati(TotErrati).sFLAG_SALDO & "' AND"
                    'End If
                    'If Not IsDBNull(DrErrati("FLAG_ACCONTO_SALDO")) Then
                    sql += " FLAG_ACCONTO_SALDO=" & ObjG1errati(TotErrati).sFLAG_ACCONTO_SALDO & " AND"
                    'End If
                    'If Not IsDBNull(DrErrati("DATA_ACCREDITO")) Then
                    sql += " DATA_ACCREDITO='" & ObjG1errati(TotErrati).sDATA_ACCREDITO & "' AND"
                    'End If
                    'If Not IsDBNull(DrErrati("DATA_VERSAMENTO")) Then
                    sql += " DATA_VERSAMENTO='" & ObjG1errati(TotErrati).sDATA_VERSAMENTO & "'"
                    'End If
                    CmdDati = New SqlClient.SqlCommand(sql, objconn)
                    CmdDati.ExecuteNonQuery()
                    'CF_PIVA, ANNO, FLAG_ACCONTO, FLAG_SALDO, FLAG_ACCONTO_SALDO, DATA_ACCREDITO, DATA_VERSAMENTO"

                Next

                CmdDati.Dispose()
                CmdDati = Nothing
                objconn.Close()

            End If

        End If

    End Sub

    Public Sub AggiornaFlagAnnoErrato(ByVal strconn As String)

        Dim objconn As New SqlConnection(strconn)
        'Dim objtrans As SqlTransaction
        Dim CmdDati As SqlClient.SqlCommand
        Dim sql As String
        Dim DrErrati As SqlClient.SqlDataReader
        Dim TotErrati As Integer
        Dim ObjG1errati() As ObjVersamentiErrati

        objconn.Open()

        sql = "SELECT CF_PIVA, FLAG_ACCONTO, FLAG_SALDO, FLAG_ACCONTO_SALDO, DATA_ACCREDITO, DATA_VERSAMENTO"
        sql += " FROM TAB_ACQUISIZIONE_F24"
        sql += " WHERE FLAG_ANNO_ERRATO=1"
        CmdDati = New SqlClient.SqlCommand(sql, objconn)

        TotErrati = 0

        DrErrati = CmdDati.ExecuteReader
        If DrErrati.HasRows = True Then
            Do While DrErrati.Read
                ReDim Preserve ObjG1errati(TotErrati)
                If Not IsDBNull(DrErrati("CF_PIVA")) Then
                    ObjG1errati(TotErrati).sCF_PIVA = CStr(DrErrati("CF_PIVA"))
                Else
                    ObjG1errati(TotErrati).sCF_PIVA = ""
                End If
                If Not IsDBNull(DrErrati("FLAG_ACCONTO")) Then
                    ObjG1errati(TotErrati).sFLAG_ACCONTO = CInt(DrErrati("FLAG_ACCONTO"))
                Else
                    ObjG1errati(TotErrati).sFLAG_ACCONTO = 0
                End If
                If Not IsDBNull(DrErrati("FLAG_SALDO")) Then
                    ObjG1errati(TotErrati).sFLAG_SALDO = CInt(DrErrati("FLAG_SALDO"))
                Else
                    ObjG1errati(TotErrati).sFLAG_SALDO = 0
                End If
                If Not IsDBNull(DrErrati("FLAG_ACCONTO_SALDO")) Then
                    ObjG1errati(TotErrati).sFLAG_ACCONTO_SALDO = CStr(DrErrati("FLAG_ACCONTO_SALDO"))
                Else
                    ObjG1errati(TotErrati).sFLAG_ACCONTO_SALDO = 0
                End If
                If Not IsDBNull(DrErrati("DATA_ACCREDITO")) Then
                    ObjG1errati(TotErrati).sDATA_ACCREDITO = CStr(DrErrati("DATA_ACCREDITO"))
                Else
                    ObjG1errati(TotErrati).sDATA_ACCREDITO = ""
                End If
                If Not IsDBNull(DrErrati("DATA_VERSAMENTO")) Then
                    ObjG1errati(TotErrati).sDATA_VERSAMENTO = CStr(DrErrati("DATA_VERSAMENTO"))
                Else
                    ObjG1errati(TotErrati).sDATA_VERSAMENTO = ""
                End If

                TotErrati += 1

            Loop
        End If
        DrErrati.Close()
        CmdDati.Dispose()
        CmdDati = Nothing

        If Not ObjG1errati Is Nothing Then
            If ObjG1errati.Length > 0 Then
                For TotErrati = 0 To ObjG1errati.Length - 1

                    sql = "UPDATE TAB_ACQUISIZIONE_F24 SET FLAG_ANNO_ERRATO=1"
                    sql += " WHERE "

                    sql += " CF_PIVA='" & ObjG1errati(TotErrati).sCF_PIVA & "' AND"

                    'If Not IsDBNull(DrErrati("ANNO")) Then
                    'sql += " ANNO='" & ObjG1errati(TotErrati).sANNO & "' AND"
                    'End If
                    'If Not IsDBNull(DrErrati("FLAG_ACCONTO")) Then
                    sql += " FLAG_ACCONTO=" & ObjG1errati(TotErrati).sFLAG_ACCONTO & " AND"
                    'End If
                    'If Not IsDBNull(DrErrati("FLAG_SALDO")) Then
                    sql += " FLAG_SALDO='" & ObjG1errati(TotErrati).sFLAG_SALDO & "' AND"
                    'End If
                    'If Not IsDBNull(DrErrati("FLAG_ACCONTO_SALDO")) Then
                    sql += " FLAG_ACCONTO_SALDO=" & ObjG1errati(TotErrati).sFLAG_ACCONTO_SALDO & " AND"
                    'End If
                    'If Not IsDBNull(DrErrati("DATA_ACCREDITO")) Then
                    sql += " DATA_ACCREDITO='" & ObjG1errati(TotErrati).sDATA_ACCREDITO & "' AND"
                    'End If
                    'If Not IsDBNull(DrErrati("DATA_VERSAMENTO")) Then
                    sql += " DATA_VERSAMENTO='" & ObjG1errati(TotErrati).sDATA_VERSAMENTO & "'"
                    'End If
                    CmdDati = New SqlClient.SqlCommand(sql, objconn)
                    CmdDati.ExecuteNonQuery()
                    'CF_PIVA, ANNO, FLAG_ACCONTO, FLAG_SALDO, FLAG_ACCONTO_SALDO, DATA_ACCREDITO, DATA_VERSAMENTO"

                Next

                CmdDati.Dispose()
                CmdDati = Nothing
                objconn.Close()

            End If

        End If
    End Sub

    '*** 20140109 - ***
    ''' <summary>
    ''' se ci sono dati nella tabella TAB_ACQUISIZIONE_F24 prendo dati per popolare oggetto 
    ''' SpostaFile((percorsoF24 + filename), percorsoDestF24, myErr)
    ''' prelevo i dati della tabella TAB_ACQUISIZIONE_F24 raggruppati per tipo di file importato
    ''' per ogni flusso richiamo 4 stored:1.inserimento nuove anagrafiche ici 2.inserimento versamenti ici 3. inserimento pagamenti e dettaglio tarsu 4. spostamento non abbinati
    ''' </summary>
    ''' <param name="CodEnte"></param>
    ''' <param name="Tributo"></param>
    ''' <param name="myConnection"></param>
    ''' <param name="percorsoF24"></param>
    ''' <param name="filename"></param>
    ''' <param name="percorsoDestF24"></param>
    ''' <param name="sBelfiore"></param>
    ''' <param name="sOperatore"></param>
    ''' <param name="IdFlussoPag"></param>
    ''' <param name="myErr"></param>
    ''' <returns></returns>
    ''' <revisionHistory>
    ''' <revision date="12/04/2019">
    ''' Modifiche da revisione manuale
    ''' </revision>
    ''' </revisionHistory>
    Public Function ImportSuTributo(ByVal CodEnte As String, ByVal Tributo As String, ByVal myConnection As String, ByVal percorsoF24 As String, ByVal filename As String, ByVal percorsoDestF24 As String, ByVal sBelfiore As String, ByVal sOperatore As String, ByVal IdFlussoPag As Integer, ByRef myErr As String) As Boolean
        Dim myDataView As New DataView
        Dim arrayFlussi As New ArrayList
        Dim sSQL As String = ""

        Try
            Dim NFileToImport As Integer = NImportazioni(myConnection, sBelfiore)
            If (NFileToImport > 0) Then
                Try
                    Using ctx As New DBModel(DbType, myConnection)
                        sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_F24GetDatiFlusso", "Belfiore")
                        myDataView = ctx.GetDataView(sSQL, "TBL", ctx.GetParam("Belfiore", sBelfiore))
                        For Each myRow As DataRowView In myDataView
                            'dati delle importazioni - filtri per il recupero dei dati
                            Dim oRigaFlusso As ObjFlussi = New ObjFlussi
                            oRigaFlusso.sCOD_BELFIORE = StringOperation.FormatString(myRow("COD_BELFIORE"))
                            oRigaFlusso.sDATA_CREAZIONE = StringOperation.FormatString(myRow("DATA_CREAZIONE"))
                            oRigaFlusso.sIDENTIFICATIVO = StringOperation.FormatString(myRow("IDENTIFICATIVO"))
                            oRigaFlusso.sPROVENIENZA = StringOperation.FormatString(myRow("PROVENIENZA"))
                            arrayFlussi.Add(oRigaFlusso)
                        Next
                        Dim objFlussi() As ObjFlussi = CType(arrayFlussi.ToArray(GetType(ObjFlussi)), ObjFlussi())
                        Try
                            If (Not (objFlussi) Is Nothing) Then
                                Dim k As Integer = 0
                                For k = 0 To objFlussi.Length - 1
                                    Try
                                        sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_F24SetNewAnag", "Belfiore", "Provenienza", "DataCreazione", "Identificativo", "Operatore", "IdEnte")
                                        ctx.ExecuteNonQuery(sSQL, ctx.GetParam("Belfiore", objFlussi(k).sCOD_BELFIORE) _
                                            , ctx.GetParam("Provenienza", objFlussi(k).sPROVENIENZA) _
                                            , ctx.GetParam("DataCreazione", objFlussi(k).sDATA_CREAZIONE) _
                                            , ctx.GetParam("Identificativo", objFlussi(k).sIDENTIFICATIVO) _
                                            , ctx.GetParam("Operatore", sOperatore) _
                                            , ctx.GetParam("IdEnte", CodEnte)
                                        )
                                    Catch ex As Exception
                                        myErr = "Errore in prc_F24SetNewAnag->" + ex.Message
                                        Return False
                                    End Try
                                    Try
                                        sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_F24SetNewICI", "Belfiore", "Provenienza", "DataCreazione", "Identificativo", "Operatore", "IdEnte", "Tributo", "FileName")
                                        ctx.ExecuteNonQuery(sSQL, ctx.GetParam("Belfiore", objFlussi(k).sCOD_BELFIORE) _
                                            , ctx.GetParam("Provenienza", objFlussi(k).sPROVENIENZA) _
                                            , ctx.GetParam("DataCreazione", objFlussi(k).sDATA_CREAZIONE) _
                                            , ctx.GetParam("Identificativo", objFlussi(k).sIDENTIFICATIVO) _
                                            , ctx.GetParam("Operatore", sOperatore) _
                                            , ctx.GetParam("IdEnte", CodEnte) _
                                            , ctx.GetParam("Tributo", Tributo) _
                                            , ctx.GetParam("FileName", filename)
                                        )
                                    Catch ex As Exception
                                        myErr = "Errore in prc_F24SetNewICI->" + ex.Message
                                        Return False
                                    End Try
                                    Try
                                        sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_F24SetNewTARSU", "Belfiore", "Provenienza", "DataCreazione", "Identificativo", "Operatore", "IdEnte", "Tributo", "FileName", "IdFlusso")
                                        ctx.ExecuteNonQuery(sSQL, ctx.GetParam("Belfiore", objFlussi(k).sCOD_BELFIORE) _
                                            , ctx.GetParam("Provenienza", objFlussi(k).sPROVENIENZA) _
                                            , ctx.GetParam("DataCreazione", objFlussi(k).sDATA_CREAZIONE) _
                                            , ctx.GetParam("Identificativo", objFlussi(k).sIDENTIFICATIVO) _
                                            , ctx.GetParam("Operatore", sOperatore) _
                                            , ctx.GetParam("IdEnte", CodEnte) _
                                            , ctx.GetParam("Tributo", "0434") _
                                            , ctx.GetParam("FileName", filename) _
                                            , ctx.GetParam("IdFlusso", IdFlussoPag)
                                        )
                                    Catch ex As Exception
                                        myErr = "Errore in prc_F24SetNewTARSU->" + ex.Message
                                        Return False
                                    End Try
                                    Try
                                        sSQL = ctx.GetSQL(DBModel.TypeQuery.StoredProcedure, "prc_F24SetNewNonAbb", "Belfiore", "Provenienza", "DataCreazione", "Identificativo", "IdFlusso")
                                        ctx.ExecuteNonQuery(sSQL, ctx.GetParam("Belfiore", objFlussi(k).sCOD_BELFIORE) _
                                            , ctx.GetParam("Provenienza", objFlussi(k).sPROVENIENZA) _
                                            , ctx.GetParam("DataCreazione", objFlussi(k).sDATA_CREAZIONE) _
                                            , ctx.GetParam("Identificativo", objFlussi(k).sIDENTIFICATIVO) _
                                            , ctx.GetParam("IdFlusso", IdFlussoPag)
                                        )
                                    Catch ex As Exception
                                        myErr = "Errore in prc_F24SetNewNonAbb->" + ex.Message
                                        Return False
                                    End Try
                                Next
                            End If
                        Catch ex As Exception
                            myErr = "Errore in 'F24SetNew'::" & ex.Message
                            Return False
                        End Try
                        ctx.Dispose()
                    End Using
                Catch ex As Exception
                    myErr = "Errore in 'prelevo i dati della tabella TAB_ACQUISIZIONE_F24 raggruppati per tipo di file importato'::" & ex.Message
                    Return False
                End Try
            End If
            Return True
        Catch ex As Exception
            myErr = "Errore in ImportSuTributo::" & ex.Message
            Return False
        End Try
    End Function
    'Public Function ImportSuTributo(ByVal CodEnte As String, ByVal Tributo As String, ByVal myConnection As String, ByVal percorsoF24 As String, ByVal filename As String, ByVal percorsoDestF24 As String, ByVal sBelfiore As String, ByVal sOperatore As String, ByVal IdFlussoPag As Integer, ByRef myErr As String) As Boolean
    '    Dim sqlConnection As New SqlConnection()
    '    Dim sqlCmd As New SqlCommand
    '    Dim drMyDati As SqlDataReader
    '    Dim arrayFlussi As New ArrayList

    '    Try
    '        Dim NFileToImport As Integer = NImportazioni(myConnection, sBelfiore)
    '        If (NFileToImport > 0) Then
    '            'se ci sono dati nella tabella TAB_ACQUISIZIONE_F24 prendo dati per popolare oggetto versamenti
    '            'SpostaFile((percorsoF24 + filename), percorsoDestF24, myErr)
    '            'prelevo i dati della tabella TAB_ACQUISIZIONE_F24 raggruppati per tipo di file importato
    '            sqlConnection.ConnectionString = myConnection
    '            sqlConnection.Open()
    '            Try
    '                sqlCmd = sqlConnection.CreateCommand()
    '                sqlCmd.CommandTimeout = 0
    '                sqlCmd.CommandType = CommandType.StoredProcedure
    '                sqlCmd.CommandText = "prc_F24GetDatiFlusso"
    '                sqlCmd.Parameters.AddWithValue("@Belfiore", sBelfiore)
    '                drMyDati = sqlCmd.ExecuteReader()
    '                If Not drMyDati Is Nothing Then
    '                    Do While drMyDati.Read
    '                        'dati delle importazioni - filtri per il recupero dei dati
    '                        Dim oRigaFlusso As ObjFlussi = New ObjFlussi
    '                        oRigaFlusso.sCOD_BELFIORE = drMyDati("COD_BELFIORE").ToString
    '                        oRigaFlusso.sDATA_CREAZIONE = drMyDati("DATA_CREAZIONE").ToString
    '                        oRigaFlusso.sIDENTIFICATIVO = drMyDati("IDENTIFICATIVO").ToString
    '                        oRigaFlusso.sPROVENIENZA = drMyDati("PROVENIENZA").ToString
    '                        arrayFlussi.Add(oRigaFlusso)
    '                    Loop
    '                End If
    '            Catch ex As Exception
    '                myErr = "Errore in 'prelevo i dati della tabella TAB_ACQUISIZIONE_F24 raggruppati per tipo di file importato'::" & ex.Message
    '                Return False
    '            Finally
    '                drMyDati.Close()
    '                If Not sqlCmd Is Nothing Then
    '                    sqlCmd.Dispose()
    '                End If
    '            End Try
    '            'per ogni flusso richiamo 4 stored:1.inserimento nuove anagrafiche ici 2.inserimento versamenti ici 3. inserimento pagamenti e dettaglio tarsu 4. spostamento non abbinati
    '            Dim objFlussi() As ObjFlussi = CType(arrayFlussi.ToArray(GetType(ObjFlussi)), ObjFlussi())
    '            Try
    '                If (Not (objFlussi) Is Nothing) Then
    '                    Dim k As Integer = 0
    '                    For k = 0 To objFlussi.Length - 1
    '                        Try
    '                            sqlCmd = New SqlCommand
    '                            sqlCmd = sqlConnection.CreateCommand()
    '                            sqlCmd.CommandTimeout = 0
    '                            sqlCmd.CommandType = CommandType.StoredProcedure
    '                            sqlCmd.CommandText = "prc_F24SetNewAnag"
    '                            sqlCmd.Parameters.AddWithValue("@Belfiore", objFlussi(k).sCOD_BELFIORE)
    '                            sqlCmd.Parameters.AddWithValue("@Provenienza", objFlussi(k).sPROVENIENZA)
    '                            sqlCmd.Parameters.AddWithValue("@DataCreazione", objFlussi(k).sDATA_CREAZIONE)
    '                            sqlCmd.Parameters.AddWithValue("@Identificativo", objFlussi(k).sIDENTIFICATIVO)
    '                            sqlCmd.Parameters.AddWithValue("@Operatore", sOperatore)
    '                            sqlCmd.Parameters.AddWithValue("@IdEnte", CodEnte)
    '                            sqlCmd.ExecuteNonQuery()
    '                        Catch ex As Exception
    '                            myErr = "Errore in prc_F24SetNewAnag"
    '                            Return False
    '                        Finally
    '                            If Not sqlCmd Is Nothing Then
    '                                sqlCmd.Dispose()
    '                            End If
    '                        End Try
    '                        Try
    '                            sqlCmd = New SqlCommand
    '                            sqlCmd = sqlConnection.CreateCommand()
    '                            sqlCmd.CommandTimeout = 0
    '                            sqlCmd.CommandType = CommandType.StoredProcedure
    '                            sqlCmd.CommandText = "prc_F24SetNewICI"
    '                            sqlCmd.Parameters.AddWithValue("@Belfiore", objFlussi(k).sCOD_BELFIORE)
    '                            sqlCmd.Parameters.AddWithValue("@Provenienza", objFlussi(k).sPROVENIENZA)
    '                            sqlCmd.Parameters.AddWithValue("@DataCreazione", objFlussi(k).sDATA_CREAZIONE)
    '                            sqlCmd.Parameters.AddWithValue("@Identificativo", objFlussi(k).sIDENTIFICATIVO)
    '                            sqlCmd.Parameters.AddWithValue("@Operatore", sOperatore)
    '                            sqlCmd.Parameters.AddWithValue("@IdEnte", CodEnte)
    '                            sqlCmd.Parameters.AddWithValue("@Tributo", Tributo)
    '                            sqlCmd.Parameters.AddWithValue("@FileName", filename)
    '                            sqlCmd.ExecuteNonQuery()
    '                        Catch ex As Exception
    '                            myErr = "Errore in prc_F24SetNewICI"
    '                            Return False
    '                        Finally
    '                            If Not sqlCmd Is Nothing Then
    '                                sqlCmd.Dispose()
    '                            End If
    '                        End Try
    '                        Try
    '                            sqlCmd = New SqlCommand
    '                            sqlCmd = sqlConnection.CreateCommand()
    '                            sqlCmd.CommandTimeout = 0
    '                            sqlCmd.CommandType = CommandType.StoredProcedure
    '                            sqlCmd.CommandText = "prc_F24SetNewTARSU"
    '                            sqlCmd.Parameters.AddWithValue("@Belfiore", objFlussi(k).sCOD_BELFIORE)
    '                            sqlCmd.Parameters.AddWithValue("@Provenienza", objFlussi(k).sPROVENIENZA)
    '                            sqlCmd.Parameters.AddWithValue("@DataCreazione", objFlussi(k).sDATA_CREAZIONE)
    '                            sqlCmd.Parameters.AddWithValue("@Identificativo", objFlussi(k).sIDENTIFICATIVO)
    '                            sqlCmd.Parameters.AddWithValue("@Operatore", sOperatore)
    '                            sqlCmd.Parameters.AddWithValue("@IdEnte", CodEnte)
    '                            sqlCmd.Parameters.AddWithValue("@Tributo", "0434")
    '                            sqlCmd.Parameters.AddWithValue("@FileName", filename)
    '                            sqlCmd.Parameters.AddWithValue("@IdFlusso", IdFlussoPag)
    '                            sqlCmd.ExecuteNonQuery()
    '                        Catch ex As Exception
    '                            myErr = "Errore in prc_F24SetNewTARSU"
    '                            Return False
    '                        Finally
    '                            If Not sqlCmd Is Nothing Then
    '                                sqlCmd.Dispose()
    '                            End If
    '                        End Try
    '                        Try
    '                            sqlCmd = New SqlCommand
    '                            sqlCmd = sqlConnection.CreateCommand()
    '                            sqlCmd.CommandTimeout = 0
    '                            sqlCmd.CommandType = CommandType.StoredProcedure
    '                            sqlCmd.CommandText = "prc_F24SetNewNonAbb"
    '                            sqlCmd.Parameters.AddWithValue("@Belfiore", objFlussi(k).sCOD_BELFIORE)
    '                            sqlCmd.Parameters.AddWithValue("@Provenienza", objFlussi(k).sPROVENIENZA)
    '                            sqlCmd.Parameters.AddWithValue("@DataCreazione", objFlussi(k).sDATA_CREAZIONE)
    '                            sqlCmd.Parameters.AddWithValue("@Identificativo", objFlussi(k).sIDENTIFICATIVO)
    '                            sqlCmd.Parameters.AddWithValue("@IdFlusso", IdFlussoPag)
    '                            sqlCmd.ExecuteNonQuery()
    '                        Catch ex As Exception
    '                            myErr = "Errore in prc_F24SetNewNonAbb"
    '                            Return False
    '                        Finally
    '                            If Not sqlCmd Is Nothing Then
    '                                sqlCmd.Dispose()
    '                            End If
    '                        End Try
    '                    Next
    '                End If
    '            Catch ex As Exception
    '                sqlConnection.Dispose()
    '            End Try
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        myErr = "Errore in ImportSuTributo::" & ex.Message
    '        Return False
    '    Finally
    '        If Not sqlConnection Is Nothing Then
    '            sqlConnection.Close()
    '        End If
    '    End Try
    'End Function

    Private Function NImportazioni(ByVal myConnection As String, ByVal Belfiore As String) As Integer
        Dim _sqlConnection As New SqlConnection()
        Dim sqlCmd As New SqlCommand
        Dim sqlRead As SqlDataReader
        Dim nFile As Integer = 0

        Try
            _sqlConnection.ConnectionString = myConnection
            _sqlConnection.Open()
            sqlCmd = _sqlConnection.CreateCommand()
            sqlCmd.CommandType = CommandType.StoredProcedure
            sqlCmd.CommandText = "prc_F24GetNImportazioni"
            sqlCmd.Parameters.AddWithValue("@Belfiore", Belfiore)
            sqlRead = sqlCmd.ExecuteReader()
            If sqlRead.Read() Then
                nFile = CInt(sqlRead("totF24"))
            End If
        Catch ex As Exception
            nFile = -1
        Finally
            If Not sqlRead Is Nothing Then
                sqlRead.Close()
            End If
            If Not sqlCmd Is Nothing Then
                sqlCmd.Dispose()
            End If
            _sqlConnection.Dispose()
        End Try
        Return nFile
    End Function

    Private Function DatiFlusso(ByVal myConnection As String, ByVal Belfiore As String, ByRef myErr As String) As SqlDataReader
        Dim _sqlConnection As New SqlConnection()
        Dim sqlCmd As New SqlCommand
        Dim sqlRead As SqlDataReader
        Try
            _sqlConnection.ConnectionString = myConnection
            _sqlConnection.Open()
            sqlCmd = _sqlConnection.CreateCommand()
            sqlCmd.CommandType = CommandType.StoredProcedure
            sqlCmd.CommandText = "prc_F24GetDatiFlusso"
            sqlCmd.Parameters.AddWithValue("@Belfiore", Belfiore)
            sqlRead = sqlCmd.ExecuteReader()
        Catch ex As Exception
            myErr = "Errore in DatiFlusso::" & ex.Message
            sqlRead = Nothing
        Finally
            If Not sqlCmd Is Nothing Then
                sqlCmd.Dispose()
            End If
            _sqlConnection.Close()
        End Try
        Return sqlRead
    End Function

    Private Function SpostaFile(ByVal fileName As String, ByVal percorsodestinazione As String, ByRef myErr As String) As String
        Dim nomefilespostato As String = ""
        Try
            Dim data As String = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString()
            Dim oraminuti As String = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString()
            Dim infoFile As New IO.FileInfo(fileName)
            Dim nomeFile As String = infoFile.Name
            Dim estensione As String = infoFile.Extension
            Dim nomeFileSenzaEstensione As String = nomeFile.Substring(0, (nomeFile.Length - estensione.Length))
            nomefilespostato = (percorsodestinazione + nomeFileSenzaEstensione + data + "-" + oraminuti + estensione)
            IO.File.Move(fileName, nomefilespostato)
        Catch ex As Exception
            myErr = "Errore in SpostaFile::" & ex.Message
        End Try
        Return nomefilespostato
    End Function
    '*** ***
    Private Structure ObjFlussi
        Dim sCOD_BELFIORE As String
        Dim sDATA_CREAZIONE As String
        Dim sPROVENIENZA As String
        Dim sIDENTIFICATIVO As String
    End Structure

    Private Structure ObjImportazioni
        Dim Codice As String
        Dim Tributo As String
        Dim Tipologia As String
        Dim Data As DateTime
        Dim Operatore As String
        Dim FileName As String
        Dim Importato As Integer
        Dim toTRecord As String
        Dim toTRecordImportati As String
    End Structure

    Private Structure ObjVersamentiErrati
        Dim sCF_PIVA As String
        Dim sANNO As String
        Dim sFLAG_ACCONTO As Integer
        Dim sFLAG_SALDO As Integer
        Dim sFLAG_ACCONTO_SALDO As Integer
        Dim sDATA_ACCREDITO As String
        Dim sDATA_VERSAMENTO As String
    End Structure

    Private Structure ObjTabellaF24
        Dim CodTributo As String
        Dim IdentificativoA1 As String
        Dim DataCreazioneFlusso As String
        Dim Provenienza As String
        Dim CodiceFiscale As String
        Dim Cognome As String
        Dim Nome As String
        Dim TipoSoggetto As String
        Dim DataNascita As String
        Dim ComuneNascita As String
        Dim PvNascita As String
        Dim Ente As String
        Dim CapEnte As String
        Dim DataVersamento As String
        Dim DataAccredito As String
        Dim FlagAccontoSaldo As Integer
        Dim FlagAcconto As Integer
        Dim FlagSaldo As Integer
        Dim Anno_Riferimento As String
        Dim FlagAnnoErrato As Integer
        Dim FlagDatiICIerrati As Integer
        Dim FlagTributoErrato As Integer
        Dim FlagCFerrato As Integer
        Dim FlagImmVariati As Integer
        Dim CodBelfiore As String
        Dim Divisa As String
        Dim Num_Fab As Integer
        Dim Bollettino_EXRurale As String
        Dim Flag_Rav_Operoso As Integer
        Dim Importo_Versamento As Double
        Dim Detrazione As Double
        Dim Bollettino_ICI As String
        '*** 20140109 - import F24 altri tributi ***
        Dim IdOperazione As String
        '*** ***
    End Structure

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
