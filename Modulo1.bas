Attribute VB_Name = "Modulo1"
Option Compare Database

'================ Parametri ===================
Const CARTELLA_DATISAP = "D:\Documenti\DatiSAP\"
Const RIEPILOGO_LAVORI = "\\ve0rs011\DTP10\GestoriTecnici\Riepilogo lavori.xlsx"
Const FILE_MILESTONE = "D:\Documenti\DatiSAP\MLST_Scadute.xlsx"
Const FILE_OPERE = "D:\Documenti\DatiSAP\Elenco Opere.xltx"
Const AVVISO_SCADENZA_CIG = 15 'giorni
Const PM_INGEGNERIA = "PERRONE FRANCESCA"
'==============================================

Const EXCEL_FORMATO_NUMERO = "0.00"
Const EXCEL_FORMATO_DATA = "m/d/yyyy"
Const EXCEL_COLORE_ROSSO = -16776961
Const EXCEL_COLORE_GIALLO_SCURO = -16727809
Const EXCEL_COLORE_VERDE_SCURO = -11489280
Const EXCEL_COLORE_GRIGIO_CHIARO = -2500135
Const EXCEL_COLORE_CELESTE = -1003520

Dim CRITERIO_RICERCA_DISPONIBILITA As String

Dim session As SAPFEWSELib.GuiSession
Dim ws As Worksheet
Dim wb As Workbook
Dim wnd As Window
Dim rg As Range
Dim gw As GuiGridView
Dim xlsapp As EXCEL.Application
Dim tabIncarichiDL As DAO.Recordset
Dim tabIncarichiAS As DAO.Recordset
Dim tabIncarichiCSE As DAO.Recordset
Dim tabVC As DAO.Recordset
Dim tabLavori As DAO.Recordset
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rsCAT2 As DAO.Recordset
Dim rsODL As DAO.Recordset
Dim rsIncarichi As DAO.Recordset
Dim rsDisponibilità As DAO.Recordset
Dim rsUO As DAO.Recordset
Dim rsPresenze As DAO.Recordset
Dim DataInEsame As String
Dim Lun, Mar, Mer, Gio, Ven As String
Dim LunTeo, MarTeo, MerTeo, GioTeo, VenTeo As Double
Dim LunAss, MarAss, MerAss, GioAss, VenAss As Double
Dim LunStr, MarStr, MerStr, GioStr, VenStr As Double

Sub test()
'    connetti
'    Dim tbl As GuiTableControl
'    Set tbl = session.FindById("wnd[0]/usr/subPERNR_LIST:SAPLCATS:1200/tblSAPLCATSTC_PERNR")
'    SAP_TableControl_PaginaGiù tbl
    MsgBox TroncaNumero(3.1264, 2)
End Sub

Sub EXCEL_VisualizzaMLSTScadute()
    Set xlsapp = New EXCEL.Application
    Set wb = xlsapp.Workbooks.Open(FILE_MILESTONE)
    Set ws = wb.Sheets("Milestone")
    ws.Activate
    Set rg = ws.Range("2:" & ws.UsedRange.Rows.Count + 1)
    rg.Delete Shift:=xlUp
    Set db = CurrentDb
    Set rg = ws.Cells
    colprogetto = EXCEL_ColonnaDaNome("Progetto")
    colIntervento = EXCEL_ColonnaDaNome("Intervento")
    colODA = EXCEL_ColonnaDaNome("ODA")
    colContratto = EXCEL_ColonnaDaNome("Contratto")
    colSpec = EXCEL_ColonnaDaNome("Spec.")
    colNetwork = EXCEL_ColonnaDaNome("Network")
    colDescrizione_Lavori = EXCEL_ColonnaDaNome("Descrizione_Lavori")
    colTesto_Milestone = EXCEL_ColonnaDaNome("Testo_Milestone")
    colDL = EXCEL_ColonnaDaNome("DL")
    colAS = EXCEL_ColonnaDaNome("AS")
    colData_Prevista = EXCEL_ColonnaDaNome("Data_Prevista")
    colRitardo = EXCEL_ColonnaDaNome("Ritardo")
    colData_Effettiva = EXCEL_ColonnaDaNome("Data_Effettiva")
    
    Set rs = db.OpenRecordset("qry_MLST_Scadute", dbOpenDynaset)
    riga = 2
    While Not rs.EOF
        DoEvents
        Form_PannelloDiControllo.lbAvanzamento.Caption = "Avanzamento: " & rs.PercentPosition
        ws.Cells(riga, colprogetto).Value = rs!Progetto
        ws.Cells(riga, colIntervento).Value = rs!Intervento
        ws.Cells(riga, colODA).Value = rs!ODA
        ws.Cells(riga, colContratto).Value = rs!Contratto
        ws.Cells(riga, colSpec).Value = rs!Specialità
        ws.Cells(riga, colNetwork).Value = rs!Network
        ws.Cells(riga, colDescrizione_Lavori).Value = rs![Descrizione lavori]
        ws.Cells(riga, colTesto_Milestone).Value = rs![testo milestone]
        ws.Cells(riga, colDL).Value = rs!DL
        ws.Cells(riga, colAS).Value = rs!AS
        ws.Cells(riga, colData_Prevista).Value = rs![data fissa cardine]
        ws.Cells(riga, colRitardo).Value = rs!ritardo
        ws.Cells(riga, colData_Effettiva).Value = rs![data eff]
        riga = riga + 1
        rs.MoveNext
    Wend
    rs.Close
    xlsapp.Visible = True
End Sub

Function ScriviValore(ByVal Valore As Object) As String
    If IsNull(Valore) Then ScriviValore = "" Else ScriviValore = Valore
End Function

Sub EXCEL_ApplicaFiltroSchedaNetwork(ByVal NomeColonnaExcel As String, ByVal NomeCampoAccess As String, ByVal IstruzioneSQLdiFiltro As String)
    Set xlsapp = New EXCEL.Application
    Set wb = xlsapp.Workbooks.Open(RIEPILOGO_LAVORI)
    Set ws = wb.Sheets("Network")
    ws.Activate
    EXCEL_CancellaFiltro
    Set db = CurrentDb
    Set rg = ws.Cells
    Dim campoFiltro As Integer
    
    campoFiltro = EXCEL_ColonnaDaNome(NomeColonnaExcel)
    Set rs = db.OpenRecordset(IstruzioneSQLdiFiltro, dbOpenDynaset)
    Dim filtro() As String
    Dim i As Integer
    i = -1
    While Not rs.EOF
        i = i + 1
        ReDim Preserve filtro(i)
        filtro(i) = rs.Fields(NomeCampoAccess).Value
        rs.MoveNext
    Wend
    rs.Close
    rg.AutoFilter Field:=campoFiltro, Criteria1:=filtro, Operator:=xlFilterValues
    xlsapp.Visible = True
End Sub

Sub connetti()
    Set SapGuiAuto = GetObject("SAPGUI")
    Set App = SapGuiAuto.GetScriptingEngine
    Set Connection = App.Children(0)
    Set session = Connection.Children(0)
End Sub

Sub AggiungiNTWDisponibiliPerCDL()
    Set db = CurrentDb
    Set rsCAT2 = db.OpenRecordset("CAT2", dbOpenDynaset)
    Dim filtroCDL As String
    filtroCDL = "([Centro di lavoro]='"
    For i = 0 To Form_CAT2.elencoCDL.ListCount - 1
        filtroCDL = filtroCDL & Form_CAT2.elencoCDL.ItemData(i) & "' OR [Centro di lavoro]='"
    Next
    filtroCDL = Left(filtroCDL, Len(filtroCDL) - 24)
    filtroCDL = filtroCDL & ")"
    Set rsNTW = db.OpenRecordset("SELECT * FROM qry_DisponibilitàCap15_perNTW_e_CDL WHERE " & filtroCDL & " AND (DisponibileWBS>0 AND PianificatoResiduo>0 AND COBL='' AND RIL='RIL.')", dbOpenDynaset)
    rsNTW.MoveFirst
    While Not rsNTW.EOF
        DoEvents
        If NTWNonEsistenteInCAT2(rsNTW!Network, rsNTW!Operazione, rsNTW!Elemento) Then
            AggiungiNTWsuCAT2 rsNTW!Network, rsNTW!Operazione, rsNTW!Elemento
        End If
        rsNTW.MoveNext
    Wend
    rsCAT2.Close
    rsNTW.Close
End Sub

Function NTWNonEsistenteInCAT2(ByVal Network As String, Operazione As String, Elemento As String) As Boolean
    Dim rsCAT2 As DAO.Recordset
    Set rsCAT2 = db.OpenRecordset("CAT2", dbOpenDynaset)
    rsCAT2.FindFirst "CID='" & Form_CAT2.cbDipendente.Value & "' AND Network='" & Network & "' AND Operazione='" & Operazione & "' AND Sottoperazione='" & Elemento & "'"
    NTWNonEsistenteInCAT2 = rsCAT2.NoMatch
    rsCAT2.Close
End Function

Sub AggiungiNTWsuCAT2(ByVal Network As String, Operazione As String, Elemento As String)
    Dim rsCAT2 As DAO.Recordset
    Set rsCAT2 = db.OpenRecordset("CAT2", dbOpenDynaset)
    rsCAT2.AddNew
    rsCAT2!CID = Form_CAT2.cbDipendente.Value
    rsCAT2!Network = Network
    rsCAT2!Operazione = Operazione
    rsCAT2!Sottoperazione = Elemento
    rsCAT2.Update
    rsCAT2.Close
End Sub

Function SAP_TableControl_ColonnaDaTitolo(ByVal Titolo As String, Tabella As GuiTableControl) As Long
    Dim Risultato As Long
    Risultato = -1
    For i = 0 To Tabella.Columns.Count - 1
        If Tabella.Columns(i).Title = Titolo Then
            Risultato = i
            Exit For
        End If
    Next
    SAP_TableControl_ColonnaDaTitolo = Risultato
End Function

Sub SAP_CompilaCAT2()
    connetti
    session.StartTransaction "CAT2"
    Set db = CurrentDb
    DoCmd.OpenQuery "qry_ELIMINA_PresenzeDaCAT2"
    filtroUO = ""
    CID = ""
    If Form_CAT2.cbCAT2CID.Value = True Then
        CID = Form_CAT2.cbDipendente.Value
        filtroUO = "SELECT DISTINCT UO FROM qry_Dettaglio_CAT2_per_inserimento_su_SAP WHERE CID='" & CID & "' Order By UO"
    Else
        If Form_CAT2.cbCAT2UOCorrente.Value = True Then
            filtroUO = "SELECT DISTINCT UO FROM qry_Dettaglio_CAT2_per_inserimento_su_SAP WHERE UO='" & Form_CAT2.cbUO.Value & "' Order By UO"
        Else
            filtroUO = "SELECT DISTINCT UO FROM qry_Dettaglio_CAT2_per_inserimento_su_SAP WHERE UO<>'' Order By UO"
        End If
    End If
    Set rsUO = db.OpenRecordset(filtroUO, dbOpenDynaset)
    rsUO.MoveFirst
    While Not rsUO.EOF
        DoEvents
        SAP_CAT2_PremiSelezionePersonale
        SAP_CAT2_InserisciUO rsUO!UO
        'SAP_CAT2_InserisciData Form_CAT2.tbData.Text
        SAP_Esegui
        SAP_CAT2_SelezionaRigheDipendenti rsUO!UO, CID
        SAP_CAT2_Matita
        If Form_CAT2.cbSettimana.Value = "Scorsa" Then
            SAP_CAT2_PremiFrecciaIndietroSettimana
        End If
        SAP_CAT2_LeggiOrarioTeorico
        SAP_CAT2_InserisciCAT2perUO rsUO!UO, CID
        MsgBox "Verificare i dati e salvare"
        rsUO.MoveNext
    Wend
    DoCmd.OpenQuery "qry_ACCODA_CAT2Storico"
    rsUO.Close
    
End Sub

Sub SAP_CAT2_SelezionaRigheDipendenti(ByVal UO As String, Optional ByVal CID As String = "")
    Dim rsDip As DAO.Recordset
    If CID = "" Then
        Set rsDip = db.OpenRecordset("SELECT * FROM qry_Dettaglio_DipendenzePerUO WHERE UO='" & UO & "'", dbOpenDynaset)
    Else
        Set rsDip = db.OpenRecordset("SELECT * FROM qry_Dettaglio_DipendenzePerUO WHERE Dipendente='" & CID & "'", dbOpenDynaset)
    End If
    Dim tbl As GuiTableControl
    Set tbl = session.FindById("wnd[0]/usr/subPERNR_LIST:SAPLCATS:1200/tblSAPLCATSTC_PERNR")
    Dim colCID As Long
    colCID = SAP_TableControl_ColonnaDaTitolo("C.I.D.", tbl)
    Dim riga As Long
    riga = 0
    While tbl.Rows(riga).Count > 0
        rsDip.FindFirst "Dipendente='" & tbl.GetCell(riga, colCID).Text & "'"
        If rsDip.NoMatch = False Then tbl.Rows(riga).Selected = True
        riga = SAP_TableControl_RigaSuccessiva(riga, tbl)
    Wend
    rsDip.Close
End Sub

Function TroncaNumero(ByVal Numero As Double, NumDecimali As Integer) As Double
    Dim Str As String
    Str = Numero
    Dim pos As Integer
    pos = InStr(1, Str, ",")
    Str = Left(Str, pos + NumDecimali)
    TroncaNumero = CDbl(Str)
End Function

Function AdattaValoreOrarioSuCAT2(ByVal Teorico As Variant, Assenze As Variant, Straordinari As Variant, Divisore As Integer) As Double
    Dim Valore, Teo, Ass, Str As Double
    If IsNull(Teorico) Then Teo = 0 Else Teo = Teorico
    If IsNull(Assenze) Then Ass = 0 Else Ass = Assenze
    If IsNull(Straordinari) Then Str = 0 Else Str = Straordinari
    Valore = TroncaNumero((Teo + Str - Ass) / Divisore, 2)
    If Valore > 0 Then AdattaValoreOrarioSuCAT2 = Valore Else AdattaValoreOrarioSuCAT2 = 0
End Function

Sub SAP_CAT2_InserisciCAT2perUO(ByVal UO As String, Optional ByVal CID As String = "")
    If CID = "" Then
        Set rsCAT2 = db.OpenRecordset("SELECT * FROM qry_Dettaglio_CAT2_per_inserimento_su_SAP WHERE UO='" & UO & "'", dbOpenDynaset)
    Else
        Set rsCAT2 = db.OpenRecordset("SELECT * FROM qry_Dettaglio_CAT2_per_inserimento_su_SAP WHERE CID='" & CID & "'", dbOpenDynaset)
    End If
    rsCAT2.MoveFirst
    Dim tbl As GuiTableControl
    Set tbl = session.FindById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD")
    colCID = SAP_TableControl_ColonnaDaTitolo("C.I.D.", tbl)
    colODL = SAP_TableControl_ColonnaDaTitolo("Ordine dest.", tbl)
    colNTW = SAP_TableControl_ColonnaDaTitolo("Network", tbl)
    colOp = SAP_TableControl_ColonnaDaTitolo("Operazione", tbl)
    colSOp = SAP_TableControl_ColonnaDaTitolo("SOp.", tbl)
    colCDL = SAP_TableControl_ColonnaDaTitolo("C. lav.", tbl)
    colLun = SAP_TableControl_ColonnaDaTitolo(Lun, tbl)
    colmar = SAP_TableControl_ColonnaDaTitolo(Mar, tbl)
    colmer = SAP_TableControl_ColonnaDaTitolo(Mer, tbl)
    colgio = SAP_TableControl_ColonnaDaTitolo(Gio, tbl)
    colven = SAP_TableControl_ColonnaDaTitolo(Ven, tbl)
    strsettimana = session.FindById("wnd[0]/usr/subCATS003:SAPLCATS:2300/txtCATSFIELDS-CATSWEEKEX").Text
    strperiododal = session.FindById("wnd[0]/usr/subCATS003:SAPLCATS:2300/ctxtCATSFIELDS-DATEFROM").Text
    strperiodoal = session.FindById("wnd[0]/usr/subCATS003:SAPLCATS:2300/ctxtCATSFIELDS-DATETO").Text
    settimana = CInt(Left(strsettimana, 2))
    anno = CInt(Right(strsettimana, 4))
    strperiododal = Replace(strperiododal, ".", "/")
    strperiodoal = Replace(strperiodoal, ".", "/")
    Dim rsDataCAT2 As DAO.Recordset
    Set rsDataCAT2 = db.OpenRecordset("CAT2", dbOpenDynaset)
    
    Dim riga As Long
    riga = SAP_CAT2_TableControl_TrovaRigaLibera(tbl)
    While Not rsCAT2.EOF
        DoEvents
        rsDataCAT2.FindFirst "ID=" & rsCAT2!ID
        rsDataCAT2.Edit
        rsDataCAT2!settimana = settimana
        rsDataCAT2!anno = anno
        rsDataCAT2!periododal = CDate(strperiododal)
        rsDataCAT2!periodoal = CDate(strperiodoal)
        rsDataCAT2.Update
        
        tbl.GetCell(riga, colCID).Text = rsCAT2!CID
        If Not IsNull(rsCAT2!ODL) Then tbl.GetCell(riga, colODL).Text = rsCAT2!ODL
        If Not IsNull(rsCAT2!Network) Then tbl.GetCell(riga, colNTW).Text = rsCAT2!Network
        If Not IsNull(rsCAT2!Operazione) Then tbl.GetCell(riga, colOp).Text = rsCAT2!Operazione
        If Not IsNull(rsCAT2!Sottoperazione) Then
            If rsCAT2!Sottoperazione <> "x" Then
                tbl.GetCell(riga, colSOp).Text = rsCAT2!Sottoperazione
            End If
        End If
        If Not IsNull(rsCAT2!CDL) Then tbl.GetCell(riga, colCDL).Text = rsCAT2!CDL
        
        If rsCAT2!Lunedì = True Then
            tbl.GetCell(riga, colLun).Text = AdattaValoreOrarioSuCAT2(rsCAT2!LunTeo, rsCAT2!LunAss, rsCAT2!LunStr, rsCAT2!NumLun)
        End If
        If rsCAT2!Martedì = True Then
            tbl.GetCell(riga, colmar).Text = AdattaValoreOrarioSuCAT2(rsCAT2!MarTeo, rsCAT2!MarAss, rsCAT2!MarStr, rsCAT2!NumMar)
        End If
        If rsCAT2!Mercoledì = True Then
            tbl.GetCell(riga, colmer).Text = AdattaValoreOrarioSuCAT2(rsCAT2!MerTeo, rsCAT2!MerAss, rsCAT2!MerStr, rsCAT2!NumMer)
        End If
        If rsCAT2!Giovedì = True Then
            tbl.GetCell(riga, colgio).Text = AdattaValoreOrarioSuCAT2(rsCAT2!GioTeo, rsCAT2!GioAss, rsCAT2!GioStr, rsCAT2!NumGio)
        End If
        If rsCAT2!Venerdì = True Then
            tbl.GetCell(riga, colven).Text = AdattaValoreOrarioSuCAT2(rsCAT2!VenTeo, rsCAT2!VenAss, rsCAT2!VenStr, rsCAT2!NumVen)
        End If
        rsCAT2.MoveNext
        riga = SAP_CAT2_RigaSuccessivaPerScrittura(riga, tbl)
        Set tbl = session.FindById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD")
    Wend
    
    SAP_CAT2_SelezionaTutto
    SAP_CAT2_InserisciOreTeoriche 'Per compensare eventuali arrotondamenti
    rsCAT2.Close
    rsDataCAT2.Close
End Sub

Sub SAP_CAT2_InserisciOreTeoriche()
    session.FindById("wnd[0]/tbar[1]/btn[36]").Press
End Sub

Function SAP_CAT2_RigaSuccessivaPerScrittura(ByVal riga As Long, ByRef Tabella As GuiTableControl) As Long
    If SAP_TableControl_èUltimaRiga(riga, Tabella) = True Then
        SAP_TableControl_PaginaGiù Tabella
        riga = 1
    Else
        riga = riga + 1
    End If
    SAP_CAT2_RigaSuccessivaPerScrittura = riga
End Function

Function LunedìSettimana(ByVal DataDaValutare As Date) As Date
    giorno = DatePart("d", DataDaValutare, vbMonday, vbFirstFourDays)
    LunedìSettimana = DateAdd("d", 1 - giorno, DataDaValutare)
End Function

Function SAP_CAT2_RigaSuccessiva(ByVal riga As Long, ByRef Tabella As GuiTableControl) As Long
    If SAP_TableControl_èUltimaRiga(riga, Tabella) = True Then
        SAP_TableControl_PaginaGiù Tabella
        riga = 0
    Else
        riga = riga + 1
    End If
    SAP_CAT2_RigaSuccessiva = riga
End Function

Function SAP_TableControl_RigaSuccessiva(ByVal riga As Long, ByRef Tabella As GuiTableControl) As Long
    If SAP_TableControl_èUltimaRiga(riga, Tabella) = True Then
        SAP_TableControl_PaginaGiù Tabella
        riga = 0
    Else
        riga = riga + 1
    End If
    SAP_TableControl_RigaSuccessiva = riga
End Function

Sub SAP_TableControl_PaginaGiù(ByRef Tabella As GuiTableControl)
    Dim scroll As GuiScrollbar
    Set scroll = Tabella.VerticalScrollbar
    Dim IDTabella As String
    IDTabella = Tabella.ID
    scroll.Position = scroll.Position + scroll.PageSize
    Set Tabella = session.FindById(IDTabella)
End Sub

Function SAP_CAT2_TableControl_TrovaRigaLibera(ByRef Tabella As GuiTableControl) As Long
    colCID = SAP_TableControl_ColonnaDaTitolo("C.I.D.", Tabella)
    Dim riga As Long
    riga = 0
    While Tabella.GetCell(riga, colCID).Text <> ""
        DoEvents
        riga = SAP_CAT2_RigaSuccessiva(riga, Tabella)
    Wend
    SAP_CAT2_TableControl_TrovaRigaLibera = riga
End Function

Sub SAP_CAT2_DefinisciNomiColonneDate()
    DataInEsame = session.FindById("wnd[0]/usr/subCATS003:SAPLCATS:2300/ctxtCATSFIELDS-DATEFROM").Text
    Dim datatemp As Date
    Dim dataStr As String
    
    Lun = Left(DataInEsame, 5)
    
    datatemp = CDate(Replace(DataInEsame, ".", "/"))
    datatemp = DateAdd("d", 1, datatemp)
    dataStr = CStr(datatemp)
    dataStr = Replace(dataStr, "/", ".")
    Mar = Left(dataStr, 5)
    
    datatemp = CDate(Replace(DataInEsame, ".", "/"))
    datatemp = DateAdd("d", 2, datatemp)
    dataStr = CStr(datatemp)
    dataStr = Replace(dataStr, "/", ".")
    Mer = Left(dataStr, 5)
    
    datatemp = CDate(Replace(DataInEsame, ".", "/"))
    datatemp = DateAdd("d", 3, datatemp)
    dataStr = CStr(datatemp)
    dataStr = Replace(dataStr, "/", ".")
    Gio = Left(dataStr, 5)
    
    datatemp = CDate(Replace(DataInEsame, ".", "/"))
    datatemp = DateAdd("d", 4, datatemp)
    dataStr = CStr(datatemp)
    dataStr = Replace(dataStr, "/", ".")
    Ven = Left(dataStr, 5)
End Sub

Sub SAP_CAT2_LeggiOrarioTeorico()
    Dim tbl As GuiTableControl
    Set tbl = session.FindById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD")
    Set rsPresenze = db.OpenRecordset("PresenzeDaCAT2", dbOpenDynaset)
    
    SAP_CAT2_DefinisciNomiColonneDate
    
    colCID = SAP_TableControl_ColonnaDaTitolo("C.I.D.", tbl)
    colTR = SAP_TableControl_ColonnaDaTitolo("TR", tbl)
    colLun = SAP_TableControl_ColonnaDaTitolo(Lun, tbl)
    colTPAP = SAP_TableControl_ColonnaDaTitolo("Tp. A/P", tbl)
    riga = 0
    Do
        CID = tbl.GetCell(riga, colCID).Text
        If tbl.GetCell(riga, colTR).IconName = "T_TIME" Then
            LunTeo = ConvertiDouble(tbl.GetCell(riga, colLun).Text)
            MarTeo = ConvertiDouble(tbl.GetCell(riga, colLun + 1).Text)
            MerTeo = ConvertiDouble(tbl.GetCell(riga, colLun + 2).Text)
            GioTeo = ConvertiDouble(tbl.GetCell(riga, colLun + 3).Text)
            VenTeo = ConvertiDouble(tbl.GetCell(riga, colLun + 4).Text)
        Else
            If tbl.Rows(riga).Count > 10 Then 'Le righe color azzurrino hanno solo 10 celle, le altre 28
                If tbl.GetCell(riga, colTPAP).Text = "ABSE" Then
                    LunAss = ConvertiDouble(tbl.GetCell(riga, colLun).Text)
                    MarAss = ConvertiDouble(tbl.GetCell(riga, colLun + 1).Text)
                    MerAss = ConvertiDouble(tbl.GetCell(riga, colLun + 2).Text)
                    GioAss = ConvertiDouble(tbl.GetCell(riga, colLun + 3).Text)
                    VenAss = ConvertiDouble(tbl.GetCell(riga, colLun + 4).Text)
                ElseIf tbl.GetCell(riga, colTPAP).Text = "STRA" Then
                    LunStr = ConvertiDouble(tbl.GetCell(riga, colLun).Text)
                    MarStr = ConvertiDouble(tbl.GetCell(riga, colLun + 1).Text)
                    MerStr = ConvertiDouble(tbl.GetCell(riga, colLun + 2).Text)
                    GioStr = ConvertiDouble(tbl.GetCell(riga, colLun + 3).Text)
                    VenStr = ConvertiDouble(tbl.GetCell(riga, colLun + 4).Text)
                    rsPresenze.AddNew
                    rsPresenze!CID = CID
                    rsPresenze!Data = Lun
                    rsPresenze!giorno = "Lun"
                    rsPresenze!Teorico = LunTeo
                    rsPresenze!assenza = LunAss
                    rsPresenze!straordinario = LunStr
                    rsPresenze.Update
                    rsPresenze.AddNew
                    rsPresenze!CID = CID
                    rsPresenze!Data = Mar
                    rsPresenze!giorno = "Mar"
                    rsPresenze!Teorico = MarTeo
                    rsPresenze!assenza = MarAss
                    rsPresenze!straordinario = MarStr
                    rsPresenze.Update
                    rsPresenze.AddNew
                    rsPresenze!CID = CID
                    rsPresenze!Data = Mer
                    rsPresenze!giorno = "Mer"
                    rsPresenze!Teorico = MerTeo
                    rsPresenze!assenza = MerAss
                    rsPresenze!straordinario = MerStr
                    rsPresenze.Update
                    rsPresenze.AddNew
                    rsPresenze!CID = CID
                    rsPresenze!Data = Gio
                    rsPresenze!giorno = "Gio"
                    rsPresenze!Teorico = GioTeo
                    rsPresenze!assenza = GioAss
                    rsPresenze!straordinario = GioStr
                    rsPresenze.Update
                    rsPresenze.AddNew
                    rsPresenze!CID = CID
                    rsPresenze!Data = Ven
                    rsPresenze!giorno = "Ven"
                    rsPresenze!Teorico = VenTeo
                    rsPresenze!assenza = VenAss
                    rsPresenze!straordinario = VenStr
                    rsPresenze.Update
                End If
            End If
        End If
        riga = SAP_CAT2_RigaSuccessiva(riga, tbl)
        Set tbl = session.FindById("wnd[0]/usr/subCATS002:SAPLCATS:2200/tblSAPLCATSTC_CATSD")
    Loop Until CID = ""
    rsPresenze.Close
End Sub

Function ConvertiDouble(ByVal Valore As String) As Double
    If Valore = "" Then
        ConvertiDouble = 0
    Else
        ConvertiDouble = CDbl(Valore)
    End If
End Function

Function SAP_TableControl_èUltimaRiga(ByVal riga As Long, ByVal Tabella As GuiTableControl) As Boolean
    Dim scroll As GuiScrollbar
    Set scroll = Tabella.VerticalScrollbar
    If riga = (scroll.PageSize - 1) Then
        SAP_TableControl_èUltimaRiga = True
    Else
        SAP_TableControl_èUltimaRiga = False
    End If
End Function

Sub SAP_Salva()
    session.FindById("wnd[0]/tbar[0]/btn[11]").Press
End Sub

Sub SAP_CAT2_PremiFrecciaIndietroSettimana()
    session.FindById("wnd[0]/usr/subCATS003:SAPLCATS:2300/btnPREVIOUS_OBJECT").Press
End Sub

Sub SAP_CAT2_Matita()
    session.FindById("wnd[0]/tbar[1]/btn[5]").Press
End Sub

Sub SAP_CAT2_SelezionaTutto()
    session.FindById("wnd[0]/tbar[1]/btn[19]").Press
End Sub

Sub SAP_CAT2_PremiSelezionePersonale()
    session.FindById("wnd[0]/usr/btnREPSELECTION_BUTTON").Press
End Sub

Sub SAP_CAT2_InserisciUO(ByVal UO As String)
    session.FindById("wnd[0]/usr/ctxtSO_OBJID-LOW").Text = UO
End Sub

Sub SAP_CAT2_InserisciData(ByVal Data As String)
    session.FindById("wnd[0]/usr/ctxtP_BEGDA").Text = Data
End Sub

Sub AggiungiODLStandard()
    Set db = CurrentDb
    Set rsCAT2 = db.OpenRecordset("CAT2", dbOpenDynaset)
    ApplicaFiltroSuODLperCDL
    While Not rsODL.EOF
        DoEvents
        If ODLNonEsistenteInCAT2(rsODL!ordine, rsODL!Operazione, rsODL!SottoOperazione) Then
            AggiungiODLsuCAT2
        End If
        rsODL.MoveNext
    Wend
    rsCAT2.Close
    rsODL.Close
End Sub

Sub AggiungiODLsuCAT2()
    rsCAT2.AddNew
    rsCAT2!CID = Form_CAT2.cbDipendente.Value
    rsCAT2!ODL = rsODL!ordine
    rsCAT2!Operazione = rsODL!Operazione
    rsCAT2!Sottoperazione = rsODL!SottoOperazione
    rsCAT2.Update
End Sub

Sub ApplicaFiltroSuODLperCDL()
    Dim strODL As String
    strODL = "SELECT Ordine, Operazione, Sottooperazione FROM OperazioniODL WHERE [Centro di lavoro]='"
    Set rs = db.OpenRecordset(Form_CAT2.elencoCDL.RowSource, dbOpenDynaset)
    rs.MoveFirst
    While Not rs.EOF
        DoEvents
        strODL = strODL & rs!CDL & "' OR [Centro di lavoro]='"
        rs.MoveNext
    Wend
    strODL = Left(strODL, Len(strODL) - 23)
    Set rsODL = db.OpenRecordset(strODL, dbOpenDynaset)
    rs.Close
End Sub

Sub AggiornaCAT2()
    Set db = CurrentDb
    Set rsCAT2 = db.OpenRecordset("CAT2", dbOpenDynaset)
    Set rsIncarichi = db.OpenRecordset("qry_Elenco_Incarichi_AS_DL_CSE_CAT2", dbOpenDynaset)
    rsIncarichi.MoveFirst
    
    While Not rsIncarichi.EOF
        DoEvents
        Form_CAT2.lbAvanzamento.Caption = rsIncarichi.PercentPosition
        AggiungiNetworkSuCAT2
        rsIncarichi.MoveNext
    Wend
    
    rsCAT2.Close
    rsIncarichi.Close
    MsgBox "Scheda aggiornata"
End Sub

Sub EliminaNetworkNonPiùDisponibili()
    Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM qry_Dettaglio_CAT2 WHERE CID='" & Form_CAT2.cbDipendente.Value & "'", dbOpenDynaset)
    rs.MoveFirst
    While Not rs.EOF
        DoEvents
        If rs!ResiduoPianificato < 0 Or rs!DisponibileWBS < 0 Or rs!COBL = "COBL" Or rs!RIL = "" Then
            DoCmd.RunSQL "DELETE * FROM CAT2 WHERE Network='" & rs!Network & "' AND Operazione='" & rs!Operazione & "' AND Sottoperazione='" & rs!Sottoperazione & "' AND CID='" & Form_CAT2.cbDipendente.Value & "'"
        End If
        rs.MoveNext
    Wend
    rs.Close
End Sub

Sub AggiungiNetworkSuCAT2()
    Set rsDisponibilità = db.OpenRecordset("qry_DisponibilitàCap15_perNTW_e_CDL", dbOpenDynaset)
    CRITERIO_RICERCA_DISPONIBILITA = "Network='" & rsIncarichi!Network & "' AND [Centro di lavoro]='" & rsIncarichi!CDL & "'"
    rsDisponibilità.FindFirst CRITERIO_RICERCA_DISPONIBILITA
    While Not rsDisponibilità.NoMatch
        DoEvents
        rsCAT2.FindFirst "CID='" & rsIncarichi!CID & "' AND Network='" & rsIncarichi!Network & "' AND Operazione='" & rsDisponibilità!Operazione & "' AND Sottoperazione='" & rsDisponibilità!Elemento & "'"
        If rsCAT2.NoMatch Then
            rsCAT2.AddNew
            rsCAT2!CID = rsIncarichi!CID
            rsCAT2!Network = rsIncarichi!Network
            rsCAT2!Operazione = rsDisponibilità!Operazione
            rsCAT2!Sottoperazione = rsDisponibilità!Elemento
            rsCAT2!Incarico = rsIncarichi!primoditipoincarico
            rsCAT2.Update
        Else
            rsCAT2.Edit
            rsCAT2!Incarico = rsIncarichi!primoditipoincarico
            rsCAT2.Update
        End If
        rsDisponibilità.FindNext CRITERIO_RICERCA_DISPONIBILITA
    Wend
    rsDisponibilità.Close
End Sub

Function ODLNonEsistenteInCAT2(ByVal ODL As String, ByVal Operazione As String, ByVal SottoOperazione) As Boolean
    rsCAT2.FindFirst "CID='" & Form_CAT2.cbDipendente.Value & "' AND ODL='" & ODL & "' AND Operazione='" & Operazione & "' AND Sottoperazione='" & SottoOperazione & "'"
    ODLNonEsistenteInCAT2 = rsCAT2.NoMatch
End Function

Function NetworkConDisponibilità(ByVal CriterioRicerca As String) As Boolean
    rsDisponibilità.FindFirst CriterioRicerca
    NetworkConDisponibilità = Not rsDisponibilità.NoMatch
End Function

Sub AggiornaDataEsportazione(ByVal NomeTabellaAggiornata As String)
    Set db = CurrentDb
    Set tabLavori = db.OpenRecordset("DateAggiornamento", dbOpenDynaset)
    tabLavori.FindFirst ("NomeTabella='" & NomeTabellaAggiornata & "'")
    If tabLavori.NoMatch = True Then
        tabLavori.AddNew
    Else
        tabLavori.Edit
    End If
    tabLavori("NomeTabella") = NomeTabellaAggiornata
    tabLavori("DataAggiornamento") = Now
    tabLavori.Update
End Sub

Sub btEsportaLavori()
    Set xlsapp = New EXCEL.Application
    xlsapp.Visible = False
    xlsapp.Workbooks.Open FileName:=RIEPILOGO_LAVORI
    Set ws = xlsapp.Sheets("Network")
    ws.Activate

    Dim af As AutoFilter
    Set af = ws.AutoFilter
    If Not af Is Nothing Then
        If af.FilterMode Then ws.ShowAllData
    End If

    Set rg = ws.Range("3:" & ws.UsedRange.Rows.Count + 2)
    rg.Delete Shift:=xlUp

    Dim db As DAO.Database
    Set db = CurrentDb
    Dim tabDateAggiornamento As DAO.Recordset
    Set tabLavori = db.OpenRecordset("qry_Dettaglio_Lavori", dbOpenDynaset)
    Set tabIncarichiDL = db.OpenRecordset("qry_Elenco_IncarichiDL_Attuali", dbOpenDynaset)
    Set tabIncarichiAS = db.OpenRecordset("qry_Elenco_IncarichiAS_Attuali", dbOpenDynaset)
    Set tabIncarichiCSE = db.OpenRecordset("qry_Elenco_IncarichiCSE_Attuali", dbOpenDynaset)
    Set tabVC = db.OpenRecordset("VCdaScaricare", dbOpenDynaset)
    Dim riga As Integer
    Dim Colonna As Integer
    riga = 3
        
    Set rg = ws.Cells
    With rg.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With

    While Not tabLavori.EOF
        Form_Maschera1.lbAvanzamento.Caption = "Record n." & (riga - 2)
        DoEvents
        
        If UCase(tabLavori("LavoriTerminati")) = "X" Then
            Set rg = ws.Rows(riga)
            rg.Select
            xlsapp.Selection.Font.Color = EXCEL_COLORE_VERDE_SCURO
        End If
        If UCase(tabLavori("Chiusa")) = "X" Then
            Set rg = ws.Rows(riga)
            rg.Select
            xlsapp.Selection.Font.Color = EXCEL_COLORE_GRIGIO_CHIARO
        End If
        
        Colonna = EXCEL_ColonnaDaNome("SpecNTW")
        ws.Cells(riga, Colonna).Value = tabLavori("Specialità")
        
        Colonna = EXCEL_ColonnaDaNome("LINTW")
        ws.Cells(riga, Colonna).Value = tabLavori("LI")
        
        Colonna = EXCEL_ColonnaDaNome("ProgettoNTW")
        If IsNull(tabLavori("Progetto")) Then
            ws.Cells(riga, Colonna).Value = tabLavori("ProgDaNTW")
        Else
            ws.Cells(riga, Colonna).Value = tabLavori("Progetto")
        End If
        
        Colonna = EXCEL_ColonnaDaNome("InterventoNTW")
        ws.Cells(riga, Colonna).Value = tabLavori("Intervento")
        
        Colonna = EXCEL_ColonnaDaNome("Network")
        ws.Cells(riga, Colonna).Value = tabLavori("Network")
        If InStr(1, tabLavori("StatoNetwork"), "ANTW") = 0 Then
            Set rg = ws.Cells(riga, Colonna)
            rg.Select
            xlsapp.Selection.Font.Color = EXCEL_COLORE_ROSSO
        End If
        If InStr(1, tabLavori("StatoNetwork"), "VNTW") > 0 Then
            Set rg = ws.Cells(riga, Colonna)
            rg.Select
            xlsapp.Selection.Font.Color = EXCEL_COLORE_GIALLO_SCURO
        End If
        
        Colonna = EXCEL_ColonnaDaNome("StatoNetwork")
        ws.Cells(riga, Colonna).Value = tabLavori("StatoNetwork")
        
        Colonna = EXCEL_ColonnaDaNome("TitoloNetwork")
        ws.Cells(riga, Colonna).Value = tabLavori("TitoloNetwork")
        
        Colonna = EXCEL_ColonnaDaNome("ODA")
        ws.Cells(riga, Colonna).Value = tabLavori("ODA")
        If tabLavori("Rilascio") = 0 Then
            Set rg = ws.Cells(riga, Colonna)
            rg.Select
            xlsapp.Selection.Font.Color = EXCEL_COLORE_ROSSO
        End If
        If tabLavori("Rilascio") = 1 Then
            Set rg = ws.Cells(riga, Colonna)
            rg.Select
            xlsapp.Selection.Font.Color = EXCEL_COLORE_GIALLO_SCURO
        End If
        
        Colonna = EXCEL_ColonnaDaNome("PMNTW")
        ws.Cells(riga, Colonna).Value = tabLavori("ProjectManager")
        If tabLavori("ProjectManager") <> PM_INGEGNERIA Then
            Set rg = ws.Cells(riga, Colonna)
            rg.Select
            xlsapp.Selection.Font.Color = EXCEL_COLORE_CELESTE
        End If

        Colonna = EXCEL_ColonnaDaNome("Contratto")
        ws.Cells(riga, Colonna).Value = tabLavori("Contratto")
        
        Colonna = EXCEL_ColonnaDaNome("DataContratto")
        ws.Cells(riga, Colonna).Value = tabLavori("Data contratto")
        
        Colonna = EXCEL_ColonnaDaNome("Impresa")
        ws.Cells(riga, Colonna).Value = tabLavori("Appaltatore")
        
        Colonna = EXCEL_ColonnaDaNome("CIG")
        ws.Cells(riga, Colonna).Value = tabLavori("CIG")
        If Not IsNull(tabLavori("ScadenzaCIG")) And tabLavori("ScadenzaCIG") < AVVISO_SCADENZA_CIG And tabLavori("CIG perfezionato") = "" Then
            Set rg = ws.Cells(riga, Colonna)
            rg.Select
            xlsapp.Selection.Font.Color = EXCEL_COLORE_GIALLO_SCURO
        End If
        
        Colonna = EXCEL_ColonnaDaNome("CIGPerfezionato")
        ws.Cells(riga, Colonna).Value = tabLavori("CIG perfezionato")
        
        Colonna = EXCEL_ColonnaDaNome("DataCIG")
        ws.Cells(riga, Colonna).Value = tabLavori("Data CIG")
        
        Colonna = EXCEL_ColonnaDaNome("ggScadenzaCIG")
        ws.Cells(riga, Colonna).Value = tabLavori("ScadenzaCIG")
        
        Colonna = EXCEL_ColonnaDaNome("LavoriTerminati")
        ws.Cells(riga, Colonna).Value = tabLavori("LavoriTerminati")
        
        Colonna = EXCEL_ColonnaDaNome("NetworkChiusa")
        ws.Cells(riga, Colonna).Value = tabLavori("Chiusa")
        
        Colonna = EXCEL_ColonnaDaNome("DescrizioneLavori")
        ws.Cells(riga, Colonna).Value = tabLavori("Descrizione lavori")
        
        Colonna = EXCEL_ColonnaDaNome("PianificatoNTW")
        ws.Cells(riga, Colonna).Value = tabLavori("TotPianificato")
        
        Colonna = EXCEL_ColonnaDaNome("EffettuatoNTW")
        ws.Cells(riga, Colonna).Value = tabLavori("TotEffettuato")
        
        Colonna = EXCEL_ColonnaDaNome("ImpFinODANTW")
        ws.Cells(riga, Colonna).Value = tabLavori("TotImpFinODA")
        
        Colonna = EXCEL_ColonnaDaNome("PercentualeNTW")
        ws.Cells(riga, Colonna).Value = tabLavori("PercentualeAvanzamento")
        
        Colonna = EXCEL_ColonnaDaNome("DL")
        ws.Cells(riga, Colonna).Value = EXCEL_Network_ElencoIncaricati(tabIncarichiDL)
        
        Colonna = EXCEL_ColonnaDaNome("AS")
        ws.Cells(riga, Colonna).Value = EXCEL_Network_ElencoIncaricati(tabIncarichiAS)
        
        riga = riga + 1
        tabLavori.MoveNext
    Wend
    EXCEL_Network_ImpostaFormatoFoglio
    Set tabDateAggiornamento = db.OpenRecordset("DateAggiornamento", dbOpenDynaset)
    tabDateAggiornamento.FindFirst "NomeTabella='GF_NTW'"
    If tabDateAggiornamento.NoMatch = True Then
        ws.Cells(1, 1).Value = ""
    Else
        ws.Cells(1, 1).Value = "Dati aggiornati al: " & Format(tabDateAggiornamento("DataAggiornamento"), "dd/mm/yyyy")
    End If
    ws.Cells(1, 1).Select
    
    '===================================================== Lettere incarico =========================================================
    
    Set ws = xlsapp.Sheets("Lettere incarico")
    ws.Activate
    
    Set af = ws.AutoFilter
    If Not af Is Nothing Then
        If af.FilterMode Then ws.ShowAllData
    End If
    
    Set rg = ws.Range("3:" & ws.UsedRange.Rows.Count + 2)
    rg.Delete Shift:=xlUp

    Set tabLavori = db.OpenRecordset("qry_Dettaglio_LI", dbOpenDynaset)
    riga = 3
    While Not tabLavori.EOF
        Form_Maschera1.lbAvanzamento.Caption = "Record n." & (riga - 2)
        DoEvents
        
        Colonna = EXCEL_ColonnaDaNome("LI")
        ws.Cells(riga, Colonna).Value = tabLavori("LI")

        Colonna = EXCEL_ColonnaDaNome("DataLI")
        ws.Cells(riga, Colonna).Value = tabLavori("Data LI")

        Colonna = EXCEL_ColonnaDaNome("SpecLI")
        ws.Cells(riga, Colonna).Value = tabLavori("Specialità")

        Colonna = EXCEL_ColonnaDaNome("ProgLI")
        ws.Cells(riga, Colonna).Value = tabLavori("Progetto")

        Colonna = EXCEL_ColonnaDaNome("Referenza")
        ws.Cells(riga, Colonna).Value = tabLavori("Referenza")

        Colonna = EXCEL_ColonnaDaNome("ST")
        ws.Cells(riga, Colonna).Value = tabLavori("Soggetto Tecnico")

        Colonna = EXCEL_ColonnaDaNome("RL")
        ws.Cells(riga, Colonna).Value = tabLavori("RL")

        Colonna = EXCEL_ColonnaDaNome("PMLI")
        ws.Cells(riga, Colonna).Value = tabLavori("PM")

        Colonna = EXCEL_ColonnaDaNome("LITerminata")
        ws.Cells(riga, Colonna).Value = tabLavori("LITerminata")

        Colonna = EXCEL_ColonnaDaNome("NTWTotali")
        ws.Cells(riga, Colonna).Value = tabLavori("NTWTot")

        Colonna = EXCEL_ColonnaDaNome("NTWConcluse")
        ws.Cells(riga, Colonna).Value = tabLavori("NTWConcluse")

        Colonna = EXCEL_ColonnaDaNome("NTWChiuse")
        ws.Cells(riga, Colonna).Value = tabLavori("NTWChiuse")

        Colonna = EXCEL_ColonnaDaNome("OggettoLI")
        ws.Cells(riga, Colonna).Value = tabLavori("Oggetto LI")

        Colonna = EXCEL_ColonnaDaNome("Assegnato")
        ws.Cells(riga, Colonna).Value = tabLavori("Assegnato")

        Colonna = EXCEL_ColonnaDaNome("PianificatoLI")
        ws.Cells(riga, Colonna).Value = tabLavori("Pianificato")

        Colonna = EXCEL_ColonnaDaNome("EffettuatoLI")
        ws.Cells(riga, Colonna).Value = tabLavori("Effettuato")

        Colonna = EXCEL_ColonnaDaNome("PercPianificati")
        ws.Cells(riga, Colonna).Value = tabLavori("PercentualePianificato")

        Colonna = EXCEL_ColonnaDaNome("PercEffettuati")
        ws.Cells(riga, Colonna).Value = tabLavori("PercentualeEffettuato")

        riga = riga + 1
        tabLavori.MoveNext
    Wend
    EXCEL_LI_ImpostaFormatoFoglio
    Set tabDateAggiornamento = db.OpenRecordset("DateAggiornamento", dbOpenDynaset)
    tabDateAggiornamento.FindFirst "NomeTabella='GF_NTW'"
    If tabDateAggiornamento.NoMatch = True Then
        ws.Cells(1, 1).Value = ""
    Else
        ws.Cells(1, 1).Value = "Dati aggiornati al: " & Format(tabDateAggiornamento("DataAggiornamento"), "dd/mm/yyyy")
    End If
    ws.Cells(1, 1).Select
    
'===================================================== Opere =========================================================
   
    EXCEL_Dettaglio_Opere
    
    xlsapp.Visible = True
End Sub

Sub EXCEL_Dettaglio_Opere()
'    Set xlsapp = New Excel.Application
'    Set wb = xlsapp.Workbooks.Open(FILE_OPERE)
    Set ws = wb.Sheets("Opere")
    ws.Activate
    
    Dim af As AutoFilter
    Set af = ws.AutoFilter
    If Not af Is Nothing Then
        If af.FilterMode Then ws.ShowAllData
    End If
    
    Set rg = ws.Range("3:" & ws.UsedRange.Rows.Count + 2)
    rg.Delete Shift:=xlUp
    Set db = CurrentDb
    colSpec = EXCEL_ColonnaDaNome("O_Spec")
    colprogetto = EXCEL_ColonnaDaNome("O_Prog")
    colOpera = EXCEL_ColonnaDaNome("O_Opera")
    colPian = EXCEL_ColonnaDaNome("O_Pianificato")
    colBudget = EXCEL_ColonnaDaNome("O_Budget")
    colRichFond = EXCEL_ColonnaDaNome("O_RichiestaFondi")
    colEff = EXCEL_ColonnaDaNome("O_Effettuato")
    colImpFin = EXCEL_ColonnaDaNome("O_ImpFinODA")
    colDisposto = EXCEL_ColonnaDaNome("O_Disposto")
    colDisponibile = EXCEL_ColonnaDaNome("O_Disponibile")
    colLavoro = EXCEL_ColonnaDaNome("O_Lavoro")
    colNetwork = EXCEL_ColonnaDaNome("O_Network")
    colODA = EXCEL_ColonnaDaNome("O_ODA")
    colContratto = EXCEL_ColonnaDaNome("O_Contratto")
    colRefProg = EXCEL_ColonnaDaNome("O_RefProg")
    colDL = EXCEL_ColonnaDaNome("O_DL")
    colAS = EXCEL_ColonnaDaNome("O_AS")
    
    colPian13 = EXCEL_ColonnaDaNome("O_Pian13")
    colEff13 = EXCEL_ColonnaDaNome("O_Eff13")
    colDiff13 = EXCEL_ColonnaDaNome("O_Diff13")
    colImpFin13 = EXCEL_ColonnaDaNome("O_ImpFin13")
    colNonImp13 = EXCEL_ColonnaDaNome("O_NonImp13")
    
    colPian14 = EXCEL_ColonnaDaNome("O_Pian14")
    colEff14 = EXCEL_ColonnaDaNome("O_Eff14")
    colDiff14 = EXCEL_ColonnaDaNome("O_Diff14")
    
    colPian15 = EXCEL_ColonnaDaNome("O_Pian15")
    colEff15 = EXCEL_ColonnaDaNome("O_Eff15")
    colDiff15 = EXCEL_ColonnaDaNome("O_Diff15")
    
    colPian16 = EXCEL_ColonnaDaNome("O_Pian16")
    colEff16 = EXCEL_ColonnaDaNome("O_Eff16")
    colDiff16 = EXCEL_ColonnaDaNome("O_Diff16")
    colStock16 = EXCEL_ColonnaDaNome("O_Stock16")
    
    colPian19 = EXCEL_ColonnaDaNome("O_Pian19")
    colEff19 = EXCEL_ColonnaDaNome("O_Eff19")
    colDiff19 = EXCEL_ColonnaDaNome("O_Diff19")
    
    colPian20 = EXCEL_ColonnaDaNome("O_Pian20")
    colEff20 = EXCEL_ColonnaDaNome("O_Eff20")
    colDiff20 = EXCEL_ColonnaDaNome("O_Diff20")
    
    colPian21 = EXCEL_ColonnaDaNome("O_Pian21")
    colEff21 = EXCEL_ColonnaDaNome("O_Eff21")
    colDiff21 = EXCEL_ColonnaDaNome("O_Diff21")
    
    colPian22 = EXCEL_ColonnaDaNome("O_Pian22")
    colEff22 = EXCEL_ColonnaDaNome("O_Eff22")
    colDiff22 = EXCEL_ColonnaDaNome("O_Diff22")
    
    colPian23 = EXCEL_ColonnaDaNome("O_Pian23")
    colEff23 = EXCEL_ColonnaDaNome("O_Eff23")
    colDiff23 = EXCEL_ColonnaDaNome("O_Diff23")
    
    colPian25 = EXCEL_ColonnaDaNome("O_Pian25")
    colEff25 = EXCEL_ColonnaDaNome("O_Eff25")
    colDiff25 = EXCEL_ColonnaDaNome("O_Diff25")
    
    
    Set rsElenco = db.OpenRecordset("qryElencoOpere", dbOpenDynaset)
    riga = 3
    While Not rsElenco.EOF
        strLavoro = ""
        strNetwork = ""
        strODA = ""
        strContratto = ""
        strDL = ""
        strAS = ""
        strSpec = ""
        rigainizio = riga

        Set rs = db.OpenRecordset("SELECT * FROM qry_GF_Lavori_Opere WHERE Opera='" & rsElenco!Opera & "'", dbOpenDynaset)
        If rs.RecordCount > 0 Then
        
            rs.MoveLast
            rs.MoveFirst
            Set rsGF_WBS = db.OpenRecordset("SELECT * FROM qry_GF_WBS_TuttiCapitoli WHERE Opera='" & rsElenco!Opera & "'", dbOpenDynaset)
            If rsGF_WBS.RecordCount > 0 Then rsGF_WBS.MoveFirst
            DoEvents
            Form_Maschera1.lbAvanzamento.Caption = rsElenco.PercentPosition
            If rs.RecordCount > 1 Then
                While Not rs.EOF
                    strLavoro = strLavoro & ScriviValore(rs!Lavoro) & "," & vbCrLf
                    strNetwork = strNetwork & ScriviValore(rs!Network) & "," & vbCrLf
                    strODA = strODA & ScriviValore(rs!ODA) & "," & vbCrLf
                    strContratto = strContratto & ScriviValore(rs!Contratto) & "," & vbCrLf
                    If InStr(1, strDL, rs!DL) = 0 Then strDL = strDL & ScriviValore(rs!DL) & "," & vbCrLf
                    If InStr(1, strAS, rs!AS) = 0 Then strAS = strAS & ScriviValore(rs!AS) & "," & vbCrLf
                    If InStr(1, strSpec, rs!Specialità) = 0 Then strSpec = strSpec & ScriviValore(rs!Specialità) & "," & vbCrLf
    '                riga = riga + 1
    '                ws.Cells(riga, colSpec).Value = ScriviValore(rs!Specialità)
    '                ws.Cells(riga, colprogetto).Value = ScriviValore(rs!Progetto)
    '                ws.Cells(riga, colOpera).Value = ScriviValore(rsElenco!Opera)
    '                ws.Cells(riga, colLavoro).Value = ScriviValore(rs!Lavoro)
    '                ws.Cells(riga, colNetwork).Value = ScriviValore(rs!Network)
    '                ws.Cells(riga, colODA).Value = ScriviValore(rs!ODA)
    '                ws.Cells(riga, colContratto).Value = ScriviValore(rs!Contratto)
    '                ws.Cells(riga, colDL).Value = ScriviValore(rs!DL)
    '                ws.Cells(riga, colAS).Value = ScriviValore(rs!AS)
    '                ws.Cells(riga, colPian).Value = Val(ScriviValore(rs!PianOp))
    '                ws.Cells(riga, colEff).Value = Val(ScriviValore(rs!EffOp))
    '                ws.Cells(riga, colImpFin).Value = Val(ScriviValore(rs!ImpFinOp))
    '                ws.Cells(riga, colDisposto).Value = Val(ScriviValore(rs!DispOp))
                    rs.MoveNext
                Wend
                If strLavoro <> "" Then strLavoro = Left(strLavoro, Len(strLavoro) - 3)
                If strNetwork <> "" Then strNetwork = Left(strNetwork, Len(strNetwork) - 3)
                If strODA <> "" Then strODA = Left(strODA, Len(strODA) - 3)
                If strContratto <> "" Then strContratto = Left(strContratto, Len(strContratto) - 3)
                If strDL <> "" Then strDL = Left(strDL, Len(strDL) - 3)
                If strAS <> "" Then strAS = Left(strAS, Len(strAS) - 3)
                If strSpec <> "" Then strSpec = Left(strSpec, Len(strSpec) - 3)
                ws.Cells(rigainizio, colSpec).Value = strSpec
                ws.Cells(rigainizio, colLavoro).Value = strLavoro
                ws.Cells(rigainizio, colNetwork).Value = strNetwork
                ws.Cells(rigainizio, colODA).Value = strODA
                ws.Cells(rigainizio, colContratto).Value = strContratto
                ws.Cells(rigainizio, colDL).Value = strDL
                ws.Cells(rigainizio, colAS).Value = strAS
    '            Set rg = ws.Range(ws.Rows(rigainizio + 1), ws.Rows(riga))
    '            rg.Interior.Color = 13434828
    '            rg.Group
            Else
                ws.Cells(riga, colSpec).Value = ScriviValore(rs!Specialità)
                ws.Cells(riga, colRefProg).Value = ScriviValore(rs!RefProg)
                ws.Cells(riga, colOpera).Value = ScriviValore(rsElenco!Opera)
                ws.Cells(riga, colLavoro).Value = ScriviValore(rs!Lavoro)
                ws.Cells(riga, colNetwork).Value = ScriviValore(rs!Network)
                ws.Cells(riga, colODA).Value = ScriviValore(rs!ODA)
                ws.Cells(riga, colContratto).Value = ScriviValore(rs!Contratto)
                ws.Cells(riga, colDL).Value = ScriviValore(rs!DL)
                ws.Cells(riga, colAS).Value = ScriviValore(rs!AS)
            End If
                   
            rs.MoveFirst
            ws.Cells(rigainizio, colprogetto).Value = ScriviValore(rs!Progetto)
            ws.Cells(rigainizio, colprogetto).Value = ScriviValore(rs!Progetto)
            ws.Cells(rigainizio, colOpera).Value = ScriviValore(rsElenco!Opera)
            ws.Cells(rigainizio, colPian).Value = Val(ScriviValore(rs!Pianificato))
            ws.Cells(rigainizio, colBudget).Value = Val(ScriviValore(rs!Budget))
            ws.Cells(rigainizio, colRichFond).Value = Val(ScriviValore(rs!RichiestaFondi))
            ws.Cells(rigainizio, colEff).Value = Val(ScriviValore(rs!Effettuato))
            ws.Cells(rigainizio, colImpFin).Value = Val(ScriviValore(rs!ImpFinODA))
            ws.Cells(rigainizio, colDisposto).Value = Val(ScriviValore(rs!Disposto))
            ws.Cells(rigainizio, colDisponibile).Value = Val(ScriviValore(rs!Disponibile))
            
            ws.Cells(rigainizio, colPian13).Value = Val(ScriviValore(rsGF_WBS!Pian13))
            ws.Cells(rigainizio, colEff13).Value = Val(ScriviValore(rsGF_WBS!Eff13))
            ws.Cells(rigainizio, colDiff13).Value = Val(ScriviValore(rsGF_WBS!Diff13))
            ws.Cells(rigainizio, colImpFin13).Value = Val(ScriviValore(rsGF_WBS!ImpFin13))
            ws.Cells(rigainizio, colNonImp13).Value = Val(ScriviValore(rsGF_WBS!NonImp13))
            
            ws.Cells(rigainizio, colPian14).Value = Val(ScriviValore(rsGF_WBS!Pian14))
            ws.Cells(rigainizio, colEff14).Value = Val(ScriviValore(rsGF_WBS!Eff14))
            ws.Cells(rigainizio, colDiff14).Value = Val(ScriviValore(rsGF_WBS!Diff14))
            
            ws.Cells(rigainizio, colPian15).Value = Val(ScriviValore(rsGF_WBS!Pian15))
            ws.Cells(rigainizio, colEff15).Value = Val(ScriviValore(rsGF_WBS!Eff15))
            ws.Cells(rigainizio, colDiff15).Value = Val(ScriviValore(rsGF_WBS!Diff15))
            
            ws.Cells(rigainizio, colPian16).Value = Val(ScriviValore(rsGF_WBS!Pian16))
            ws.Cells(rigainizio, colEff16).Value = Val(ScriviValore(rsGF_WBS!Eff16))
            ws.Cells(rigainizio, colDiff16).Value = Val(ScriviValore(rsGF_WBS!Diff16))
            ws.Cells(rigainizio, colStock16).Value = Val(ScriviValore(rsGF_WBS!Stock16))
            
            ws.Cells(rigainizio, colPian19).Value = Val(ScriviValore(rsGF_WBS!Pian19))
            ws.Cells(rigainizio, colEff19).Value = Val(ScriviValore(rsGF_WBS!Eff19))
            ws.Cells(rigainizio, colDiff19).Value = Val(ScriviValore(rsGF_WBS!Diff19))
            
            ws.Cells(rigainizio, colPian20).Value = Val(ScriviValore(rsGF_WBS!Pian20))
            ws.Cells(rigainizio, colEff20).Value = Val(ScriviValore(rsGF_WBS!Eff20))
            ws.Cells(rigainizio, colDiff20).Value = Val(ScriviValore(rsGF_WBS!Diff20))
            
            ws.Cells(rigainizio, colPian21).Value = Val(ScriviValore(rsGF_WBS!Pian21))
            ws.Cells(rigainizio, colEff21).Value = Val(ScriviValore(rsGF_WBS!Eff21))
            ws.Cells(rigainizio, colDiff21).Value = Val(ScriviValore(rsGF_WBS!Diff21))
            
            ws.Cells(rigainizio, colPian22).Value = Val(ScriviValore(rsGF_WBS!Pian22))
            ws.Cells(rigainizio, colEff22).Value = Val(ScriviValore(rsGF_WBS!Eff22))
            ws.Cells(rigainizio, colDiff22).Value = Val(ScriviValore(rsGF_WBS!Diff22))
            
            ws.Cells(rigainizio, colPian23).Value = Val(ScriviValore(rsGF_WBS!Pian23))
            ws.Cells(rigainizio, colEff23).Value = Val(ScriviValore(rsGF_WBS!Eff23))
            ws.Cells(rigainizio, colDiff23).Value = Val(ScriviValore(rsGF_WBS!Diff23))
            
            ws.Cells(rigainizio, colPian25).Value = Val(ScriviValore(rsGF_WBS!Pian25))
            ws.Cells(rigainizio, colEff25).Value = Val(ScriviValore(rsGF_WBS!Eff25))
            ws.Cells(rigainizio, colDiff25).Value = Val(ScriviValore(rsGF_WBS!Diff25))
            
            riga = riga + 1
            rs.Close
            rsGF_WBS.Close
        End If
        rsElenco.MoveNext
    Wend
    rsElenco.Close
    ws.Outline.ShowLevels RowLevels:=1
    Form_Maschera1.lbAvanzamento.Caption = "Fine"
    'xlsapp.Visible = True
End Sub

Function EXCEL_Network_ElencoVC(ByVal Tabella As DAO.Recordset) As String
    Dim strTemp As String
    Dim CRITERIO_RICERCA As String
    strTemp = ""
    CRITERIO_RICERCA = "Network='" & tabLavori("Network") & "'"
    Tabella.FindFirst CRITERIO_RICERCA
    While Not Tabella.NoMatch
        strTemp = strTemp & Tabella("NumeroVC") & ", " & Chr(10)
        Tabella.FindNext CRITERIO_RICERCA
    Wend
    If strTemp <> "" Then strTemp = Left(strTemp, Len(strTemp) - 3)
    EXCEL_Network_ElencoVC = strTemp
End Function


Function EXCEL_Network_ElencoIncaricati(ByVal Tabella As DAO.Recordset) As String
    Dim strTemp As String
    Dim CRITERIO_RICERCA As String
    strTemp = ""
    CRITERIO_RICERCA = "Lavoro='" & tabLavori("Descrizione Lavori") & "'"
    Tabella.FindFirst CRITERIO_RICERCA
    While Not Tabella.NoMatch
        strTemp = strTemp & Tabella("Cognome_e_Nome") & ", "
        Tabella.FindNext CRITERIO_RICERCA
    Wend
    If strTemp <> "" Then strTemp = Left(strTemp, Len(strTemp) - 2)
    EXCEL_Network_ElencoIncaricati = strTemp
End Function

Function EXCEL_LetteraColonnaDaNumero(ByVal NumeroColonna As Integer) As String
    EXCEL_LetteraColonnaDaNumero = Split(Cells(1, NumeroColonna).Address, "$")(1)
End Function

Sub EXCEL_LI_ImpostaFormatoFoglio()
    Set ws = xlsapp.Sheets("Lettere incarico")
    Set wb = xlsapp.Workbooks("Riepilogo lavori.xlsx")
    ws.Activate
    ws.Cells.Select
    With xlsapp.Selection
        .VerticalAlignment = xlCenter
    End With
    With xlsapp.Selection.Font
        .Name = "Calibri"
        .Size = 12
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("LI")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("DataLI")).Select
    With xlsapp.Selection
        .NumberFormat = EXCEL_FORMATO_DATA
        .HorizontalAlignment = xlLeft
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("SpecLI")).Select
    With xlsapp.Selection
        .WrapText = True
    End With

    ws.Columns(EXCEL_ColonnaDaNome("Referenza")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("RL")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
   
    ws.Columns(EXCEL_ColonnaDaNome("PMLI")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
        
    ws.Columns(EXCEL_ColonnaDaNome("LITerminata")).Select
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(EXCEL_ColonnaDaNome("NTWTotali")).Select
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(EXCEL_ColonnaDaNome("NTWConcluse")).Select
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(EXCEL_ColonnaDaNome("NTWChiuse")).Select
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With

    ws.Columns(EXCEL_ColonnaDaNome("OggettoLI")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("Assegnato")).Select
    xlsapp.Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    
    ws.Columns(EXCEL_ColonnaDaNome("PianificatoLI")).Select
    xlsapp.Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    
    ws.Columns(EXCEL_ColonnaDaNome("EffettuatoLI")).Select
    xlsapp.Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    
    ws.Columns(EXCEL_ColonnaDaNome("PercPianificati")).Select
    xlsapp.Selection.Style = "Percent"
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With
    
    xlsapp.Selection.FormatConditions.Delete
    xlsapp.Selection.FormatConditions.AddDatabar
    xlsapp.Selection.FormatConditions(xlsapp.Selection.FormatConditions.Count).ShowValue = True
    xlsapp.Selection.FormatConditions(xlsapp.Selection.FormatConditions.Count).SetFirstPriority
    With xlsapp.Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
    End With
    With xlsapp.Selection.FormatConditions(1).BarColor
        .Color = 8700771
        .TintAndShade = 0
    End With
    xlsapp.Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
    xlsapp.Selection.FormatConditions(1).Direction = xlContext
    xlsapp.Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    xlsapp.Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
    xlsapp.Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With xlsapp.Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With xlsapp.Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With

    ws.Columns(EXCEL_ColonnaDaNome("PercEffettuati")).Select
    xlsapp.Selection.Style = "Percent"
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With
    xlsapp.Selection.FormatConditions.Delete
    xlsapp.Selection.FormatConditions.AddDatabar
    xlsapp.Selection.FormatConditions(xlsapp.Selection.FormatConditions.Count).ShowValue = True
    xlsapp.Selection.FormatConditions(xlsapp.Selection.FormatConditions.Count).SetFirstPriority
    With xlsapp.Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
    End With
    With xlsapp.Selection.FormatConditions(1).BarColor
        .Color = 8700771
        .TintAndShade = 0
    End With
    xlsapp.Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
    xlsapp.Selection.FormatConditions(1).Direction = xlContext
    xlsapp.Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    xlsapp.Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
    xlsapp.Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With xlsapp.Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With xlsapp.Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
End Sub


Sub EXCEL_Network_ImpostaFormatoFoglio()
    Set ws = xlsapp.Sheets("Network")
    Set wb = xlsapp.Workbooks("Riepilogo lavori.xlsx")
    ws.Cells.Select
    With xlsapp.Selection
        .VerticalAlignment = xlCenter
    End With
    With xlsapp.Selection.Font
        .Name = "Calibri"
        .Size = 12
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("SpecNTW")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("LINTW")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
        
    ws.Columns(EXCEL_ColonnaDaNome("StatoNetwork")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("PMNTW")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
   
    ws.Columns(EXCEL_ColonnaDaNome("Contratto")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("DataContratto")).Select
    With xlsapp.Selection
        .NumberFormat = EXCEL_FORMATO_DATA
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("Impresa")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("DataCIG")).Select
    With xlsapp.Selection
        .NumberFormat = EXCEL_FORMATO_DATA
    End With


    ws.Columns(EXCEL_ColonnaDaNome("LavoriTerminati")).Select
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("NetworkChiusa")).Select
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("DescrizioneLavori")).Select
    With xlsapp.Selection
        .WrapText = True
    End With
    
    
    ws.Columns(EXCEL_ColonnaDaNome("PianificatoNTW")).Select
    xlsapp.Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    
    ws.Columns(EXCEL_ColonnaDaNome("EffettuatoNTW")).Select
    xlsapp.Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    
    ws.Columns(EXCEL_ColonnaDaNome("ImpFinODANTW")).Select
    xlsapp.Selection.NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    
    ws.Columns(EXCEL_ColonnaDaNome("PercentualeNTW")).Select
    xlsapp.Selection.Style = "Percent"
    With xlsapp.Selection
        .HorizontalAlignment = xlCenter
    End With
    
    xlsapp.Selection.FormatConditions.Delete
    xlsapp.ActiveWorkbook.KeepChangeHistory = True
    xlsapp.Selection.FormatConditions.AddDatabar
    xlsapp.Selection.FormatConditions(xlsapp.Selection.FormatConditions.Count).ShowValue = True
    xlsapp.Selection.FormatConditions(xlsapp.Selection.FormatConditions.Count).SetFirstPriority
    With xlsapp.Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
    End With
    With xlsapp.Selection.FormatConditions(1).BarColor
        .Color = 8700771
        .TintAndShade = 0
    End With
    xlsapp.Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
    xlsapp.Selection.FormatConditions(1).Direction = xlContext
    xlsapp.Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    xlsapp.Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
    xlsapp.Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With xlsapp.Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With xlsapp.Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    
    ws.Columns(EXCEL_ColonnaDaNome("DL")).Select
    With xlsapp.Selection
        .WrapText = True
    End With

    ws.Columns(EXCEL_ColonnaDaNome("AS")).Select
    With xlsapp.Selection
        .WrapText = True
    End With

End Sub

Sub EXCEL_Network_e_LI_CancellaTuttiIDati()
    Set ws = xlsapp.ActiveSheet
    Set rg = ws.Rows("3:" & ws.UsedRange.Rows.Count)
    rg.Clear
End Sub

Function EXCEL_ColonnaDaNome(ByVal Nome As String) As Long
    Set wb = xlsapp.ActiveWorkbook
    EXCEL_ColonnaDaNome = wb.Names(Nome).RefersToRange.Column
End Function

Sub btAggiornaTuttoClick()
    Set xlsapp = New EXCEL.Application
    xlsapp.Visible = True
    connetti
    If Form_Maschera1.cbNTW.Value = True Then AggiornaNTW
    If Form_Maschera1.cbOpNTW.Value = True Then AggiornaOpNTW
    If Form_Maschera1.cbMLST.Value = True Then AggiornaMLST
    If Form_Maschera1.cbWBS.Value = True Then AggiornaWBS
    If Form_Maschera1.cbODA.Value = True Then AggiornaODA
    If Form_Maschera1.cbConferme.Value = True Then AggiornaConferme
    If Form_Maschera1.cbPresenze.Value = True Then AggiornaPresenze
    If Form_Maschera1.cbTimesheet.Value = True Then AggiornaTimeSheet
    If Form_Maschera1.cbGFWBS.Value = True Then
        AvviaGestFondWBS
        AggiornaGestFondWBS
        AggiornaGestFondWBSCap13
        AggiornaGestFondWBSCap "14"
        AggiornaGestFondWBSCap "15"
        AggiornaGestFondWBSCap "16"
        AggiornaGestFondWBSCap "19"
        AggiornaGestFondWBSCap "20"
        AggiornaGestFondWBSCap "21"
        AggiornaGestFondWBSCap "22"
        AggiornaGestFondWBSCap "23"
        AggiornaGestFondWBSCap "25"
        AggiornaGestFondWBSCap "ST"
    End If
    If Form_Maschera1.cbGFNTW.Value = True Then
        AvviaGestFondNTW
        AggiornaGestFondNTW
        AggiornaGestFondNTWCap13
        AggiornaGestFondNTWCap "14"
        AggiornaGestFondNTWCap "15"
        AggiornaGestFondNTWCap "16"
        AggiornaGestFondNTWCap "19"
        AggiornaGestFondNTWCap "20"
        AggiornaGestFondNTWCap "21"
        AggiornaGestFondNTWCap "22"
        AggiornaGestFondNTWCap "23"
        AggiornaGestFondNTWCap "25"
        AggiornaGestFondNTWCap "ST"
    End If
    DoCmd.RunMacro "ImportaDatiSAPDaExcel"
    xlsapp.Quit
    
    Set db = CurrentDb
    Dim settimana, anno As Integer
    settimana = DatePart("ww", Now, vbMonday, vbFirstFourDays)
    anno = DatePart("yyyy", Now, vbMonday, vbFirstFourDays)
    
    Set rs = db.OpenRecordset("GF_NTW_Settimana", dbOpenDynaset)
    rs.FindFirst "Settimana=" & settimana & " AND Anno=" & anno
    If rs.NoMatch Then
        DoCmd.OpenQuery "qry_Accoda_GF_NTW_Settimana"
    End If
    
    Set rs = db.OpenRecordset("GF_NTW_15_Settimana", dbOpenDynaset)
    rs.FindFirst "Settimana=" & settimana & " AND Anno=" & anno
    If rs.NoMatch Then
        DoCmd.OpenQuery "qry_Accoda_GF_NTW_15_Settimana"
    End If
    
    DoCmd.OpenQuery "qry_AGGIORNA_LavoriTerminati"
    DoCmd.OpenQuery "qry_AGGIORNA_Incarichi_LavoriTerminati"
    DoCmd.OpenQuery "qry_AGGIORNA_LavoriSenzaIntervento"
    
    rs.Close
    MsgBox "Finito"
End Sub

Sub AggiornaGestFondNTWPeriodi()
    SAP_GestFond_AnnullaDettaglio
    SAP_GestFond_SelezionaVoceNavigazione "Periodo/esercizio"
    SAP_GestFond_SelezionaTestata
    SAP_GestFond_SelezionaVoceElenco_PeriodoPerAnno Year(Now)
    SAP_GestFond_SelezionaVoceNavigazione "Oggetto"
    Dim meseSelezionato As String
    Dim annoSelezionato As String
    Dim periodoSelezionato As String
    periodoSelezionato = SAP_GestFond_LeggiPeriodoSelezionato
    annoSelezionato = Right(periodoSelezionato, 4)
    While annoSelezionato = Year(Now)
        SAP_GestFond_ComprimiTutto
        SAP_GestFond_EsportaExcel
        EXCEL_Apriworkbook_Testo "D:\Documenti\Desktop\x.dat"
        meseSelezionato = Left(periodoSelezionato, 3)
        EXCEL_GestFond_CopiaDatiDaFileEsportato "GFNTW" & meseSelezionato
        EXCEL_ChiudiWorkbook "x.dat"
        SAP_GestFond_PremiFrecciaGiù
        periodoSelezionato = SAP_GestFond_LeggiPeriodoSelezionato
        annoSelezionato = Right(periodoSelezionato, 4)
    Wend
End Sub

Function SAP_GestFond_LeggiPeriodoSelezionato() As String
    SAP_GestFond_LeggiPeriodoSelezionato = session.FindById("/app/con[0]/ses[0]/wnd[0]/usr/lbl[51,3]").Text
End Function

Sub SAP_GestFond_AnnullaDettaglio()
    session.FindById("wnd[0]/usr/lbl[4,6]").SetFocus
    session.FindById("wnd[0]").sendVKey 2
End Sub

Sub SAP_GestFond_PremiFrecciaGiù()
    session.FindById("wnd[0]/usr/lbl[46,3]").SetFocus
    session.FindById("wnd[0]").sendVKey 2
End Sub

Sub AggiornaTimeSheet()
    session.StartTransaction "CAOR"
    SAP_InserisciListaUO
    SAP_TimeSheet_ImpostaLayoutMarchi
    SAP_TimeSheet_ImpostaPeriodoEsercizioInCorso
    SAP_Esegui
    SAP_TimeSheet_SelezionaRigaRossa
    SAP_TimeSheet_EsportaFoglioElettronico "exportTimeSheet.xlsx"
    AggiornaDataEsportazione "TimeSheet"
End Sub

Sub SAP_TimeSheet_EsportaFoglioElettronico(ByVal NomeFileDiDestinazione As String)
    session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").CurrentCellColumn = "LIGHTS" 'Seleziona la prima cella
    session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu 'Preme il tasto destro del mouse per il menu contestuale
    session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL" 'Seleziona la voce di menu Foglio elettronico
    session.FindById("wnd[1]/usr/cmbG_LISTBOX").SetFocus 'Seleziona il check item "Sel. da tutti i formati disponibili / EXCEL XLSX"
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press 'Esegui
    session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = CARTELLA_DATISAP
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = NomeFileDiDestinazione
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press 'Bottone sostituisci
End Sub

Sub SAP_TimeSheet_ImpostaLayoutMarchi()
    session.FindById("wnd[0]/usr/ctxtVARIANT").Text = "/marchi"
End Sub

Sub SAP_TimeSheet_SelezionaRigaRossa()
    session.FindById("wnd[0]/usr/lbl[8,3]").SetFocus
    session.FindById("wnd[0]").sendVKey 2
End Sub

Sub SAP_TimeSheet_ImpostaPeriodoEsercizioInCorso()
    session.FindById("wnd[0]/usr/radLFDJR").Select
End Sub

Sub AggiornaPresenze()
    session.StartTransaction "ZTSTM001"
    SAP_InserisciListaUO
    SAP_Presenze_InserisciMeseEAnno Mese:=Month(Now), anno:=Year(Now)
    SAP_Presenze_DeselezionaRecordDiscordanti
    SAP_Esegui
    SAP_Presenze_EsportaFoglioElettronico "exportPresenzeMeseCorrente.xlsx"
    SAP_FrecciaVerde
    Dim MeseScorso As Date
    MeseScorso = DateAdd(interval:="m", Number:=-1, Date:=Now)
    SAP_Presenze_InserisciMeseEAnno Mese:=Month(MeseScorso), anno:=Year(MeseScorso)
    SAP_Esegui
    SAP_Presenze_EsportaFoglioElettronico "exportPresenzeMesePrecedente.xlsx"
    AggiornaDataEsportazione "Presenze"
End Sub

Sub EXCEL_SostituisciCaratteriNonValidiInTestata(ByVal NumeroRigaTestata As Integer)
    Set ws = xlsapp.ActiveSheet
    ws.Rows(NumeroRigaTestata).Select
    xlsapp.Selection.Replace What:=".", Replacement:="_", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Sub EXCEL_Presenze_SelezionaPrimaCellaLibera()
    Set ws = xlsapp.ActiveSheet
    Dim primacolonnadati As Integer
    primacolonnadati = EXCEL_TrovaIndicePrimaColonnaDatiEsportati
    Dim riga As Integer
    Dim rigatrovata As Integer
    For riga = 1 To ws.UsedRange.Rows.Count
        If ws.Cells(riga, primacolonnadati) = "" Then
            rigatrovata = riga
            Exit For
        End If
    Next riga
    Set rg = ws.Cells(riga, primacolonnadati)
    rg.Select
End Sub

Sub EXCEL_SelezionaSoloDatiDaFileEsportato()
    Set ws = xlsapp.ActiveSheet
    Set rg = Range(ws.Cells(2, "A"), ws.Cells(ws.UsedRange.Rows.Count, "I"))
    rg.Select
End Sub

Sub SAP_FrecciaVerde()
    session.FindById("wnd[0]/tbar[0]/btn[3]").Press
End Sub

Sub SAP_Presenze_DeselezionaRecordDiscordanti()
    session.FindById("wnd[0]/usr/chkP_ERROR").Selected = False
End Sub

Sub SAP_Presenze_InserisciMeseEAnno(ByVal Mese As Integer, ByVal anno As Integer)
    session.FindById("wnd[0]/usr/txtP_ANNO").Text = anno
    session.FindById("wnd[0]/usr/txtP_MESE").Text = Mese
End Sub

Sub AggiornaConferme()
    session.StartTransaction "CN48N"
    SAP_AzzeraListaProgetti
    SAP_AzzeraListaWBS
    SAP_InserisciListaNTW
    SAP_ImpostaLayoutMarchi
    SAP_Esegui
    SAP_EsportaFoglioElettronico "exportConferme.xlsx"
    AggiornaDataEsportazione "Conferme"
End Sub

Sub AggiornaODA()
    session.StartTransaction "ME2J"
    SAP_AzzeraListaProgetti
    SAP_AzzeraListaWBS
    SAP_InserisciListaNTW
    SAP_ODA_ImpostaContenutoLista "ALV"
    SAP_AzzeraListaGruppoAcquisti
    SAP_AzzeraListaDivisione
    SAP_AzzeraListaFornitori
    SAP_ODA_AzzeraListaMateriali
    SAP_Esegui
    SAP_ODA_ImpostaLayoutMarchi
    SAP_ODA_EsportaFoglioElettronico "exportODA.xlsx"
    SAP_ODA_LeggiTutteLeTestate
    AggiornaDataEsportazione "ODA"
End Sub

Sub SAP_ODA_LeggiTutteLeTestate()
    Dim tbl As GuiShell
    Dim btn As GuiButton
    Dim usr As GuiUserArea
    Dim cnt1 As GuiComponent
    Dim cnt2 As GuiComponent
    Dim cnt3 As GuiComponent
    Dim cnt4 As GuiComponent
    Dim cnt5 As GuiComponent
    Dim cntBtnTestata As GuiComponent
    Dim cntTestata As GuiComponent
    Dim cmp As GuiComponent

    Set tbl = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    tbl.SetCurrentCell 0, "EBELN"
    tbl.DoubleClickCurrentCell
    
    Set db = CurrentDb
    Dim ctxt As GuiCTextField
    Dim txt As GuiTextField
    Dim cb As GuiCheckBox
    Dim rsElencoODA As DAO.Recordset
    Set rsElencoODA = db.OpenRecordset("qry_ODA_COIN_COFM_LOMI", dbOpenDynaset)
    rsElencoODA.MoveFirst
    While Not rsElencoODA.EOF
        session.FindById("wnd[0]/tbar[1]/btn[17]").Press
        session.FindById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").Text = rsElencoODA!ODA
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press

        Dim DATI_CLIENTE As String
        DATI_CLIENTE = "wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT10"
        session.FindById(DATI_CLIENTE).Select 'Seleziona scheda Dati Cliente

        Set ctxt = session.FindById(DATI_CLIENTE & "/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/ctxtEKKO-DTCON")
        dataconsegna = ctxt.Text
        Set ctxt = session.FindById(DATI_CLIENTE & "/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/ctxtEKKO-DTFIN")
        datafinepresunta = ctxt.Text
        Set txt = session.FindById(DATI_CLIENTE & "/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/txtEKKO-TUTIL")
        tempoutile = txt.Text
        Set txt = session.FindById(DATI_CLIENTE & "/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/txtZECI_ACTUAL_SAL-VERSION")
        SALAttuale = txt.Text
        Set cb = session.FindById(DATI_CLIENTE & "/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/chkZECI_ACTUAL_SAL-FINAL")
        SALFinale = cb.Selected
        Set txt = session.FindById(DATI_CLIENTE & "/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1227/ssubCUSTOMER_DATA_HEADER:SAPLXM06:0101/txtZECI_SAP_SAL-STATUS")
        StatoSAL = txt.Text
        ODA = rsElencoODA!ODA
        Set rs = db.OpenRecordset("ODATestata", dbOpenDynaset)
        rs.FindFirst "ODA='" & ODA & "'"
        If rs.NoMatch Then
            rs.AddNew
        Else
            rs.Edit
        End If
        rs!ODA = ODA
        If dataconsegna <> "" Then rs!dataconsegna = Replace(dataconsegna, ".", "/")
        If datafinepresunta <> "" Then rs!datafinepresunta = Replace(datafinepresunta, ".", "/")
        If tempoutile <> "" Then rs!tempoutile = tempoutile
        If SALAttuale <> "" Then rs!SALAttuale = Replace(SALAttuale, ".", "/")
        rs!SALFinale = SALFinale
        If StatoSAL <> "" Then rs!StatoSAL = Replace(StatoSAL, ".", "/")
        rs.Update
        rs.Close
        rsElencoODA.MoveNext
    Wend
    rsElencoODA.Close
End Sub

Sub SAP_ODA_ImpostaLayoutMarchi()
    session.FindById("wnd[0]/tbar[1]/btn[33]").Press 'Bottone Selezionare layout
    Dim riga As Integer
    Dim ID As String
    ID = "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
    riga = SAP_GridView_TrovaRigaPerContenuto(IDGridView:=ID, TitoloColonna:="Layout", ContenutoDaCercare:="/MARCHI")
    session.FindById(ID).SelectedRows = CStr(riga)
    session.FindById(ID).ClickCurrentCell
End Sub

Function SAP_GridView_RigheTotali(ByVal IDGridView As String) As Integer
    Set gw = session.FindById(IDGridView)
    SAP_GridView_RigheTotali = gw.RowCount
End Function

Function SAP_GridView_ColonneTotali(ByVal IDGridView As String) As Integer
    Set gw = session.FindById(IDGridView)
    Dim coll As GuiCollection
    Set coll = gw.ColumnOrder
    SAP_GridView_ColonneTotali = coll.Count
End Function

Function SAP_GridView_TrovaIDColonnaPerTitoloVisualizzato(ByVal IDGridView As String, ByVal TitoloColonna As String) As String
    Set gw = session.FindById(IDGridView)
    Dim ID As String
    Dim IDTrovato As String
    IDTrovato = ""
    For Each elem In gw.ColumnOrder
        ID = gw.GetDisplayedColumnTitle(elem)
        If ID = TitoloColonna Then
            IDTrovato = elem
            Exit For
        End If
    Next
    SAP_GridView_TrovaIDColonnaPerTitoloVisualizzato = IDTrovato
End Function

Function SAP_GridView_TrovaRigaPerContenuto(ByVal IDGridView As String, ByVal TitoloColonna As String, ByVal ContenutoDaCercare As String) As Integer
    Dim riga As Integer
    Dim rigatrovata As Integer
    Set gw = session.FindById(IDGridView)
    rigatrovata = -1
    Dim IDColonna As String
    IDColonna = SAP_GridView_TrovaIDColonnaPerTitoloVisualizzato(IDGridView, TitoloColonna)
    For riga = 0 To SAP_GridView_RigheTotali(IDGridView) - 1
        SAP_GridView_SistemaBugCambioPagina IDGridView, riga
        If gw.GetCellValue(riga, IDColonna) = ContenutoDaCercare Then
            rigatrovata = riga
            Exit For
        End If
    Next riga
    SAP_GridView_TrovaRigaPerContenuto = rigatrovata
End Function

Sub SAP_GridView_SistemaBugCambioPagina(ByVal IDGridView As String, RigaAttuale As Integer)
    Set gw = session.FindById(IDGridView)
    If RigaAttuale Mod gw.VisibleRowCount = 0 Then
        gw.SetCurrentCell RigaAttuale, gw.ColumnOrder(0)
    End If
End Sub

Sub SAP_AzzeraListaMateriali()
    session.FindById("wnd[0]/usr/btn%_CN_MATNR_%_APP_%-VALU_PUSH").Press 'Bottone lista materiali
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone esegui
End Sub

Sub SAP_ODA_AzzeraListaMateriali()
    session.FindById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").Press 'Bottone lista materiali
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone esegui
End Sub

Sub SAP_AzzeraListaFornitori()
    session.FindById("wnd[0]/usr/btn%_S_LIFNR_%_APP_%-VALU_PUSH").Press 'Bottone lista fornitori
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone esegui
End Sub

Sub SAP_AzzeraListaDivisione()
    session.FindById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").Press 'Bottone lista divisione
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone esegui
End Sub

Sub SAP_AzzeraListaGruppoAcquisti()
    session.FindById("wnd[0]/usr/btn%_S_EKGRP_%_APP_%-VALU_PUSH").Press 'Bottone lista gruppo acquisti
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone esegui
End Sub

Sub SAP_ODA_ImpostaContenutoLista(ByVal Testo As String)
    session.FindById("wnd[0]/usr/ctxtLISTU").Text = Testo
End Sub

Sub AvviaGestFondNTW()
    session.StartTransaction "ZGEST_FND"
    SAP_AzzeraListaProgetti
    SAP_AzzeraListaWBS
    SAP_AzzeraListaMateriali
    SAP_InserisciListaNTW
    SAP_SelezionaReportRicercaClassico
    SAP_Esegui
    'SAP_GestFond_ComprimiTutto
    SAP_GestFond_ImpostaParametrizzazione_e_RigaTotali
End Sub

Sub AggiornaGestFondNTW()
    SAP_GestFond_EsportaExcel "GF_NTW.dat"
    EXCEL_Apriworkbook_Testo "GF_NTW.dat"
    EXCEL_AdattaTestataGestFondSuUnaRiga
    EXCEL_SalvaComeXLSX "GF_NTW.xlsx"
    EXCEL_ChiudiWorkbook "GF_NTW.xlsx"
    AggiornaDataEsportazione "GF_NTW"
End Sub

Sub AggiornaGestFondNTWCap13()
    SAP_GestFond_VisualizzaCatValori "CAP-PER-13"
    'SAP_GestFond_ComprimiTutto
    SAP_GestFond_EsportaExcel "GF_NTW_13.dat"
    EXCEL_Apriworkbook_Testo "GF_NTW_13.dat"
    EXCEL_AdattaTestataGestFondSuUnaRiga
    EXCEL_SalvaComeXLSX "GF_NTW_13.xlsx"
    EXCEL_ChiudiWorkbook "GF_NTW_13.xlsx"
    AggiornaDataEsportazione "GF_NTW_13"
End Sub

Sub AggiornaGestFondNTWCap(ByVal Capitolo As String)
    If SAP_GestFond_TrovaCapitolo(Capitolo) Then
        'SAP_GestFond_ComprimiTutto
        SAP_GestFond_EsportaExcel "GF_NTW_" & Capitolo & ".dat"
        EXCEL_Apriworkbook_Testo "GF_NTW_" & Capitolo & ".dat"
        EXCEL_AdattaTestataGestFondSuUnaRiga
    Else
        DoCmd.RunSQL "DELETE FROM GF_NTW_" & Capitolo
        EXCEL_Apriworkbook "GF_NTW_" & Capitolo & ".xlsx"
        EXCEL_SelezionaRighe_da_2_a_fine
        EXCEL_CancellaSelezione
    End If
    EXCEL_SalvaComeXLSX "GF_NTW_" & Capitolo & ".xlsx"
    EXCEL_ChiudiWorkbook "GF_NTW_" & Capitolo & ".xlsx"
    AggiornaDataEsportazione "GF_NTW_" & Capitolo & ""
End Sub

Sub SAP_GestFond_ComprimiTutto()
    SAP_GestFond_PremiMeno
    SAP_GestFond_PremiMeno
    SAP_GestFond_PremiMeno
    SAP_GestFond_PremiMeno
End Sub

Sub SAP_GestFond_VisualizzaCatValori(ByVal NomeCapitolo As String)
    SAP_GestFond_SelezionaVoceNavigazione "Cat. valori"
    SAP_GestFond_SelezionaTestata
    SAP_GestFond_SelezionaVoceElenco NomeCapitolo 'Mette Oggetto sulla freccetta del capitolo desiderato
    SAP_GestFond_SelezionaVoceNavigazione "Oggetto"
End Sub

Sub SAP_GestFond_VisualizzaPeriodoEsercizio(ByVal NomePeriodo As String)
    SAP_GestFond_SelezionaTestata
    SAP_GestFond_SelezionaVoceNavigazione "Periodo/esercizio"
    SAP_GestFond_SelezionaTestata 'Mette Periodo/esercizio sulla testata
    SAP_GestFond_SelezionaVoceNavigazione "Oggetto"
    SAP_GestFond_SelezionaVoceElenco NomePeriodo 'Mette Oggetto sulla freccetta del NomePeriodo desiderato
End Sub

Sub SAP_GestFond_SelezionaVoceElenco(ByVal NomeVoce As String)
    Dim trovato As Boolean
    trovato = False
    Dim i As Integer
    i = 13
    While Not trovato
        On Error GoTo finepagina:
        If session.FindById("wnd[0]/usr/lbl[2," & i & "]").Text = NomeVoce Then
            session.FindById("wnd[0]/usr/lbl[1," & i & "]").SetFocus 'Posiziona il mouse sulla voce
            session.FindById("wnd[0]").sendVKey 2 'Tasto destro del mouse
            trovato = True
        End If
        i = i + 1
        If SAP_GestFond_VerificaNecessitàCambioPagina(RigaAttuale:=i) = True Then
            SAP_btPaginaGiù
            i = 13
        End If
    Wend
finepagina:
End Sub

Sub SAP_GestFond_SelezionaVoceElenco_PeriodoPerAnno(ByVal anno As Integer)
    Dim trovato As Boolean
    trovato = False
    Dim i As Integer
    i = 13
    While Not trovato
        On Error GoTo finepagina:
        If Right(session.FindById("wnd[0]/usr/lbl[2," & i & "]").Text, 4) = CStr(anno) Then
            session.FindById("wnd[0]/usr/lbl[1," & i & "]").SetFocus 'Posiziona il mouse sulla voce
            session.FindById("wnd[0]").sendVKey 2 'Tasto destro del mouse
            trovato = True
        End If
        i = i + 1
        If SAP_GestFond_VerificaNecessitàCambioPagina(RigaAttuale:=i) = True Then
            SAP_btPaginaGiù
            i = 13
        End If
    Wend
finepagina:
End Sub

Sub SAP_btPaginaGiù()
    session.FindById("wnd[0]/tbar[0]/btn[82]").Press
End Sub

Function SAP_GestFond_VerificaNecessitàCambioPagina(ByVal RigaAttuale As Integer) As Boolean
    Dim Risultato As Boolean
    Risultato = False
    Dim vb As GuiScrollbar
    Set vb = session.FindById("wnd[0]/usr").VerticalScrollbar
    If (RigaAttuale > vb.PageSize) And (vb.Position < vb.Maximum) Then
        Risultato = True
    End If
    SAP_GestFond_VerificaNecessitàCambioPagina = Risultato
End Function

Sub SAP_GestFond_SelezionaVoceNavigazione(ByVal NomeVoce As String)
    For i = 3 To 6
        If session.FindById("wnd[0]/usr/lbl[1," & i & "]").Text = NomeVoce Then
            session.FindById("wnd[0]/usr/lbl[1," & i & "]").SetFocus 'Posiziona il mouse sulla voce
            session.FindById("wnd[0]").sendVKey 2 'Tasto destro del mouse
        End If
    Next i
End Sub

Sub EXCEL_ImpostaFormato(ByVal formato As String)
    xlsapp.Selection.NumberFormat = formato
End Sub

Sub AvviaGestFondWBS()
    session.StartTransaction "ZGEST_FND"
    SAP_AzzeraListaProgetti
    SAP_AzzeraListaNTW
    SAP_AzzeraListaMateriali
    SAP_InserisciListaWBS
    SAP_SelezionaReportRicercaClassico
    SAP_Esegui
    SAP_GestFond_ComprimiTutto
    SAP_GestFond_ImpostaParametrizzazione_e_RigaTotali
End Sub

Sub AggiornaGestFondWBS()
    SAP_GestFond_EsportaExcel "GF_WBS.dat"
    EXCEL_Apriworkbook_Testo "GF_WBS.dat"
    EXCEL_AdattaTestataGestFondSuUnaRiga
    EXCEL_SalvaComeXLSX "GF_WBS.xlsx"
    EXCEL_ChiudiWorkbook "GF_WBS.xlsx"
    AggiornaDataEsportazione "GF_WBS"
End Sub
    
Sub AggiornaGestFondWBSCap13()
    SAP_GestFond_VisualizzaCatValori "CAP-PER-13"
    SAP_GestFond_ComprimiTutto
    SAP_GestFond_EsportaExcel "GF_WBS_13.dat"
    EXCEL_Apriworkbook_Testo "GF_WBS_13.dat"
    EXCEL_AdattaTestataGestFondSuUnaRiga
    EXCEL_SalvaComeXLSX "GF_WBS_13.xlsx"
    EXCEL_ChiudiWorkbook "GF_WBS_13.xlsx"
    AggiornaDataEsportazione "GF_WBS_13"
End Sub

Sub AggiornaGestFondWBSCap(ByVal Capitolo As String)
    SAP_GestFond_PremiFrecciaGiù
    If SAP_GestFond_TrovaCapitolo(Capitolo) Then
        SAP_GestFond_ComprimiTutto
        SAP_GestFond_EsportaExcel "GF_WBS_" & Capitolo & ".dat"
        EXCEL_Apriworkbook_Testo "GF_WBS_" & Capitolo & ".dat"
        EXCEL_AdattaTestataGestFondSuUnaRiga
    Else
        DoCmd.RunSQL "DELETE FROM GF_WBS_" & Capitolo
        EXCEL_Apriworkbook "GF_WBS_" & Capitolo & ".xlsx"
        EXCEL_SelezionaRighe_da_2_a_fine
        EXCEL_CancellaSelezione
    End If
    EXCEL_SalvaComeXLSX "GF_WBS_" & Capitolo & ".xlsx"
    EXCEL_ChiudiWorkbook "GF_WBS_" & Capitolo & ".xlsx"
    AggiornaDataEsportazione "GF_WBS_" & Capitolo & ""
End Sub

Function SAP_GestFond_TrovaCapitolo(ByVal Capitolo As String) As Boolean
    Dim lbl As GuiLabel
    Set lbl = session.FindById("wnd[0]/usr/lbl[51,3]")
    While lbl.Text <> ("CAP-PER-" & Capitolo) And lbl.Text <> "CAP-PER-ST" And Right(lbl.Text, 2) < Capitolo
        SAP_GestFond_PremiFrecciaGiù
        Set lbl = session.FindById("wnd[0]/usr/lbl[51,3]")
    Wend
    If lbl.Text = ("CAP-PER-" & Capitolo) Then SAP_GestFond_TrovaCapitolo = True Else SAP_GestFond_TrovaCapitolo = False
End Function

Sub EXCEL_SalvaComeXLSX(ByVal NomeFile As String)
    xlsapp.DisplayAlerts = False
    xlsapp.ActiveWorkbook.SaveAs FileName:=CARTELLA_DATISAP & NomeFile, FileFormat:=xlOpenXMLWorkbook, ConflictResolution:=xlLocalSessionChanges
    xlsapp.DisplayAlerts = True
End Sub

Sub EXCEL_AdattaTestataGestFondSuUnaRiga()
    Set ws = xlsapp.ActiveSheet
    Set rg = ws.Rows(1)
    rg.Select
    xlsapp.Selection.Delete Shift:=xlUp
    
    ws.Cells(1, 2).Value = "TotPianificato"
    ws.Cells(1, 3).Value = "TotBudget"
    ws.Cells(1, 4).Value = "TotRichiestaFondi"
    ws.Cells(1, 5).Value = "TotEffettuato"
    ws.Cells(1, 6).Value = "TotImpFinODA"
    ws.Cells(1, 7).Value = "TotDisposto"
    ws.Cells(1, 8).Value = "TotDisponibile"
    
    ws.Cells(1, 9).Value = "EsPrecPianificato"
    ws.Cells(1, 10).Value = "EsPrecBudget"
    ws.Cells(1, 11).Value = "EsPrecRichiestaFondi"
    ws.Cells(1, 12).Value = "EsPrecEffettuato"
    ws.Cells(1, 13).Value = "EsPrecImpFinODA"
    ws.Cells(1, 14).Value = "EsPrecDisposto"
    ws.Cells(1, 15).Value = "EsPrecDisponibile"

    ws.Cells(1, 16).Value = "CorsoPianificato"
    ws.Cells(1, 17).Value = "CorsoBudget"
    ws.Cells(1, 18).Value = "CorsoRichiestaFondi"
    ws.Cells(1, 19).Value = "CorsoEffettuato"
    ws.Cells(1, 20).Value = "CorsoImpFinODA"
    ws.Cells(1, 21).Value = "CorsoDisposto"
    ws.Cells(1, 22).Value = "CorsoDisponibile"

    ws.Cells(1, 23).Value = "PiùUnoPianificato"
    ws.Cells(1, 24).Value = "PiùUnoBudget"
    ws.Cells(1, 25).Value = "PiùUnoRichiestaFondi"
    ws.Cells(1, 26).Value = "PiùUnoEffettuato"
    ws.Cells(1, 27).Value = "PiùUnoImpFinODA"
    ws.Cells(1, 28).Value = "PiùUnoDisposto"
    ws.Cells(1, 29).Value = "PiùUnoDisponibile"

    ws.Cells(1, 30).Value = "PiùDuePianificato"
    ws.Cells(1, 31).Value = "PiùDueBudget"
    ws.Cells(1, 32).Value = "PiùDueRichiestaFondi"
    ws.Cells(1, 33).Value = "PiùDueEffettuato"
    ws.Cells(1, 34).Value = "PiùDueImpFinODA"
    ws.Cells(1, 35).Value = "PiùDueDisposto"
    ws.Cells(1, 36).Value = "PiùDueDisponibile"

    ws.Cells(1, 37).Value = "OltreDuePianificato"
    ws.Cells(1, 38).Value = "OltreDueBudget"
    ws.Cells(1, 39).Value = "OltreDueRichiestaFondi"
    ws.Cells(1, 40).Value = "OltreDueEffettuato"
    ws.Cells(1, 41).Value = "OltreDueImpFinODA"
    ws.Cells(1, 42).Value = "OltreDueDisposto"
    ws.Cells(1, 43).Value = "OltreDueDisponibile"

    ws.Cells(1, 44).Value = "TotEsPianificato"
    ws.Cells(1, 45).Value = "TotEsBudget"
    ws.Cells(1, 46).Value = "TotEsRichiestaFondi"
    ws.Cells(1, 47).Value = "TotEsEffettuato"
    ws.Cells(1, 48).Value = "TotEsImpFinODA"
    ws.Cells(1, 49).Value = "TotEsDisposto"
    ws.Cells(1, 50).Value = "TotEsDisponibile"
End Sub

Sub EXCEL_Apriworkbook_Testo(ByVal NomeFile As String)
    xlsapp.Workbooks.OpenText FileName:=CARTELLA_DATISAP & NomeFile, TextQualifier:=xlNone
End Sub

Sub EXCEL_GestFond_CopiaDatiDaFileEsportato(ByVal NomeFoglioDiDestinazione As String)
    EXCEL_SelezionaFinestra "x.dat"
    EXCEL_SelezionaTutto
    EXCEL_CopiaSelezione
    EXCEL_SelezionaFinestra "Riepilogo lavori.xlsm"
    EXCEL_SelezionaFoglio NomeFoglioDiDestinazione
    EXCEL_SelezionaTutto
    EXCEL_IncollaSelezione
End Sub

Sub EXCEL_SelezionaPrimaCellaDati()
    Set ws = xlsapp.ActiveSheet
    ws.Cells(5, 2).Activate
End Sub

Sub EXCEL_CancellaSelezione()
    xlsapp.Selection.Clear
End Sub

Sub EXCEL_CancellaTutto()
    xlsapp.Cells.Clear
End Sub

Sub EXCEL_IncollaSpecialeNumeri()
    Set ws = xlsapp.ActiveSheet
    ws.Cells(1, 1).Activate
    xlsapp.ActiveCell.PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:=False, Transpose:=False
End Sub

Sub EXCEL_GestFond_SelezionaSoloDati()
    Set ws = xlsapp.ActiveSheet
    Set rg = Range(ws.Cells(5, 2), ws.Cells(ws.UsedRange.Rows.Count, ws.UsedRange.Columns.Count))
    rg.Select
End Sub

Sub EXCEL_SelezionaTutto()
    Set ws = xlsapp.ActiveSheet
    ws.Cells.Select
End Sub

Sub SAP_GestFond_EsportaExcel(ByVal NomeFileDiDestinazione As String)
    session.FindById("wnd[0]/tbar[1]/btn[47]").Press 'Bottone esporta
    session.FindById("wnd[1]/usr/chkCFDOWNLOAD-HEADER").Selected = False
    session.FindById("wnd[1]/usr/chkCFDOWNLOAD-COL_HEADER").Selected = True
    session.FindById("wnd[1]/usr/chkCFDOWNLOAD-FORMAT_DIS").Selected = False
    session.FindById("wnd[1]/usr/chkCFDOWNLOAD-VALUE_AREA").Selected = True
    session.FindById("wnd[1]/usr/ctxtCFDOWNLOAD-FILE").Text = CARTELLA_DATISAP & NomeFileDiDestinazione 'File output
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press 'Bottone esegui
    session.FindById("wnd[2]/usr/btnSPOP-VAROPTION1").Press 'Bottone Sostituire
End Sub

Sub SAP_GestFond_ImpostaParametrizzazione_e_RigaTotali()
    SAP_GestFond_SelezionaTestata
    SAP_GestFond_ImpostaParametrizzazione_RapprCaratteristiche_ColonnaChiave
    SAP_GestFond_SelezionaTestata
    SAP_GestFond_ImpostaRigaTotali
End Sub

Sub SAP_GestFond_ImpostaRigaTotali()
    session.FindById("wnd[0]/mbar/menu[5]/menu[3]").Select 'Menu Riga Totali
    session.FindById("wnd[1]/usr/sub:SAPLKEC1:0110/radCEC01-CHOICE[0,0]").Select 'Check Omettere riga totali
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press 'Bottone esegui
End Sub

Sub SAP_GestFond_PremiMeno()
    session.FindById("wnd[0]/tbar[1]/btn[45]").Press
End Sub

Sub SAP_GestFond_ImpostaParametrizzazione_RapprCaratteristiche_ColonnaChiave()
    session.FindById("wnd[0]/mbar/menu[5]/menu[2]").Select 'Menu Parametrizzazione_RapprCaratteristiche
    session.FindById("wnd[1]/usr/btnD2000_PUSH_01").Press 'Bottone Colonna Chiave
    session.FindById("wnd[1]/usr/subSSC_LC:SAPLKYAT:0210/tblSAPLKYATTC_LC").GetAbsoluteRow(1).Selected = True 'Seleziona seconda riga (Descrizione)
    session.FindById("wnd[1]/usr/btnPUSH_DESE").Press 'Bottone freccia a destra
    session.FindById("wnd[1]/tbar[0]/btn[0]").Press 'Bottone Rilevare
End Sub

Function SAP_GestFond_TrovaLabelPerTesto(ByVal Testo As String) As GuiLabel
    Dim usr As GuiUserArea
    Set usr = session.FindById("/app/con[0]/ses[0]/wnd[0]/usr")
    Dim lbl As GuiLabel
    For Each lbl In usr.Children
        If lbl.Text = Testo Then
            Exit For
        End If
    Next lbl
    Set SAP_GestFond_TrovaLabelPerTesto = lbl
End Function

Sub SAP_GestFond_SelezionaTestata()
    'La testata si trova più o meno entro questi limiti di posizione
    Const CharLeft_minimo = 1
    Const CharTop_minimo = 8
    Const CharLeft_massimo = 15
    Const CharTop_massimo = 14
    
    Dim usr As GuiUserArea
    Set usr = session.FindById("wnd[0]/usr")
    Dim lbl As GuiLabel
    For Each lbl In usr.Children
        If (lbl.Text <> "") And (lbl.CharLeft >= CharLeft_minimo) And (lbl.CharTop >= CharTop_minimo) And (lbl.CharLeft <= CharLeft_massimo) And (lbl.CharTop <= CharTop_massimo) Then
            Exit For
        End If
    Next lbl
    lbl.SetFocus
    
    session.FindById("wnd[0]").sendVKey 2 'Preme il tasto del mouse
End Sub

Sub SAP_SelezionaReportRicercaClassico()
    session.FindById("wnd[0]/usr/radLISTE").Select
End Sub

Sub AggiornaWBS()
    session.StartTransaction "CN43N"
    SAP_AzzeraListaProgetti
    SAP_InserisciListaWBS
    SAP_ImpostaLayoutMarchi
    SAP_Esegui
    SAP_EsportaFoglioElettronico "exportWBS.xlsx"
    AggiornaDataEsportazione "WBS"
End Sub

Sub SAP_InserisciListaWBS()
    ACCESS_EsportaListaWBSsuFileDiTesto
    SAP_IncollaListaWBSdaFileDiTesto
End Sub

Sub SAP_IncollaListaWBSdaFileDiTesto()
    session.FindById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").Press 'Bottone elenco WBS
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[23]").Press 'Bottone incolla da file di testo
    session.FindById("wnd[2]/usr/ctxtDY_PATH").Text = CARTELLA_DATISAP
    session.FindById("wnd[2]/usr/ctxtDY_FILENAME").Text = "ElencoWBS.txt"
    session.FindById("wnd[2]/tbar[0]/btn[0]").Press
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone SAP_Esegui
End Sub

Sub ACCESS_EsportaListaWBSsuFileDiTesto()
    DoCmd.RunMacro "EsportaElencoWBS"
End Sub

Sub EXCEL_CopiaListaWBSsuClipboard()
    EXCEL_SelezionaFoglio "OpNTW"
    EXCEL_CancellaFiltro
    EXCEL_SelezionaColonnaPerTitolo "senza com.", 1
    EXCEL_CopiaSelezione
End Sub

Sub SAP_IncollaListaWBSdaClipboard()
    session.FindById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").Press 'Bottone elenco WBS
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[24]").Press 'Bottone incolla da clipboard
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone SAP_Esegui
End Sub

Sub AggiornaMLST()
    session.StartTransaction "CN53N"
    SAP_AzzeraListaProgetti
    SAP_AzzeraListaWBS
    SAP_InserisciListaNTW
    SAP_ImpostaLayoutMarchi
    SAP_Esegui
    SAP_EsportaFoglioElettronico "exportMLST.xlsx"
    AggiornaDataEsportazione "MLST"
End Sub

Sub AggiornaOpNTW()
    session.StartTransaction "CN47N"
    SAP_AzzeraListaProgetti
    SAP_AzzeraListaWBS
    SAP_InserisciListaNTW
    SAP_ImpostaLayoutMarchi
    SAP_Esegui
    SAP_EsportaFoglioElettronico "exportOpNTW.xlsx"
    AggiornaDataEsportazione "OpNTW"
End Sub

Sub AggiornaNTW()
    session.StartTransaction "CN46N"
    SAP_AzzeraListaProgetti
    SAP_AzzeraListaWBS
    SAP_InserisciListaNTW
    SAP_ImpostaLayoutMarchi
    SAP_Esegui
    SAP_EsportaFoglioElettronico "exportNTW.xlsx"
    AggiornaDataEsportazione "NTW"
End Sub

Sub EXCEL_ChiudiWorkbook(ByVal NomeWorkBook As String)
    Set wb = xlsapp.Workbooks(NomeWorkBook)
    wb.Close SaveChanges:=False
End Sub

Sub EXCEL_Apriworkbook(ByVal NomeFile As String)
    xlsapp.Workbooks.Open FileName:=CARTELLA_DATISAP & NomeFile
End Sub

Sub EXCEL_IncollaSelezione()
    xlsapp.ActiveSheet.Paste
End Sub

Sub EXCEL_SelezionaColonne(ByVal IndiceInizio As Integer, ByVal IndiceFine As Integer)
    Set ws = xlsapp.ActiveSheet
    Set rg = ws.Range(ws.Columns(IndiceInizio), ws.Columns(IndiceFine))
    rg.Select
End Sub

Sub EXCEL_SelezionaRighe(ByVal IndiceInizio As Integer, ByVal IndiceFine As Integer)
    Set ws = xlsapp.ActiveSheet
    Set rg = ws.Range(ws.Rows(IndiceInizio), ws.Rows(IndiceFine))
    rg.Select
End Sub

Sub EXCEL_SelezionaRighe_da_2_a_fine()
    Set ws = xlsapp.ActiveSheet
    Set rg = ws.Range(ws.Rows(2), ws.Rows(ws.UsedRange.Rows.Count + 1))
    rg.Select
End Sub

Function EXCEL_TrovaIndiceUltimaColonna() As Integer
    EXCEL_TrovaIndiceUltimaColonna = xlsapp.ActiveSheet.UsedRange.Columns.Count
End Function

Function EXCEL_TrovaIndicePrimaColonnaDatiEsportati() As Integer
    Set ws = ActiveSheet
    Dim Colonna As Integer
    For i = 1 To ws.UsedRange.Columns.Count
        Set rg = ws.Cells(1, i)
        If rg.Interior.Color = 12632256 Then
            Colonna = i
            Exit For
        End If
    Next i
    EXCEL_TrovaIndicePrimaColonnaDatiEsportati = Colonna
End Function

Sub EXCEL_SelezionaTutteColonne()
    Set rg = xlsapp.ActiveSheet.UsedRange.Columns
    rg.Select
End Sub

Sub EXCEL_SelezionaFinestra(ByVal Nome As String)
    Dim i As EXCEL.Window
    For Each i In xlsapp.Windows
        If UCase(i.Caption) = UCase(Nome) Then
            Set wnd = i
            Exit For
        End If
    Next i
    wnd.Activate
End Sub

Sub SAP_Presenze_EsportaFoglioElettronico(ByVal NomeFileDestinazione As String)
    session.FindById("wnd[0]/tbar[1]/btn[14]").Press 'Bottone esporta
    session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = CARTELLA_DATISAP
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = NomeFileDestinazione
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press 'Bottone sovrascrivi
End Sub

Sub SAP_ODA_EsportaFoglioElettronico(ByVal NomeFileDestinazione As String)
    session.FindById("wnd[0]/tbar[1]/btn[43]").Press 'Bottone esporta
    session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = CARTELLA_DATISAP
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = NomeFileDestinazione
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press 'Bottone sovrascrivi
End Sub

Sub SAP_EsportaFoglioElettronico(ByVal NomeFileDestinazione As String)
    session.FindById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT" 'Bottone esporta
    session.FindById("wnd[0]/usr/cntlALVCONTAINER/shellcont/shell").selectContextMenuItem "&XXL" 'Voce menu "Foglio elettronico"
    session.FindById("wnd[1]/usr/ctxtDY_PATH").Text = CARTELLA_DATISAP
    session.FindById("wnd[1]/usr/ctxtDY_FILENAME").Text = NomeFileDestinazione
    session.FindById("wnd[1]/tbar[0]/btn[11]").Press 'Bottone Sostituire
End Sub

Function EXCEL_WorkBookAperto(ByVal NomeWorkBook As String) As Boolean
    Dim Risultato As Boolean
    Risultato = False
    For Each wb In xlsapp.Workbooks
        If wb.Name = NomeWorkBook Then Risultato = True
    Next wb
    EXCEL_WorkBookAperto = Risultato
End Function

Sub SAP_AzzeraListaProgetti()
    session.FindById("wnd[0]/usr/btn%_CN_PROJN_%_APP_%-VALU_PUSH").Press 'Bottone elenco Progetti
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone SAP_Esegui
End Sub

Sub SAP_AzzeraListaWBS()
    session.FindById("wnd[0]/usr/btn%_CN_PSPNR_%_APP_%-VALU_PUSH").Press 'Bottone elenco WBS
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone SAP_Esegui
End Sub

Sub SAP_AzzeraListaNTW()
    session.FindById("wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH").Press 'Bottone elenco network
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone SAP_Esegui
End Sub

Sub SAP_InserisciListaUO()
    ACCESS_EsportaListaUOsuFileDiTesto
    SAP_IncollaListaUOdaFileDiTesto
End Sub

Sub SAP_IncollaListaUOdaFileDiTesto()
    session.FindById("wnd[0]/usr/btn%_SO_OBJID_%_APP_%-VALU_PUSH").Press 'Bottone elenco UO
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[23]").Press 'Bottone incolla da file di testo
    session.FindById("wnd[2]/usr/ctxtDY_PATH").Text = CARTELLA_DATISAP
    session.FindById("wnd[2]/usr/ctxtDY_FILENAME").Text = "ElencoUO.txt"
    session.FindById("wnd[2]/tbar[0]/btn[0]").Press
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press ' Bottone esegui
End Sub

Sub ACCESS_EsportaListaUOsuFileDiTesto()
    DoCmd.RunMacro "EsportaElencoUO"
End Sub

Sub EXCEL_CopiaListaUOsuClipboard()
    EXCEL_SelezionaFoglio "CAT2"
    EXCEL_CancellaFiltro
    EXCEL_SelezionaColonnaPerTitolo NomeColonna:="UO", RigaTestata:=2
    EXCEL_CopiaSelezione
End Sub

Sub SAP_InserisciListaNTW()
    ACCESS_EsportaElencoNTWsuFileDiTesto
    SAP_IncollaListaNTWdaFileDiTesto
End Sub

Sub SAP_IncollaListaNTWdaFileDiTesto()
    session.FindById("wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH").Press 'Bottone elenco network
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[23]").Press 'Bottone incolla da file di testo
    session.FindById("wnd[2]/usr/ctxtDY_PATH").Text = CARTELLA_DATISAP
    session.FindById("wnd[2]/usr/ctxtDY_FILENAME").Text = "ElencoNTW.txt"
    session.FindById("wnd[2]/tbar[0]/btn[0]").Press
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone SAP_Esegui
End Sub

Sub ACCESS_EsportaElencoNTWsuFileDiTesto()
    DoCmd.RunMacro "EsportaElencoNTW"
End Sub

Sub EXCEL_CopiaListaNTWsuClipboard()
    EXCEL_SelezionaFoglio "Network"
    EXCEL_CancellaFiltro
    EXCEL_SelezionaColonnaPerTitolo "Network", 2
    EXCEL_CopiaSelezione
End Sub

Sub EXCEL_SelezionaFoglio(ByVal NomeFoglio As String)
    xlsapp.Sheets(NomeFoglio).Activate
End Sub

Sub EXCEL_CancellaFiltro()
    Dim af As AutoFilter
    Set af = xlsapp.ActiveSheet.AutoFilter
    Set ws = xlsapp.ActiveSheet
    If af.FilterMode Then ws.ShowAllData
End Sub

Sub EXCEL_SelezionaColonnaPerTitolo(ByVal NomeColonna As String, ByVal RigaTestata As Integer)
    Dim Colonna As Integer
    Set ws = xlsapp.ActiveSheet
    For i = 1 To ws.UsedRange.Columns.Count
        If UCase(ws.Cells(RigaTestata, i)) = UCase(NomeColonna) Then
            Colonna = i
            Exit For
        End If
    Next i
    ws.Columns(Colonna).Select
End Sub

Sub EXCEL_CopiaSelezione()
    xlsapp.Selection.Copy
End Sub

Sub SAP_IncollaListaNTWdaClipboard()
    session.FindById("wnd[0]/usr/btn%_CN_NETNR_%_APP_%-VALU_PUSH").Press 'Bottone elenco network
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[24]").Press 'Bottone incolla da clipboard
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press 'Bottone SAP_Esegui
End Sub

Sub SAP_IncollaListaUOdaClipboard()
    session.FindById("wnd[0]/usr/btn%_SO_OBJID_%_APP_%-VALU_PUSH").Press 'Bottone elenco UO
    session.FindById("wnd[1]/tbar[0]/btn[16]").Press 'Bottone cestino
    session.FindById("wnd[1]/tbar[0]/btn[24]").Press 'Bottone incolla da clipboard
    session.FindById("wnd[2]/usr/btnBUTTON_1").Press 'Ignora finestra di errore
    session.FindById("wnd[1]/tbar[0]/btn[8]").Press ' Bottone esegui
End Sub

Sub SAP_ImpostaLayoutMarchi()
    session.FindById("wnd[0]/usr/ctxtP_DISVAR").Text = "/marchi"
End Sub

Sub SAP_Esegui()
    session.FindById("wnd[0]/tbar[1]/btn[8]").Press
End Sub



