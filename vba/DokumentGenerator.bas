' =====================================================
' DokumentGenerator.bas
' VBA-Modul zur automatischen Dokumentgenerierung
' für deutsche Behörden
' =====================================================
Option Explicit

' Konfiguration
Private Const KI_PROXY_URL As String = "http://localhost:5000"
Private Const LOG_DATEI As String = "dokument_log.csv"
Private Const VORLAGEN_ORDNER As String = "templates\"

' =====================================================
' Hauptfunktion: Dokument generieren
' =====================================================
Public Sub GenerateDokument()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Prüfe ob Daten vorhanden
    If ws.Range("A2").Value = "" Then
        MsgBox "Keine Daten in Zeile 2 gefunden.", vbExclamation, "Fehler"
        Exit Sub
    End If
    
    ' Daten aus Excel lesen
    Dim daten As Collection
    Set daten = ReadDataFromExcel(ws)
    
    ' Word-Vorlage füllen
    Dim vorlagenPfad As String
    vorlagenPfad = ws.Range("Q1").Value  ' Zelle Q1 = Vorlagenpfad
    
    If vorlagenPfad = "" Then
        vorlagenPfad = ThisWorkbook.Path & "\" & VORLAGEN_ORDNER & "Bescheid-Vorlage.dotx"
    End If
    
    ' Dokument erstellen
    Dim docPath As String
    docPath = FillWordTemplate(vorlagenPfad, daten, ws)
    
    ' KI-Optimierung (optional)
    Dim kiAktiv As Boolean
    kiAktiv = ws.Range("Q2").Value = "JA"
    
    If kiAktiv Then
        OptimizeWithKI docPath, daten
    End If
    
    ' Protokollieren
    LogGenerierung daten, docPath, kiAktiv
    
    ' Erfolgsmeldung
    MsgBox "Dokument erfolgreich generiert:" & vbCrLf & docPath, vbInformation, "Erfolg"
    
    Exit Sub
    
ErrorHandler:
    LogFehler Err.Number, Err.Description, "GenerateDokument"
    MsgBox "Fehler bei der Generierung: " & Err.Description, vbCritical, "Fehler"
End Sub

' =====================================================
' Daten aus Excel lesen
' =====================================================
Private Function ReadDataFromExcel(ws As Worksheet) As Collection
    Dim daten As New Collection
    Dim felder As Variant
    Dim i As Integer
    
    ' Spaltenüberschriften in Zeile 1
    felder = Array("VORNAME", "NACHNAME", "STRASSE", "PLZ", "ORT", _
                    "GEBURTSDATUM", "AKTENZEICHEN", "DATUM", "BEHOERDE", _
                    "ABTEILUNG", "BETREFF", "INHALT", "FRIST", "BETRAG", _
                    "KONTAKT", "ANREDE")
    
    ' Werte aus Zeile 2 lesen
    For i = LBound(felder) To UBound(felder)
        Dim wert As String
        wert = ws.Cells(2, i + 1).Value
        If wert = "" Then wert = "[NICHT ANGEGEBEN]"
        daten.Add wert, felder(i)
    Next i
    
    Set ReadDataFromExcel = daten
End Function

' =====================================================
' Word-Vorlage füllen
' =====================================================
Private Function FillWordTemplate(vorlagenPfad As String, daten As Collection, ws As Worksheet) As String
    On Error GoTo ErrorHandler
    
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim ausgabePfad As String
    
    ' Word erstellen
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = False
    
    ' Vorlage öffnen
    If Dir(vorlagenPfad) = "" Then
        Err.Raise vbObjectError + 1, , "Vorlage nicht gefunden: " & vorlagenPfad
    End If
    
    Set wordDoc = wordApp.Documents.Add(Template:=vorlagenPfad)
    
    ' Platzhalter ersetzen
    Dim feld As Variant
    For Each feld In daten
        Dim platzhalter As String
        platzhalter = "{{" & feld & "}}"
        
        With wordApp.Selection
            .HomeKey Unit:=6  ' wdStory
            With .Find
                .Text = platzhalter
                .Replacement.Text = daten(feld)
                .Wrap = 1  ' wdFindContinue
                .Execute Replace:=2  ' wdReplaceAll
            End With
        End With
    Next feld
    
    ' Ausgabe-Pfad generieren
    Dim nachname As String
    nachname = daten("NACHNAME")
    Dim aktenzeichen As String
    aktenzeichen = daten("AKTENZEICHEN")
    
    ausgabePfad = ThisWorkbook.Path & "\output\" & nachname & "_" & aktenzeichen & ".docx"
    
    ' Ordner erstellen falls nötig
    If Dir(ThisWorkbook.Path & "\output", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\output"
    End If
    
    ' Speichern
    wordDoc.SaveAs2 ausgabePfad, 16  ' wdFormatDocumentDefault
    wordDoc.Close False
    wordApp.Quit False
    
    FillWordTemplate = ausgabePfad
    Exit Function
    
ErrorHandler:
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit False
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

' =====================================================
' KI-Optimierung
' =====================================================
Private Sub OptimizeWithKI(docPath As String, daten As Collection)
    On Error Resume Next  ' KI ist optional
    
    Dim kiConnector As New KIConnector
    Dim inhalt As String
    inhalt = daten("INHALT")
    
    If inhalt = "[NICHT ANGEGEBEN]" Then Exit Sub
    
    Dim optimiert As String
    optimiert = kiConnector.CallKIAPI("bescheid_optimierung", inhalt)
    
    If optimiert <> "" Then
        ' In Word-Dokument einfügen
        Dim wordApp As Object
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False
        
        Dim wordDoc As Object
        Set wordDoc = wordApp.Documents.Open(docPath)
        
        With wordApp.Selection
            .HomeKey Unit:=6
            With .Find
                .Text = daten("INHALT")
                .Replacement.Text = optimiert
                .Wrap = 1
                .Execute Replace:=2
            End With
        End With
        
        wordDoc.Save
        wordDoc.Close False
        wordApp.Quit False
    End If
    
    On Error GoTo 0
End Sub

' =====================================================
' Serienbrief generieren
' =====================================================
Public Sub GenerateSerienbrief()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim letzteZeile As Long
    letzteZeile = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If letzteZeile < 2 Then
        MsgBox "Keine Datenzeilen gefunden.", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long
    Dim erfolg As Long
    Dim fehlgeschlagen As Long
    
    For i = 2 To letzteZeile
        On Error Resume Next
        GenerateDokumentForRow ws, i
        If Err.Number = 0 Then
            erfolg = erfolg + 1
        Else
            fehlgeschlagen = fehlgeschlagen + 1
            LogFehler Err.Number, Err.Description, "Serienbrief Zeile " & i
        End If
        On Error GoTo ErrorHandler
    Next i
    
    MsgBox "Serienbrief abgeschlossen:" & vbCrLf & _
           erfolg & " erfolgreich, " & fehlgeschlagen & " fehlgeschlagen", _
           vbInformation, "Ergebnis"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Fehler: " & Err.Description, vbCritical
End Sub

Private Sub GenerateDokumentForRow(ws As Worksheet, zeile As Long)
    Dim daten As New Collection
    Dim felder As Variant
    Dim i As Integer
    
    felder = Array("VORNAME", "NACHNAME", "STRASSE", "PLZ", "ORT", _
                    "GEBURTSDATUM", "AKTENZEICHEN", "DATUM", "BEHOERDE", _
                    "ABTEILUNG", "BETREFF", "INHALT", "FRIST", "BETRAG", _
                    "KONTAKT", "ANREDE")
    
    For i = LBound(felder) To UBound(felder)
        Dim wert As String
        wert = ws.Cells(zeile, i + 1).Value
        If wert = "" Then wert = "[NICHT ANGEGEBEN]"
        daten.Add wert, felder(i)
    Next i
    
    Dim vorlagenPfad As String
    vorlagenPfad = ThisWorkbook.Path & "\" & VORLAGEN_ORDNER & "Bescheid-Vorlage.dotx"
    
    FillWordTemplate vorlagenPfad, daten, ws
End Sub

' =====================================================
' Protokollierung
' =====================================================
Private Sub LogGenerierung(daten As Collection, docPath As String, kiAktiv As Boolean)
    On Error Resume Next
    
    Dim fn As Integer
    fn = FreeFile
    
    Dim logPfad As String
    logPfad = ThisWorkbook.Path & "\" & LOG_DATEI
    
    Open logPfad For Append As #fn
    Print #fn, Now & ";" & daten("AKTENZEICHEN") & ";" & _
              daten("NACHNAME") & ";" & docPath & ";" & _
              IIf(kiAktiv, "JA", "NEIN")
    Close #fn
    
    On Error GoTo 0
End Sub

Private Sub LogFehler(fehlerNr As Long, beschreibung As String, kontext As String)
    On Error Resume Next
    
    Dim fn As Integer
    fn = FreeFile
    
    Dim logPfad As String
    logPfad = ThisWorkbook.Path & "\" & LOG_DATEI
    
    Open logPfad For Append As #fn
    Print #fn, "FEHLER;" & Now & ";" & fehlerNr & ";" & _
              beschreibung & ";" & kontext
    Close #fn
    
    On Error GoTo 0
End Sub
