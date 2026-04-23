' =====================================================
' KIConnector.bas
' VBA-Modul für KI-Anbindung über lokalen Proxy
' =====================================================
Option Explicit

Private Const DEFAULT_PROXY_URL As String = "http://localhost:5000"
Private Const REQUEST_TIMEOUT As Integer = 30  ' Sekunden

' =====================================================
' Hauptfunktion: KI-API aufrufen
' =====================================================
Public Function CallKIAPI(promptType As String, inputText As String) As String
    On Error GoTo ErrorHandler
    
    ' Prüfe Eingabe
    If Trim(inputText) = "" Then
        CallKIAPI = ""
        Exit Function
    End If
    
    ' Proxy-URL aus Registry oder Standard
    Dim proxyUrl As String
    proxyUrl = GetProxyURL()
    
    ' HTTP-Anfrage erstellen
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP.6.0")
    
    If xmlHttp Is Nothing Then
        Set xmlHttp = CreateObject("MSXML2.XMLHTTP.3.0")
    End If
    
    If xmlHttp Is Nothing Then
        Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
    End If
    
    If xmlHttp Is Nothing Then
        Err.Raise vbObjectError + 1, , "Kein XMLHTTP-Objekt verfügbar"
    End If
    
    ' JSON-Payload erstellen
    Dim payload As String
    payload = BuildJsonPayload(promptType, inputText)
    
    ' Anfrage senden
    xmlHttp.Open "POST", proxyUrl & "/api/optimize", False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.setTimeouts REQUEST_TIMEOUT * 1000, REQUEST_TIMEOUT * 1000, _
                        REQUEST_TIMEOUT * 1000, REQUEST_TIMEOUT * 1000
    
    xmlHttp.send payload
    
    ' Antwort prüfen
    If xmlHttp.Status <> 200 Then
        Err.Raise vbObjectError + 2, , "KI-Proxy Fehler " & xmlHttp.Status & ": " & xmlHttp.responseText
    End If
    
    ' JSON-Antwort parsen
    CallKIAPI = ParseJsonResponse(xmlHttp.responseText)
    
    Exit Function
    
ErrorHandler:
    ' KI ist optional - Fehler nicht kritisch
    Debug.Print "KI-Fehler: " & Err.Description
    CallKIAPI = ""
End Function

' =====================================================
' JSON-Payload erstellen (ohne externe Bibliothek)
' =====================================================
Private Function BuildJsonPayload(promptType As String, inputText As String) As String
    ' Einfacher JSON-String ohne Escaping-Lib
    Dim cleanText As String
    cleanText = EscapeJsonString(inputText)
    
    Dim promptText As String
    Select Case promptType
        Case "bescheid_optimierung"
            promptText = "Verbessere folgenden Behördentext. Formal, rechtssicher, verständlich behalten:"
        Case "berichts_zusammenfassung"
            promptText = "Fasse folgende Daten als strukturierten Behördenbericht zusammen:"
        Case Else
            promptText = "Verbessere folgenden Text:"
    End Select
    
    BuildJsonPayload = "{" & _
        """prompt"": """ & EscapeJsonString(promptText) & """," & _
        """text"": """ & cleanText & """" & _
        "}"
End Function

' =====================================================
' JSON-Antwort parsen
' =====================================================
Private Function ParseJsonResponse(jsonStr As String) As String
    ' Einfacher Parser: extrahiere "result" Feld
    Dim startIdx As Long
    Dim endIdx As Long
    
    startIdx = InStr(jsonStr, """result"":")
    If startIdx = 0 Then
        startIdx = InStr(jsonStr, """result"" :")
    End If
    
    If startIdx = 0 Then
        ParseJsonResponse = ""
        Exit Function
    End If
    
    ' Zum Wert springen
    startIdx = InStr(startIdx, jsonStr, """")
    If startIdx = 0 Then Exit Function
    startIdx = startIdx + 1
    
    ' Ende finden
    endIdx = InStr(startIdx, jsonStr, """")
    If endIdx = 0 Then Exit Function
    
    ParseJsonResponse = Mid(jsonStr, startIdx, endIdx - startIdx)
End Function

' =====================================================
' JSON-String escapen
' =====================================================
Private Function EscapeJsonString(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbTab, "\t")
    EscapeJsonString = result
End Function

' =====================================================
' Proxy-URL ermitteln
' =====================================================
Private Function GetProxyURL() As String
    On Error Resume Next
    ' Aus Registry lesen (optional)
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    Dim url As String
    url = ws.RegRead("HKCU\Software\VerwaltungsDokument\ProxyURL")
    If url <> "" Then
        GetProxyURL = url
        Exit Function
    End If
    On Error GoTo 0
    
    GetProxyURL = DEFAULT_PROXY_URL
End Function

' =====================================================
' KI-Verfügbarkeit prüfen
' =====================================================
Public Function CheckKIAvailable() As Boolean
    On Error GoTo NotAvailable
    
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP.6.0")
    
    xmlHttp.Open "GET", GetProxyURL() & "/health", False
    xmlHttp.setTimeouts 3000, 3000, 3000, 3000
    xmlHttp.send
    
    CheckKIAvailable = (xmlHttp.Status = 200)
    Exit Function
    
NotAvailable:
    CheckKIAvailable = False
End Function
