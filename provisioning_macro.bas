Attribute VB_Name = "ProvisioningMacro"
Option Explicit

' ── API endpoints ─────────────────────────────────────────────────────────────
' Replace YOUR_API_ID and YOUR_REGION with your actual API Gateway values.
Private Const PROVISION_API_URL As String = "https://YOUR_API_ID.execute-api.YOUR_REGION.amazonaws.com/Prod/provision"
Private Const STATUS_API_URL    As String = "https://YOUR_API_ID.execute-api.YOUR_REGION.amazonaws.com/Prod/status"

' API Gateway API key — required when the stage is protected by a Usage Plan.
' Store this in a named range or a separate config sheet rather than hardcoding in production.
Private Const API_KEY As String = "YOUR_API_GATEWAY_API_KEY"

' ── Column layout of Sheet1 ───────────────────────────────────────────────────
' A  = Gear ID       B  = OS (linux / windows)    C  = Region
' D  = Technology    E  = Environment (DEV/QA/PROD)
' F  = HTTP Status (written back by macros)
' G  = Response Body (written back by macros)

' =============================================================================
' SubmitProvisioningRequest
' Select any cell in the row you want to provision, then run this macro.
' It POST-s the row data to the Validator Lambda via API Gateway.
' =============================================================================
Public Sub SubmitProvisioningRequest()
    Dim ws          As Worksheet
    Dim rowNum      As Long
    Dim gearId      As String
    Dim osName      As String
    Dim regionName  As String
    Dim technology  As String
    Dim envName     As String
    Dim payload     As String
    Dim http        As Object

    Set ws   = ThisWorkbook.Worksheets("Sheet1")
    rowNum   = ActiveCell.Row

    gearId     = Trim(CStr(ws.Cells(rowNum, 1).Value))
    osName     = LCase(Trim(CStr(ws.Cells(rowNum, 2).Value)))
    regionName = Trim(CStr(ws.Cells(rowNum, 3).Value))
    technology = LCase(Trim(CStr(ws.Cells(rowNum, 4).Value)))
    envName    = UCase(Trim(CStr(ws.Cells(rowNum, 5).Value)))

    If gearId = "" Then
        MsgBox "Gear ID (column A) is required.", vbExclamation
        Exit Sub
    End If

    payload = "{" & _
              """gear_id"":"     & Chr(34) & EscapeJson(gearId)     & Chr(34) & "," & _
              """os"":"          & Chr(34) & EscapeJson(osName)     & Chr(34) & "," & _
              """region"":"      & Chr(34) & EscapeJson(regionName) & Chr(34) & "," & _
              """environment"":" & Chr(34) & EscapeJson(envName)    & Chr(34) & "," & _
              """technology"":"  & Chr(34) & EscapeJson(technology) & Chr(34) & _
              "}"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", PROVISION_API_URL, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "x-api-key", API_KEY   ' Required by API Gateway Usage Plan
    http.Send payload

    ws.Cells(rowNum, 6).Value = http.Status
    ws.Cells(rowNum, 7).Value = http.ResponseText

    If http.Status = 202 Then
        MsgBox "Request accepted for Gear ID: " & gearId & vbCrLf & _
               "Poll status using RefreshGearStatus.", vbInformation
    Else
        MsgBox "Unexpected response (" & http.Status & "):" & vbCrLf & _
               http.ResponseText, vbExclamation
    End If
End Sub


' =============================================================================
' RefreshGearStatus
' Select any cell in the row, then run this macro to fetch the latest
' provisioning status from DynamoDB via the Status Lambda.
' =============================================================================
Public Sub RefreshGearStatus()
    Dim ws     As Worksheet
    Dim rowNum As Long
    Dim gearId As String
    Dim http   As Object

    Set ws   = ThisWorkbook.Worksheets("Sheet1")
    rowNum   = ActiveCell.Row
    gearId   = Trim(CStr(ws.Cells(rowNum, 1).Value))

    If gearId = "" Then
        MsgBox "Select a row that has a Gear ID in column A.", vbExclamation
        Exit Sub
    End If

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", STATUS_API_URL & "?gear_id=" & UrlEncode(gearId), False
    http.SetRequestHeader "x-api-key", API_KEY   ' Required by API Gateway Usage Plan
    http.Send

    ws.Cells(rowNum, 6).Value = http.Status
    ws.Cells(rowNum, 7).Value = http.ResponseText

    If http.Status = 200 Then
        ' Optionally parse and write status / instance_id back to dedicated columns here.
        MsgBox "Status for " & gearId & ":" & vbCrLf & http.ResponseText, vbInformation
    ElseIf http.Status = 404 Then
        MsgBox "Gear ID not found: " & gearId, vbExclamation
    Else
        MsgBox "Error (" & http.Status & "):" & vbCrLf & http.ResponseText, vbExclamation
    End If
End Sub


' =============================================================================
' Private helpers
' =============================================================================
Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\",  "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function

Private Function UrlEncode(ByVal s As String) As String
    ' Basic URL encoding for gear_id values (alphanumeric + safe chars assumed).
    ' If gear IDs can contain spaces or special chars, use a fuller encoding here.
    UrlEncode = Replace(s, " ", "%20")
End Function
