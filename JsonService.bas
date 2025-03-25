Attribute VB_Name = "JsonService"
' JsonService v1.2.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/CactusData/VBA.CVRAPI
'
' Set of functions to retrieve data from a Json service
' and return these as a response text or data collection.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Requires: A reference to "Microsoft XML, v6.0".
'
Option Compare Text
Option Explicit

' Retrieve a Json response from a service URL.
'
' Returns a data collection.
'
Public Function RetrieveDataCollection( _
    ByVal ServiceUrl As String, _
    Optional ByVal UserAgent As String) _
    As Collection
    
    Dim DataCollection      As Collection
    
    Dim ResponseText        As String
    
    If ServiceUrl <> "" Then
        If RetrieveDataResponse(ServiceUrl, ResponseText, UserAgent) = True Then
            Set DataCollection = CollectJson(ResponseText)
            ' Debug.Print ServiceUrl
            ' Debug.Print ResponseText
        End If
    End If
    
    Set RetrieveDataCollection = DataCollection

End Function

' Retrieve a Json response from a service URL.
' Retrieved data is returned in parameter ResponseText.
'
' Returns True if success.
'
Public Function RetrieveDataResponse( _
    ByVal ServiceUrl As String, _
    ByRef ResponseText As String, _
    Optional ByVal UserAgent As String, _
    Optional ByVal UserName As String, _
    Optional ByVal Password As String) _
    As Boolean

    ' ServiceUrl is expected to have URL encoded parameters.
    
    ' User defined constants.
    Const DefaultUserAgent  As String = "Min organisation - Mit projekt"
    ' Fixed constants.
    Const Async             As Boolean = False
    Const StatusOk          As Integer = 200
    Const StatusNotFound    As Integer = 404
    
    ' Engine to communicate with the Json service.
    Dim XmlHttp             As XMLHTTP60
    
    Dim Result              As Boolean
  
    On Error GoTo Err_RetrieveDataResponse
    
    Set XmlHttp = New XMLHTTP60
    
    If UserAgent = "" Then
        ' Set default string for User-Agent.
        UserAgent = DefaultUserAgent
    End If
    
    XmlHttp.Open "GET", ServiceUrl, Async, UserName, Password
    XmlHttp.setRequestHeader "User-Agent", UserAgent

    XmlHttp.Send

    ResponseText = XmlHttp.ResponseText
    Select Case XmlHttp.Status
        Case StatusOk
            Result = True
        Case StatusNotFound
            ' Special case for CVRAPI which returns 404 and a valid Json error message:
            ' {"error":"NOT_FOUND","t":0,"version":6}
            If _
                InStr(1, ResponseText, "{", vbBinaryCompare) = 1 And _
                InStr(1, ResponseText, "error", vbBinaryCompare) = 3 And _
                InStr(1, ResponseText, "}", vbBinaryCompare) = 39 Then
                ' Json error message received.
                Result = True
            End If
    End Select
    If Result = False Then
        If ResponseText = "" Then
            ResponseText = XmlHttp.statusText
        End If
        ResponseText = CStr(XmlHttp.Status) & ":" & vbCrLf & ResponseText
    End If
    
    RetrieveDataResponse = Result

Exit_RetrieveDataResponse:
    Set XmlHttp = Nothing
    Exit Function

Err_RetrieveDataResponse:
    MsgBox "Error" & Str(Err.Number) & ": " & Err.Description, vbCritical + vbOKOnly, "Web Service Error"
    Resume Exit_RetrieveDataResponse

End Function

