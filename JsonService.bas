Attribute VB_Name = "JsonService"
' JsonService v1.0.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/CactusData/VBA.CVRAPI
'
' Set of functions to retrieve and decode data from a Json service
' and return these as a response text or data collection.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Requires: A reference to "Microsoft XML, v6.0".
'
Option Compare Text
Option Explicit

' Build a URL string from its components to call a service.
' Parameter Query must be URL encoded.
'
' Returns: URL string.
'
Public Function BuildServiceUrl( _
    Optional ByVal Scheme As String = "http", _
    Optional ByVal Host As String = "localhost", _
    Optional ByVal Path As String, _
    Optional ByVal Query As String) _
    As String

    Dim ServiceUrl          As String

    ' Verify scheme.
    If Scheme = "" Then
        Scheme = "http"
    End If
    ' Append scheme separator.
    If Right(Scheme, 3) <> "://" Then
        Scheme = Scheme & "://"
    End If
    
    ' Verify host.
    If Host = "" Then
        Host = "localhost"
    End If
    ' Append a trailing slash.
    If Right(Host, 1) <> "/" Then
        Host = Host & "/"
    End If
    
    ' Verify path.
    If Path <> "" Then
        ' Remove a leading slash.
        If Left(Path, 1) = "/" Then
            Path = Mid(Path, 2)
        End If
        ' Remove a trailing slash.
        If Right(Path, 1) = "/" Then
            Path = Mid(Path, 1, Len(Path) - 1)
        End If
        ' Remove an empty path.
        If Replace(Path, "/", "") = "" Then
            Path = ""
        End If
    End If
    
    ' Verify query.
    If Left(Query, 1) <> "?" Then
        Query = ""
    End If
    
    ServiceUrl = Scheme & Host & Path & Query
    
    BuildServiceUrl = ServiceUrl
    
End Function

' Build the query element of a URL string from a
' parameter array of key/value pairs.
'
' Returns: String of encoded query elements
'
Public Function BuildUrlQuery( _
    ParamArray QueryElements() As Variant) _
    As String
    
    ' Key/Value pairs of QueryElements must be URL encoded.
    
    Const SeparatorFirst    As String = "?"
    Const SeparatorNext     As String = "&"
    
    Dim QueryString         As String
    
    If UBound(QueryElements) > -1 Then
        QueryString = SeparatorFirst & Join(QueryElements, SeparatorNext)
    End If
    
    BuildUrlQuery = QueryString

End Function

' Build a URL encoded query element from its key/value pairs.
'
' Returns a key/value string: key=value.
'
Public Function BuildUrlQueryParameter( _
    ByVal Key As String, _
    ByVal Value As Variant) _
    As String

    Const Separator         As String = "="
    
    Dim QueryElement        As String
    Dim ValueString         As String
    
    ' Trim and URL encode the key/value pair.
    If Trim(Key) <> "" And IsEmpty(Value) = False Then
        ValueString = Trim(CStr(Nz(Value)))
        If ValueString = "" Then
            ValueString = "''"
        End If
        QueryElement = EncodeUrl(Trim(Key)) & Separator & EncodeUrl(ValueString)
    End If
    
    BuildUrlQueryParameter = QueryElement
    
End Function

' Retrieve a Json response from a service URL.
' Retrieved data is returned in parameter ResponseText.
'
' Returns True if success.
'
Public Function RetrieveDataResponse( _
    ByVal ServiceUrl As String, _
    ByRef ResponseText As String, _
    Optional ByVal UserAgent As String) _
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
    
    XmlHttp.Open "GET", ServiceUrl, Async
    XmlHttp.setRequestHeader "User-Agent", UserAgent

    XmlHttp.send

    ResponseText = XmlHttp.ResponseText
    Select Case XmlHttp.status
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
        ResponseText = CStr(XmlHttp.status) & ": " & XmlHttp.statusText
    End If
    
    RetrieveDataResponse = Result

Exit_RetrieveDataResponse:
    Set XmlHttp = Nothing
    Exit Function

Err_RetrieveDataResponse:
    MsgBox "Error" & Str(Err.Number) & ": " & Err.Description, vbCritical + vbOKOnly, "Web Service Error"
    Resume Exit_RetrieveDataResponse

End Function

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
    Dim Result              As Boolean
    
    If ServiceUrl <> "" Then
        If RetrieveDataResponse(ServiceUrl, ResponseText, UserAgent) = True Then
            Set DataCollection = CollectJson(ResponseText)
            ' Debug.Print ServiceUrl
            ' Debug.Print ResponseText
        End If
    End If
    
    Set RetrieveDataCollection = DataCollection

End Function
