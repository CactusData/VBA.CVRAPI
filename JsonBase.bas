Attribute VB_Name = "JsonBase"
' JsonBase v1.2.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/CactusData/VBA.CVRAPI
'
' Supporting functions for retrieval of data from a Json service.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
Option Compare Text
Option Explicit

' Enum for HTTP methods.
Public Enum HttpVerb
    hvDelete = 8    ' Requests that a specified URI be deleted.
    hvGet = 1       ' Retrieves the information or entity that is identified by the URI of the request.
    hvHead = 16     ' Retrieves the message headers for the information or entity that is identified by the URI of the request.
    hvOptions = 64  ' Represents a request for information about the communication options available on the request/response chain identified by the Request-URI.
    hvPatch = 32    ' Requests that a set of changes described in the request entity be applied to the resource identified by the Request- URI.
    hvPost = 2      ' Posts a new entity as an addition to a URI.
    hvPut = 4       ' Replaces an entity that is identified by a URI.
End Enum

Public Sub AppendContentKeyValue( _
    ByRef Body As String, _
    ByVal Key As String, _
    ByVal Value As String, _
    Optional ByVal UseNull As Boolean)
    
    Const Delimiter As String = ","
    Const Separator As String = ":"
    
    Dim Pair        As String
    Dim Pairs()     As String
    
    If Key <> "" And (Value <> "" Or UseNull) Then
        If Value <> "" Then
            Pair = """" & Key & """" & Separator & """" & Value & """"
        ElseIf UseNull Then
            Pair = """" & Key & """" & Separator & "null"
        End If
        Pairs = Split(Body, Delimiter)
        ReDim Preserve Pairs(UBound(Pairs) + 1)
        Pairs(UBound(Pairs)) = Pair
        Body = Join(Pairs, Delimiter)
    End If

End Sub

Public Sub AppendSubPath( _
    ByRef Url As String, _
    ByVal SubPath As String)
    
    Const SeparatorFirst    As String = "?"
    Const SeparatorPath     As String = "/"
    
    Dim UrlParts()  As String
    
    If SubPath <> "" Then
        UrlParts = Split(Url, SeparatorFirst)
        If Right(UrlParts(0), 1) <> SeparatorPath Then
            SubPath = SeparatorPath & SubPath
        End If
        UrlParts(0) = UrlParts(0) & SubPath
        ' Return modified URL including the appended subpath.
        Url = Join(UrlParts, SeparatorFirst)
    End If
    
End Sub

' Break one-lined multiple elements into separate lines.
' For better read-out when debugging.
' Returns the modified Json by reference.
'
' Example input:
'   {"taxCodes": [{"id": "bd1e66f3","code": "U-00"},{"id": "453c1e3b","code": "U-01"}]}
' Example output:
'   {"taxCodes": [{"id": "bd1e66f3","code": "U-00"},
'   {"id": "453c1e3b","code": "U-01"}]}
'
' 2022-05-24. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub BreakJson(ByRef Json As String)

    Const BreakMark As String = "},{"
    Const BreakLine As String = "}," & vbNewLine & "{"
    
    Dim Text        As String
    
    ' Break one-lined multiple elements into separate lines.
    Text = Replace(Json, BreakMark, BreakLine)
    ' Return the sanitised json.
    Json = Text

End Sub

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
        QueryElement = Trim(Key) & Separator & ValueString
    End If
    
    BuildUrlQueryParameter = QueryElement
    
End Function

' Converts a name to general proper case leaving the company type abreviation intact.
'
' Note: Often company and city names are received in uppercase only.
'
' Example:
'   CACTUS DATA APS -> Cactus Data ApS
'   BERGEN -> Bergen
'
Public Function FormatCompany( _
    ByVal Company As String) _
    As String
    
    Dim CompanyTypes()      As Variant
    Dim ProperCompany       As String
    Dim Index               As Integer
    
    CompanyTypes() = Array("AmbA", "A.m.b.A", "ApS", "AS", "A/S", "I/S", "IVS", "K/S", "P/S")
    
    ProperCompany = Replace(StrConv(Replace(Company, "v/", "¤v/ "), vbProperCase), "¤v/ ", "v/")
    
    For Index = LBound(CompanyTypes) To UBound(CompanyTypes)
        If Left(ProperCompany, Len(CompanyTypes(Index)) + 1) = CompanyTypes(Index) & " " Then
            Mid(ProperCompany, 1) = CompanyTypes(Index)
        End If
        If Right(ProperCompany, Len(CompanyTypes(Index)) + 1) = " " & CompanyTypes(Index) Then
            Mid(ProperCompany, Len(ProperCompany) - Len(CompanyTypes(Index)) + 1) = CompanyTypes(Index)
        End If
    Next
    
    FormatCompany = ProperCompany
    
End Function

' Create the literal uppercased HTTP method from a passed HTTP verb.
'
' 2022-04-28. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function HttpMethod( _
    ByVal Method As HttpVerb) _
    As String
    
    Dim LiteralMethod   As String
    
    Select Case Method
        Case hvDelete
            LiteralMethod = "DELETE"
        Case hvGet
            LiteralMethod = "GET"
        Case hvHead
            LiteralMethod = "HEAD"
        Case hvOptions
            LiteralMethod = "OPTIONS"
        Case hvPatch
            LiteralMethod = "PATCH"
        Case hvPost
            LiteralMethod = "POST"
        Case hvPut
            LiteralMethod = "PUT"
    End Select
    
    HttpMethod = LiteralMethod

End Function

