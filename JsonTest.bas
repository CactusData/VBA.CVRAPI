Attribute VB_Name = "JsonTest"
Option Compare Database
Option Explicit
'
' Functions for simple testing and listing of retrieved
' data from a Json service.

' Call a Json service and return result as a collection and a messagebox.
'
Public Sub TestJsonService()

    Dim DataCollection      As Collection
    
    Dim ServiceUrl          As String
    Dim ResponseText        As String
    Dim UserAgent           As String
    
    Const Username          As String = "demo"
    Const App_id            As String = "b492b663ae3e458d9f0b042e8edb8c63"
    
    ' Register at http://www.geonames.org/login
    ServiceUrl = "http://api.geonames.org/citiesJSON?north=44.1&south=-9.9&east=-22.4&west=55.2&lang=de&username=" & Username
    
    ' Register at https://openexchangerates.org/signup/free
    'ServiceUrl = "http://openexchangerates.org/api/latest.json?app_id=" & App_id
    
    'ServiceUrl = "http://cvrapi.dk/api?name=lagkagehuset&country=dk&format=json"
    UserAgent = "Example Org. - TestApp"
    
    If RetrieveDataResponse(ServiceUrl, ResponseText, UserAgent) = True Then
        Set DataCollection = CollectJson(ResponseText)
        MsgBox "Retrieved" & Str(DataCollection.Count) & " root member(s)", vbInformation + vbOKOnly, "Web Service Success"
    ElseIf ResponseText <> "" Then
        MsgBox ResponseText, vbCritical + vbOKOnly, "Web Service Error"
    End If
    
    Call ListFieldNames(DataCollection)
    
    Set DataCollection = Nothing
    
End Sub

' Analyze a manually entered Json string.
'
Public Sub TestJsonResponseText( _
    ByVal ResponseText As String)

    Dim DataCollection      As Collection
    ResponseText = InputBox("Json")
    If ResponseText <> "" Then
        Set DataCollection = CollectJson(ResponseText)
        MsgBox "Retrieved" & Str(DataCollection.Count) & " root member(s)", vbInformation + vbOKOnly, "Web Service Success"
    End If
    
    Call ListFieldNames(DataCollection)
    
    Set DataCollection = Nothing
    
End Sub

' List field names of a collection of arrays.
'
Public Sub ListFieldNames( _
    ByVal DataCollection As Collection, _
    Optional Indent As String)

    On Error GoTo Err_ListFieldNames
    
    Dim Index               As Long
    Dim MemberName          As String
    
    For Index = 1 To DataCollection.Count
        MemberName = Space(16)
        LSet MemberName = DataCollection(Index)(CollectionItem.Name)
        Debug.Print Indent & MemberName, ;
        If VarType(DataCollection(Index)(CollectionItem.Data)) = vbObject Then
            Debug.Print
            Call ListFieldNames(DataCollection(Index)(CollectionItem.Data), Indent & vbTab)
        Else
            Debug.Print Trim(DataCollection(Index)(CollectionItem.Data))
        End If
    Next
    
Exit_ListFieldNames:
    Exit Sub
    
Err_ListFieldNames:
    Debug.Print "Error" & Str(Err.Number) & ": " & Err.Description
    Resume Exit_ListFieldNames
    
End Sub
