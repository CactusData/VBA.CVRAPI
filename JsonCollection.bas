Attribute VB_Name = "JsonCollection"
' Set of low-level functions to retrieve and decode data from a Json service
' and return these as a data collection.
'
Option Compare Text
Option Explicit

' Decode a Json response text and convert it to a collection of arrays.
'
Public Function CollectJson( _
    ByVal ResponseText As String) _
    As Collection

    Const CollectionName    As String = "root"
    
    Dim col                 As Collection
    Dim colRoot             As Collection
    Dim JsonObject          As Object

    Set col = New Collection
    Set JsonObject = DecodeJsonString(ResponseText)
    
    Set col = FillCollection(JsonObject, CollectionName)
    If VarType(col(1)(CollectionItem.Name)) <> vbObject Then
        ' Append the field collection to a root object.
        Set colRoot = New Collection
        colRoot.Add Array(CollectionName, col), CollectionName
        Set col = colRoot
    End If
    
    Set CollectJson = col
    
    ' Finished using the script engine.
    Call TerminateScriptEngine
    
End Function

' Collect members of a Json object recursively.
' Returns a collection of arrays of key/value pairs.
'
Private Function FillCollection( _
    ByRef JsonObject As Object, _
    Optional ByVal CollectionName As String) _
    As Collection
    
    Dim col         As Collection
    
    Dim Keys()      As String
    Dim Key         As String
    Dim KeyValue    As Variant
    Dim Index       As Long
    
    Set col = New Collection
    
    ' Collect array of key and value of members of JsonObject recursively.
    ' Note: CollectionName is not implemented. Could be used for a tree build.
    Keys = GetKeys(JsonObject)

    For Index = LBound(Keys) To UBound(Keys)
        Key = Keys(Index)
        KeyValue = GetProperty(JsonObject, Key)
        If InStr(KeyValue, "[object Object]") > 0 Then
            ' Subcollection.
            col.Add Array(Key, FillCollection(GetObjectProperty(JsonObject, Key), Key)), Key
        Else
            ' Field value.
            col.Add Array(Key, KeyValue), Key
        End If
    Next
    
    Set FillCollection = col
    
End Function
