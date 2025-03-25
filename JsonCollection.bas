Attribute VB_Name = "JsonCollection"
' Set of low-level functions to retrieve and decode data from a Json service
' and return these as a data collection.
'
Option Compare Text
Option Explicit

Public Enum CollectionItem
    Name = 0
    Data = 1
End Enum

' Decode a Json response text and convert it to a collection of arrays.
'
Public Function CollectJson( _
    ByVal ResponseText As String) _
    As Collection

    Const CollectionName    As String = "root"
    
    Dim Col                 As Collection
    Dim colRoot             As Collection
    Dim JsonObject          As Object

    Set Col = New Collection
    Set JsonObject = DecodeJsonString(ResponseText)
    
    If Not JsonObject Is Nothing Then
        Set Col = FillCollection(JsonObject)
        If Not Col Is Nothing Then
            If VarType(Col(1)(CollectionItem.Name)) <> vbObject Then
                ' Append the field collection to a root object.
                Set colRoot = New Collection
                colRoot.Add Array(CollectionName, Col), CollectionName
                Set Col = colRoot
            End If
        End If
    End If
    
    Set CollectJson = Col
    
    ' Finished using the script engine.
    Call TerminateScriptEngine
    
End Function

' Collect members of a Json object recursively.
' Returns a collection of arrays of key/value pairs.
'
Private Function FillCollection( _
    ByRef JsonObject As Object) _
    As Collection
    
    Dim Col         As Collection
    
    Dim Keys()      As String
    Dim Key         As String
    Dim KeyValue    As Variant
    Dim Index       As Long
        
    If Not JsonObject Is Nothing Then
        ' Collect array of key and value of members of JsonObject recursively.
        ' Note: CollectionName is not implemented. Could be used for a tree build.
        Keys = GetKeys(JsonObject)
    
        
        If LBound(Keys) <= UBound(Keys) Then
            Set Col = New Collection
        Else
            ' Empty array.
            Set Col = Nothing
        End If
    
        For Index = LBound(Keys) To UBound(Keys)
            Key = Keys(Index)
            KeyValue = GetProperty(JsonObject, Key)
            If InStr(KeyValue, "[object Object]") > 0 Then
                ' Subcollection.
                Col.Add Array(Key, FillCollection(GetObjectProperty(JsonObject, Key))), Key
            Else
                ' Field value.
                Col.Add Array(Key, KeyValue), Key
            End If
        Next
    End If
    
    Set FillCollection = Col
    
End Function

