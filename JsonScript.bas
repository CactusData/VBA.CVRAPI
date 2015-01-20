Attribute VB_Name = "JsonScript"
' JsonService v1.0.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/CactusData/VBA.CVRAPI
'
' Low-level wrapper functions to retrieve and encode/decode Json data by JavaScript.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' Requires: A reference to "Microsoft Script Control 1.0".
'
Option Compare Database
Option Explicit

' Script engine to run JavaScript (Microsoft JScript).
Private ScriptEngine        As ScriptControl

' Initialize the engine.
'
Public Sub InitiateScriptEngine()

    If ScriptEngine Is Nothing Then
        Set ScriptEngine = New ScriptControl
    
        ScriptEngine.Language = "JScript"
        ScriptEngine.AddCode "function encode(plainString) {return encodeURIComponent(plainString);}"
        ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) {return jsonObj[propertyName];}"
        ScriptEngine.AddCode "function getKeys(jsonObj) {var keys = new Array(); for (var i in jsonObj) {keys.push(i);} return keys;}"
    End If
    
End Sub

' Terminate the engine.
'
Public Sub TerminateScriptEngine()

    Set ScriptEngine = Nothing
    
End Sub

' Get the keys of a Json object.
'
Public Function GetKeys( _
    ByVal JsonObject As Object) _
    As String()

    Dim KeysObject  As Object
    
    Dim Keys()      As String
    Dim Length      As Integer
    Dim Index       As Integer
    Dim Key         As Variant

    Set KeysObject = ScriptEngine.Run("getKeys", JsonObject)

    Length = GetProperty(KeysObject, "length")
    ReDim Keys(Length - 1)

    For Each Key In KeysObject
        Keys(Index) = Key
        Index = Index + 1
    Next

    GetKeys = Keys

End Function

' Get a property by name.
'
Public Function GetProperty(ByVal JsonObject As Object, ByVal propertyName As String) As Variant

    GetProperty = ScriptEngine.Run("getProperty", JsonObject, propertyName)

End Function

' Get a property as an object by name.
'
Public Function GetObjectProperty(ByVal JsonObject As Object, ByVal propertyName As String) As Object

    Set GetObjectProperty = ScriptEngine.Run("getProperty", JsonObject, propertyName)

End Function

' URL Encode a string.
'
Public Function EncodeUrl( _
    ByVal PlainString As String) _
    As String
    
    Dim EncodedString       As String
    
    Call InitiateScriptEngine
    
    EncodedString = ScriptEngine.Run("encode", PlainString)
    
    EncodeUrl = EncodedString

End Function

' URL decode a Json string.
'
Public Function DecodeJsonString(ByVal JSonString As String) As Object

    Call InitiateScriptEngine
    
    Set DecodeJsonString = ScriptEngine.Eval("(" + JSonString + ")")

End Function
