Attribute VB_Name = "JsonScript"
' JsonScript v1.2.0
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/CactusData/VBA.CVRAPI
'
' Low-level wrapper functions to retrieve and encode/decode Json data by JavaScript.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' 2018-05-03:   Binding of Script Control changed to late binding for simplicity.
'               Added option for 64-bit script control (third-party)
'
' Requires:
'   32-bit VBA: Presence of "Microsoft Script Control 1.0"
'   64-bit VBA: Install of third-party script control "Tablacus Script Control 64"
'               https://tablacus.github.io/scriptcontrol_en.html
'
Option Compare Text
Option Explicit

' Script engine to run JavaScript (Microsoft JScript).
Private ScriptEngine        As Object

' Initialize the engine.
'
Public Sub InitiateScriptEngine()

    Dim Prompt  As String
    Dim Buttons As VbMsgBoxStyle
    Dim Title   As String

    On Error GoTo Err_InitiateScriptEngine
    
    If ScriptEngine Is Nothing Then
        Set ScriptEngine = CreateObject("ScriptControl")
    
        ScriptEngine.Language = "JScript"
        ScriptEngine.AddCode "function encode(plainString) {return encodeURIComponent(plainString);}"
        ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) {return jsonObj[propertyName];}"
        ScriptEngine.AddCode "function getKeys(jsonObj) {var keys = new Array(); for (var i in jsonObj) {keys.push(i);} return keys;}"
    End If
    
Exit_InitiateScriptEngine:
    Exit Sub

Err_InitiateScriptEngine:
    Prompt = "Error " & Err.Number & ":" & vbCrLf & Err.Description
    Buttons = vbCritical + vbOKOnly
    Title = "Script Control Objcet Error"
    MsgBox Prompt, Buttons, Title
    Resume Exit_InitiateScriptEngine
    
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

' Runs in 32-bit but fails in 64-bit.
''    For Each Key In KeysObject
''        Keys(Index) = Key
''        Index = Index + 1
''    Next
' However:

    ' KeysObject is just a comma separated string ...
    Keys = Split(KeysObject, ",")

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

