Attribute VB_Name = "Access"
' Access v1.0.2
' (c) 2018-05-01. Gustav Brock, Cactus Data ApS, CPH
' https://github.com/GustavBrock/VBA.Round
'
' Replacement for functions only found in Access.
'
' NOT needed, if a reference to Access has been established,
' for example for Access 2016:
'
'   Microsoft Access 16.0 Object Library
'
' If a reference to Access is established, this module can NOT
' be added and will raise a name conflict error, if an attempt is made.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)

Option Explicit

' Replacement for the function Application.Eval() of Access.
'
' 2018-04-12. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Eval(ByRef StringExpr As String) As Variant

    Eval = Application.Evaluate(StringExpr)
    
End Function

' Replacement for the function Application.Nz() of Access.
'
' Returns by default Empty if argument Value is Null and
' no value for argument ValueIfNull is passed.
'
' 2020-10-13. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Nz( _
    ByRef Value As Variant, _
    Optional ByRef ValueIfNull = Empty) _
    As Variant

    Dim ValueNz     As Variant

    If Not IsEmpty(Value) Then
        If IsNull(Value) Then
            ValueNz = ValueIfNull
        Else
            ValueNz = Value
        End If
    End If
        
    Nz = ValueNz
    
End Function
