Attribute VB_Name = "CvrUtil"
Option Compare Text
Option Explicit

' Replacement for Microsoft Access' function Application.Nz().
' For use in Word/Excel to eliminate the need for a reference to:
'
'     Microsoft Office 15.0 Access database engine Object Library.
'
' 2015-12-10. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function Nz( _
    ByRef Value As Variant, _
    Optional ByRef ValueIfNull As Variant = "") _
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
