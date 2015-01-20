Attribute VB_Name = "CvrDemo"
Option Compare Database
Option Explicit
'
' Example functions for using CVRAPI at application level.

' Basic second-level example function to retrieve company info from a vat number.
'
' Vat is searched.
' If found, True is returned, Vat is returned cleaned, and other parameters are filled.
' If not found, False is returned. Other parameters are left untouched.
'
' Example usage: See GetDkCompanyInfo().
'
Public Function RetrieveCvrAddress( _
    ByVal Country As CvrCountrySelect, _
    ByRef Vat As String, _
    ByRef Company As String, _
    ByRef Address As String, _
    ByRef PostalCode As String, _
    ByRef City As String) _
    As Boolean

    Dim DataCollection      As Collection
    Dim Result              As Boolean
    Dim FullResult          As CvrVat
    
    Set DataCollection = CvrLookup(Result, VatNo, Vat, Country)
    
    If Result = True Then
        ' Success.
        ' Purify data.
        FullResult = FillType(DataCollection)
        ' Return cleaned VAT number.
        Vat = FullResult.Vat
        ' Return info.
        Company = FullResult.Name
        Address = FullResult.Address
        PostalCode = FullResult.ZipCode
        City = FullResult.City
    End If
    
    Set DataCollection = Nothing
    
    RetrieveCvrAddress = Result

End Function

' Basic top-level example function to retrieve company info from a vat number.
' If found, info will be printed.
' If not found, nothing will be printed.
'
' Example:
'   Call GetDkCompanyInfo("20-21-30-94")
' will print:
'   VAT:          20213094
'   Company:      Lagkagehuset A/S
'   Street:       Amerikavej 21
'   City:         1756 København V
'
Public Sub GetDkCompanyInfo( _
    ByVal Vat As String)

    Const Country           As Long = CvrCountrySelect.Denmark
    
    Dim Company             As String
    Dim Address             As String
    Dim PostalCode          As String
    Dim City                As String
    
    If RetrieveCvrAddress(Country, Vat, Company, Address, PostalCode, City) Then
        Debug.Print "VAT:", Vat
        Debug.Print "Company:", Company
        Debug.Print "Street:", Address
        Debug.Print "City:", PostalCode & " " & City
    End If
    
End Sub

' Basic top-level example function to verify the existence of a vat number.
' Returns True if found.
'
' Examples:
'   IsCvr("20-21-30-94")
' returns True.
'   IsCvr("20-21-30-94", Norway)
' returns False.
'
Public Function IsCvr( _
    ByVal Vat As String, _
    Optional ByVal Country As CvrCountrySelect = CvrCountrySelect.Denmark) _
    As Boolean

    IsCvr = RetrieveCvrAddress(Country, Vat, "", "", "", "")
    
End Function
