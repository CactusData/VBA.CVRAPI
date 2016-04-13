Attribute VB_Name = "CvrDemo"
Option Compare Text
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
    ByRef VAT As String, _
    ByRef Company As String, _
    ByRef Address As String, _
    ByRef PostalCode As String, _
    ByRef City As String) _
    As Boolean

    Dim DataCollection      As Collection
    Dim Result              As Boolean
    Dim FullResult          As CvrVat
    
    Set DataCollection = CvrLookup(Result, VatNo, VAT, Country)
    
    If Result = True Then
        ' Success.
        ' Purify data.
        FullResult = FillType(DataCollection)
        ' Return cleaned VAT number.
        VAT = FullResult.VAT
        ' Return info.
        Company = FullResult.Name
        Address = FullResult.Address
        PostalCode = FullResult.ZipCode
        City = FullResult.City
    End If
    
    Set DataCollection = Nothing
    
    RetrieveCvrAddress = Result

End Function

' Basic second-level example function to retrieve vat number and address from company name.
'
' Company is searched.
' If found, True is returned, Company is returned cleaned, and other parameters are filled.
' If not found, False is returned. Other parameters are left untouched.
'
' LIMITATION:
'   Company name must be unique as CVRAPI currently returns a first match only.
'
' Example usage: See GetDkVat().
'
Public Function RetrieveCvrVat( _
    ByVal Country As CvrCountrySelect, _
    ByRef VAT As String, _
    ByRef Company As String, _
    ByRef Address As String, _
    ByRef PostalCode As String, _
    ByRef City As String) _
    As Boolean

    Dim DataCollection      As Collection
    Dim Result              As Boolean
    Dim FullResult          As CvrVat
    
    Set DataCollection = CvrLookup(Result, CompanyName, Company, Country)
    
    If Result = True Then
        ' Success.
        ' Purify data.
        FullResult = FillType(DataCollection)
        ' Return cleaned VAT number.
        VAT = FullResult.VAT
        ' Return info.
        Company = FullResult.Name
        Address = FullResult.Address
        PostalCode = FullResult.ZipCode
        City = FullResult.City
    End If
    
    Set DataCollection = Nothing
    
    RetrieveCvrVat = Result

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
    ByVal VAT As String)

    Const Country           As Long = CvrCountrySelect.Denmark
    
    Dim Company             As String
    Dim Address             As String
    Dim PostalCode          As String
    Dim City                As String
    
    If RetrieveCvrAddress(Country, VAT, Company, Address, PostalCode, City) Then
        Debug.Print "VAT:", VAT
        Debug.Print "Company:", Company
        Debug.Print "Street:", Address
        Debug.Print "City:", PostalCode & " " & City
    End If
    
End Sub

' Basic top-level example function to retrieve vat number and address from company name.
' If found, info will be printed.
' If not found, nothing will be printed.
'
' LIMITATION:
'   Company name must be unique as CVRAPI currently returns a first match only.
'
' Example:
'   Call GetDkVat("nydata")
' will print:
'   VAT:          33402996
'   Company:      Nydata.dk v/Per Stenholt Andersen
'   Street:       Roskildevej 278A, st. tv.
'   City:         2610 Rødovre

Public Sub GetDkVat( _
    ByVal Company As String)

    Const Country           As Long = CvrCountrySelect.Denmark
    
    Dim VAT                 As String
    Dim Address             As String
    Dim PostalCode          As String
    Dim City                As String
    
    If RetrieveCvrVat(Country, VAT, Company, Address, PostalCode, City) Then
        Debug.Print "VAT:", VAT
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
    ByVal VAT As String, _
    Optional ByVal Country As CvrCountrySelect = CvrCountrySelect.Denmark) _
    As Boolean

    IsCvr = RetrieveCvrAddress(Country, VAT, "", "", "", "")
    
End Function

' Basic top-level example function to retrieve a full set of data
' typically from a search by VAT number or company name.
'
' Returns the UDT (User Defined Type) CvrVat.
'
' Example:
'   Dim CvrVatResult As CvrVat
'   Dim Result       As Boolean
'   CvrVatResult = GetCvrData(CvrSearchKey.CompanyName, "TheUniqueCompanyName", Result)
' returns full info in CvrVatResult.
'
Public Function GetCvrData( _
    ByVal SearchKey As CvrSearchKey, _
    ByVal SearchValue As String, _
    ByRef Result As Boolean, _
    Optional ByVal CountryValue As CvrCountrySelect = CvrCountrySelect.Denmark) _
    As CvrVat

    Dim DataCollection      As Collection
    Dim FullResult          As CvrVat
    
    Set DataCollection = CvrLookup(Result, SearchKey, SearchValue, CountryValue)
    
    If Result = True Then
        ' Success.
        ' Purify data.
        FullResult = FillType(DataCollection)
    End If
    
    Set DataCollection = Nothing
    
    GetCvrData = FullResult

End Function
