Attribute VB_Name = "CvrService"
' CvrService v1.1.2
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/CactusData/VBA.CVRAPI
'
' Set of base functions to retrieve and decode data from CVRAPI.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
Option Compare Text
Option Explicit


' Enums.
'
Public Enum CvrSearchKey
    GeneralSearch
    VatNo
    CompanyName
    Productionunit
    PhoneNumber
End Enum

Public Enum CvrCountryKey
    Country
End Enum

Public Enum CvrFormatKey
    DataFormat
End Enum
    
Public Enum CvrVersionKey
    Version
End Enum

Public Enum CvrCountrySelect
    Denmark
    Norway
End Enum

Public Enum CvrFormatSelect
    Json
    XML
End Enum

Public Enum CvrVersionSelect
    Version0
    Version1
    Version2
    Version3
    Version4
    Version5
    Version6
    Newest
End Enum

Public Enum CvrPnoCodeLength
    Dk = 10
    No = 9
End Enum

Public Enum CvrVatCodeLength
    Dk = 8
    No = 9
End Enum

Public Enum CvrPnoFirstDigitMin
    Dk = 1
    No = 8
End Enum

Public Enum CvrVatFirstDigitMin
    Dk = 1
    No = 8
End Enum

' User defined data types.
'
Public Type CvrOwner
    Name            As String
End Type

Public Type CvrProductionunit
    Pno             As Double
    Main            As Boolean
    Name            As String
    Address         As String
    ZipCode         As String
    City            As String
    Protected       As Boolean
    Phone           As String
    Email           As String
    Fax             As String
    Startdate       As Variant
    Enddate         As Variant
    Employees       As String
    Addressco       As String
    Industrycode    As Long
    Industrydesc    As String
End Type

Public Type CvrVat
    VAT             As Long
    Name            As String
    Address         As String
    ZipCode         As String
    City            As String
    Protected       As Boolean
    Phone           As String
    Email           As String
    Fax             As String
    Startdate       As Variant
    Enddate         As Variant
    Employees       As String
    Addressco       As String
    Industrycode    As String
    Industrydesc    As String
    Companycode     As Integer
    Companydesc     As String
    Creditstartdate As Variant
    Creditstatus    As Integer
    Creditbankrupt  As Boolean
    Owner           As CvrOwner
    Productionunit  As CvrProductionunit
    T               As Integer
    Version         As Integer
End Type

Public Type CvrError
    Error           As String
    T               As Integer
    Version         As Integer
End Type

' Magic numbers.
'
' Highest production unit number (P-nummer) with custom
' Modulus 11 check digit calculation.
Private Const PNumberOldMax     As Double = 1006959421

' Main function for searching the CVRAPI.
' Result returns success or error as True or False.
'
' Returns a Collection that can be transformed to user defined types
' by functions FillType and FillError.
'
' Note:
'   FormatValue is only implemented for CvrFormatSelect.Json.
'   VersionValue is only implemented for CvrVersionSelect.Version6 and .Newest.
'
' Note:
'   Application specific constants must be adjusted prior to production usage.
'
Public Function CvrLookup( _
    ByRef Result As Boolean, _
    ByVal SearchKey As CvrSearchKey, _
    ByVal SearchValue As String, _
    Optional CountryValue As CvrCountrySelect = CvrCountrySelect.Denmark, _
    Optional FormatValue As CvrFormatSelect = CvrFormatSelect.Json, _
    Optional VersionValue As CvrVersionSelect = CvrVersionSelect.Newest) _
    As Collection

    ' Application specific constants.
    Const UserAgentOrg      As String = "Cactus Data ApS"  ' Your organisation.
    Const UserAgentApp      As String = "Accesstest"       ' Your app name.
    
    ' API specific constants.
    Const Host              As String = "cvrapi.dk"         ' Do not change.
    Const Path              As String = "api"               ' Do not change.
    Const UserAgent         As String = UserAgentOrg & " - " & UserAgentApp
    
    ' Constants for this procedure.
    Const CountryKey        As Long = CvrCountryKey.Country
    Const FormatKey         As Long = CvrFormatKey.DataFormat
    Const VersionKey        As Long = CvrVersionKey.Version
    ' First (only) item in error message collection.
    Const RootItem          As Integer = 1
    ' Count of elements in error message.
    Const ErrorItems        As Integer = 3
    ' Default error code.
    Const DefaultErrorCode  As String = "NOT_FOUND"
    
    Dim DataCollection      As Collection
    
    Dim SearchParamKey      As String
    Dim SearchParamVal      As String
    Dim CountryParamKey     As String
    Dim CountryParamVal     As String
    Dim FormatParamKey      As String
    Dim FormatParamVal      As String
    Dim VersionParamKey     As String
    Dim VersionParamVal     As String
    
    Dim ServiceUrl          As String
    Dim Query               As String
    Dim ResponseText        As String
    Dim ErrorCode           As String
    
    If Trim(SearchValue) <> "" Then
        
        If ValidateSearch(CountryValue, SearchKey, SearchValue, ErrorCode) = True Then
            SearchParamKey = CvrSearchKeyLabel(SearchKey)
            SearchParamVal = SearchValue
        
            CountryParamKey = CvrCountryKeyLabel(CountryKey)
            CountryParamVal = CvrCountryValue(CountryValue)
            
            FormatParamKey = CvrFormatKeyLabel(FormatKey)
            FormatParamVal = CvrFormatValue(FormatValue)
            
            VersionParamKey = CvrVersionKeyLabel(VersionKey)
            VersionParamVal = CvrVersionValue(VersionValue)
        
            Query = BuildUrlQuery( _
                BuildUrlQueryParameter(SearchParamKey, SearchParamVal), _
                BuildUrlQueryParameter(CountryParamKey, CountryParamVal), _
                BuildUrlQueryParameter(FormatParamKey, FormatParamVal), _
                BuildUrlQueryParameter(VersionParamKey, VersionParamVal))
            ServiceUrl = BuildServiceUrl(, Host, Path, Query)
            ' Retrieve data.
            Set DataCollection = RetrieveDataCollection(ServiceUrl, UserAgent)
            If DataCollection Is Nothing Then
                ' CVRAPI service is not available.
                ' Has your IP been blocked?
            Else
                If DataCollection(RootItem)(CollectionItem.Data).Count > ErrorItems Then
                    Result = True
                Else
                    ' Error message returned.
                End If
            End If
        Else
            ' Search data didn't validate.
            ' No reason to bother CVRAPI.
            ' ErrorCode has been returned from ValidateSearch.
        End If
    Else
        ' Nothing to search for.
        ErrorCode = DefaultErrorCode
    End If
    
    If DataCollection Is Nothing Then
        ' Nothing to look up.
        ' Return pseudo error data.
        Set DataCollection = CvrError(ErrorCode)
    End If
    
    Set CvrLookup = DataCollection
    
End Function

' Converts an error collection to user defined type CvrError.
'
Public Function FillError( _
    ByVal DataCollection As Collection) _
    As CvrError
    
    ' Always only one root item.
    Const RootItem          As Integer = 1
    
    Dim FieldName           As String
    Dim FieldValue          As Variant
    Dim Item                As Integer
    Dim Items               As Integer
    Dim TypeError           As CvrError
    
    Items = DataCollection(RootItem)(CollectionItem.Data).Count
    
    ' Purify data and fill user defined type.
    For Item = 1 To Items
        FieldName = DataCollection(RootItem)(CollectionItem.Data)(Item)(CollectionItem.Name)
        FieldValue = DataCollection(RootItem)(CollectionItem.Data)(Item)(CollectionItem.Data)
        Select Case FieldName
            Case "error"
                TypeError.Error = Nz(FieldValue)
            Case "t"
                TypeError.T = Nz(FieldValue, 0)
            Case "version"
                TypeError.Version = Nz(FieldValue, 0)
        End Select
    Next
    
    FillError = TypeError
    
End Function

' Converts and cleans one item of a sub collection to user defined type CvrOwner.
'
Public Function FillTypeOwner( _
    ByVal DataCollection As Collection) _
    As CvrOwner
    
    Dim FieldName           As String
    Dim FieldValue          As Variant
    Dim Item                As Integer
    Dim Items               As Integer
    Dim TypeOwner           As CvrOwner
    
    Items = DataCollection.Count
    
    ' Purify data and fill user defined type.
    For Item = 1 To Items
        FieldName = DataCollection(Item)(CollectionItem.Name)
        FieldValue = DataCollection(Item)(CollectionItem.Data)
        Select Case FieldName
            Case "name"
                TypeOwner.Name = Nz(FieldValue)
        End Select
    Next
    
    FillTypeOwner = TypeOwner
    
End Function

' Converts and cleans one item of a sub collection to user defined type CvrProductionunit.
'
Public Function FillTypeProductionunit( _
    ByVal DataCollection As Collection) _
    As CvrProductionunit
    
    Dim FieldName           As String
    Dim FieldValue          As Variant
    Dim Item                As Integer
    Dim Items               As Integer
    Dim TypeProductionunit  As CvrProductionunit
    
    Items = DataCollection.Count
    
    ' Purify data and fill user defined type.
    For Item = 1 To Items
        FieldName = DataCollection(Item)(CollectionItem.Name)
        FieldValue = DataCollection(Item)(CollectionItem.Data)
        Select Case FieldName
            Case "pno"
                TypeProductionunit.Pno = FieldValue
            Case "main"
                TypeProductionunit.Main = FieldValue
            Case "name"
                TypeProductionunit.Name = FormatCompany(Nz(FieldValue))
            Case "address"
                TypeProductionunit.Address = Nz(FieldValue)
            Case "zipcode"
                TypeProductionunit.ZipCode = Nz(FieldValue)
            Case "city"
                TypeProductionunit.City = StrConv(Nz(FieldValue), vbProperCase)
            Case "protected"
                TypeProductionunit.Protected = FieldValue
            Case "phone"
                TypeProductionunit.Phone = Nz(FieldValue)
            Case "email"
                TypeProductionunit.Email = Nz(FieldValue)
            Case "fax"
                TypeProductionunit.Fax = Nz(FieldValue)
            Case "startdate"
                TypeProductionunit.Startdate = ConvertCvrDate(Nz(FieldValue))
            Case "enddate"
                TypeProductionunit.Enddate = ConvertCvrDate(Nz(FieldValue))
            Case "employees"
                TypeProductionunit.Employees = Nz(FieldValue)
            Case "addressco"
                TypeProductionunit.Addressco = Nz(FieldValue)
            Case "industrycode"
                TypeProductionunit.Industrycode = CStr(Nz(FieldValue))
            Case "industrydesc"
                TypeProductionunit.Industrydesc = Nz(FieldValue)
        End Select
    Next
    
    FillTypeProductionunit = TypeProductionunit
    
End Function

' Converts and cleans one item of a sub collection to user defined type CvrVat.
'
' 2018-05-15: Wrapped TypeVat.Creditbankrupt value in Nz().
'
Public Function FillTypeVat( _
    ByVal DataCollection As Collection) _
    As CvrVat
    
    Dim FieldName           As String
    Dim FieldValue          As Variant
    Dim Item                As Integer
    Dim Items               As Integer
    Dim TypeVat             As CvrVat
    
    ' Purify data and fill user defined type.
    Items = DataCollection.Count
    For Item = 1 To Items
        FieldName = DataCollection(Item)(CollectionItem.Name)
        Select Case FieldName
            Case "owners"
                ' Filled by FillTypeOwners.
            Case "productionunits"
                ' Filled by FillTypeProductionunit.
            Case Else
                FieldValue = DataCollection(Item)(CollectionItem.Data)
                Select Case FieldName
                    Case "vat"
                        TypeVat.VAT = FieldValue
                    Case "name"
                        TypeVat.Name = FormatCompany(Nz(FieldValue))
                    Case "address"
                        TypeVat.Address = Nz(FieldValue)
                    Case "zipcode"
                        TypeVat.ZipCode = Nz(FieldValue)
                    Case "city"
                        TypeVat.City = StrConv(Nz(FieldValue), vbProperCase)
                    Case "protected"
                        TypeVat.Protected = FieldValue
                    Case "phone"
                        TypeVat.Phone = Nz(FieldValue)
                    Case "email"
                        TypeVat.Email = Nz(FieldValue)
                    Case "fax"
                        TypeVat.Fax = Nz(FieldValue)
                    Case "startdate"
                        TypeVat.Startdate = ConvertCvrDate(Nz(FieldValue))
                    Case "enddate"
                        TypeVat.Enddate = ConvertCvrDate(Nz(FieldValue))
                    Case "employees"
                        TypeVat.Employees = Nz(FieldValue)
                    Case "addressco"
                        TypeVat.Addressco = Nz(FieldValue)
                    Case "industrycode"
                        TypeVat.Industrycode = CStr(Nz(FieldValue))
                    Case "industrydesc"
                        TypeVat.Industrydesc = Nz(FieldValue)
                    Case "companycode"
                        TypeVat.Companycode = Nz(FieldValue, 0)
                    Case "companydesc"
                        TypeVat.Companydesc = Nz(FieldValue)
                    Case "creditstartdate"
                        TypeVat.Creditstartdate = ConvertCvrDate(Nz(FieldValue))
                    Case "creditstatus"
                        TypeVat.Creditstatus = Nz(FieldValue, 0)
                    Case "creditbankrupt"
                        TypeVat.Creditbankrupt = Nz(FieldValue, 0)
                    Case "t"
                        TypeVat.T = Nz(FieldValue, 0)
                    Case "version"
                        TypeVat.Version = Nz(FieldValue, 0)
                End Select
        End Select
    Next
    
    FillTypeVat = TypeVat

End Function

' Converts and cleans one item of a full collection to user defined type CvrVat.
'
Public Function FillType( _
    ByVal DataCollection As Collection) _
    As CvrVat
    
    Dim FieldName           As String
    Dim FieldValue          As Variant
    Dim Item                As Integer
    Dim Items               As Integer
    Dim RootItem            As Integer
    Dim TypeVat             As CvrVat
    
    ' Fill user defined type.
    ' Find first active organisation/company.
    ' Note: Currently CVRAPI always returns one company only.
    Items = DataCollection.Count
    For Item = 1 To Items
        If IsNull(DataCollection(Item)(CollectionItem.Data).Item("enddate")(CollectionItem.Data)) Then
            Exit For
        End If
    Next
    If Item > Items Then
        ' No active company found.
        ' Select the first.
        Item = 1
    End If
    RootItem = Item
    TypeVat = FillTypeVat(DataCollection(RootItem)(CollectionItem.Data))
    
    ' Fill user defined sub types.
    ' Find main production unit.
    If IsNull(DataCollection(RootItem)(CollectionItem.Data)("productionunits")(CollectionItem.Data)) Then
        ' No production units.
        ' Should not happen.
    Else
        Items = DataCollection(RootItem)(CollectionItem.Data)("productionunits")(CollectionItem.Data).Count
        For Item = 1 To Items
            If DataCollection(RootItem)(CollectionItem.Data).Item("productionunits")(CollectionItem.Data).Item(Item)(CollectionItem.Data)("main")(CollectionItem.Data) = True Then
                Exit For
            End If
        Next
        If Item > Items Then
            ' No main production unit found.
            ' Select the first.
            Item = 1
        End If
        TypeVat.Productionunit = FillTypeProductionunit(DataCollection(RootItem)(CollectionItem.Data)("productionunits")(CollectionItem.Data).Item(Item)(CollectionItem.Data))
    End If
    
    ' Find owner(s).
    If IsNull(DataCollection(RootItem)(CollectionItem.Data)("owners")(CollectionItem.Data)) Then
        ' No owners registered yet.
        ' May happen.
    Else
        Items = DataCollection(RootItem)(CollectionItem.Data)("owners")(CollectionItem.Data).Count
        ' Select the first owner.
        Item = 1
        TypeVat.Owner = FillTypeOwner(DataCollection(RootItem)(CollectionItem.Data)("owners")(CollectionItem.Data).Item(Item)(CollectionItem.Data))
    End If
    
    FillType = TypeVat

End Function

' Converts CVR DK legacy ("Grandma") string date to Date value (in a Variant).
' For empty or invalid CvrDate, Null is returned.
'
' Example:
'   "01/03 - 1988" -> #1988-03-01#
'
Public Function ConvertCvrDate( _
    ByVal CvrDate As String) _
    As Variant
    
    Dim DateMonth           As String
    Dim Year                As Integer
    Dim Month               As Integer
    Dim Day                 As Integer
    Dim TrueDate            As Variant
    
    If IsDate(CvrDate) Then
        DateMonth = Split(CvrDate, "-")(0)
        Year = CInt(Split(CvrDate, "-")(1))
        Month = CInt(Split(DateMonth, "/")(1))
        Day = CInt(Split(DateMonth, "/")(0))
        TrueDate = DateSerial(Year, Month, Day)
    Else
        TrueDate = Null
    End If
    
    ConvertCvrDate = TrueDate

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

' Returns a friendly (long) error message matching an error code.
' Accepts some local error codes in addition to the CVRAPI error codes.
'
' Example:
'   Error code "NOT_FOUND" -> "Nothing matched the search criteria."
'
Public Function CvrErrorText( _
    ByVal ErrorCode As String) _
    As String
    
    Dim FriendlyError       As String
    
    ' CVRAPI error codes.
    ' NO_SEARCH     ' No useful search criteria was supplied.
    ' BANNED        ' Your IP address or IP range has been blocked. Stop further attempts.
    ' INVALID_VAT   ' Invalid VAT number or wrong format.
    ' NOT_FOUND     ' Nothing matched the search criteria.
    ' Local error codes.
    ' INVALID_PHN   ' Invalid phone number or wrong format.
    ' INVALID_PNO   ' Invalid Production Unit number or wrong format.
    
    Select Case ErrorCode
        Case "NO_SEARCH"
            FriendlyError = "No useful search criteria was supplied."
        Case "BANNED"
            FriendlyError = "Your IP address or IP range has been blocked. Stop further attempts."
        Case "INVALID_VAT"
            FriendlyError = "Invalid VAT number or wrong format."
        Case "NOT_FOUND"
            FriendlyError = "Nothing matched the search criteria."
        Case "INVALID_PNO"
            FriendlyError = "Invalid Production Unit number or wrong format."
        Case "INVALID_PHN"
            FriendlyError = "Invalid phone number or wrong format."
        Case Else
            FriendlyError = "Unspecified search error."
    End Select
    
    CvrErrorText = FriendlyError
    
End Function

' Private functions.


' Returns name of key for key/value pair of query parameter to search CVRAPI.
'
Private Function CvrSearchKeyLabel( _
    ByVal CvrKey As CvrSearchKey) _
    As String
    
    Dim Labels()            As Variant
    
    Labels() = Array("search", "vat", "name", "produ", "phone")

    CvrSearchKeyLabel = Labels(CvrKey)

End Function

' Returns name of key for key/value pair of query parameter to search CVRAPI.
'
Private Function CvrCountryKeyLabel( _
    ByVal CvrKey As CvrCountryKey) _
    As String
    
    Dim Labels()            As Variant
    
    Labels() = Array("country")

    CvrCountryKeyLabel = Labels(CvrKey)

End Function

' Returns name of key for key/value pair of query parameter to search CVRAPI.
'
Private Function CvrFormatKeyLabel( _
    ByVal CvrKey As CvrFormatKey) _
    As String
    
    Dim Labels()            As Variant
    
    Labels() = Array("format")

    CvrFormatKeyLabel = Labels(CvrKey)

End Function

' Returns name of key for key/value pair of query parameter to search CVRAPI.
'
Private Function CvrVersionKeyLabel( _
    ByVal CvrKey As CvrVersionKey) _
    As String
    
    Dim Labels()            As Variant
    
    Labels() = Array("version")

    CvrVersionKeyLabel = Labels(CvrKey)

End Function

' Returns spelled out value for key/value pair of query parameter to search CVRAPI.
'
Private Function CvrCountryValue( _
    ByVal CvrKey As CvrCountrySelect) _
    As String
    
    Dim Values()            As Variant
    
    Values() = Array("dk", "no")

    CvrCountryValue = Values(CvrKey)

End Function

' Returns spelled out value for key/value pair of query parameter to search CVRAPI.
'
Private Function CvrFormatValue( _
    ByVal CvrKey As CvrFormatSelect) _
    As String
    
    Dim Values()            As Variant
    
    Values() = Array("json", "xml")

    CvrFormatValue = Values(CvrKey)

End Function

' Returns spelled out value for key/value pair of query parameter to search CVRAPI.
'
Private Function CvrVersionValue( _
    ByVal CvrKey As CvrVersionSelect) _
    As String
    
    Dim VersionValue        As String
    
'    ' 2015-12-10.
'    ' Value 0 is now valid but does NOT return the newest version.
'
'    If CvrKey = Newest Then
'        ' Value 0 of Newest is not a valid parameter value.
'        ' Return an empty string.
'    Else
'        VersionValue = CStr(CvrKey)
'    End If

    ' Verify version.
    Select Case CvrKey
        ' Allowed versions.
        Case Newest
            ' Change to newest version as expected.
            CvrKey = Version6
        Case Version4
        Case Version5
        Case Version6
        Case Else
            ' Change disallowed versions to newest.
            CvrKey = Version6
    End Select
    
    VersionValue = CStr(CvrKey)
    
    CvrVersionValue = VersionValue

End Function

' Cleans and returns a numeric string if only valid characters are met.
' Returns an empty string if non-valid characters are located.
'
' Example:
'   CleanSearchValue(CvrSearchKey.PhoneNumber, "(12) 33.23-98") returns:
'   "12332398"
'
Private Sub CleanSearchValue( _
    ByVal SearchKey As CvrSearchKey, _
    ByRef SearchValue As String)
    
    Dim Index               As Integer
    Dim Digit               As String
    Dim Digits              As Integer
    Dim CleanValue          As String
    Dim ValueLength         As Integer
    
    Select Case SearchKey
        Case CvrSearchKey.GeneralSearch, CvrSearchKey.CompanyName
            SearchValue = Trim(SearchValue)
            
        Case CvrSearchKey.PhoneNumber, CvrSearchKey.Productionunit, CvrSearchKey.VatNo
            ValueLength = Len(SearchValue)
            CleanValue = Space(ValueLength)
            
            For Index = 1 To ValueLength
                Digit = Mid(SearchValue, Index, 1)
                Select Case Asc(Digit)
                    Case 32 To 47
                        ' Special characters allowed but ignored.
                    Case 48 To 57
                        If Digit = "0" And Digits = 0 Then
                            ' Leading zero(es) not allowed.
                            Exit For
                        Else
                            ' Insert found digit.
                            Digits = Digits + 1
                            Mid(CleanValue, Digits) = Digit
                        End If
                    Case Else
                        ' Only digits and special characters allowed.
                        CleanValue = ""
                        Exit For
                End Select
            Next
            
            SearchValue = RTrim(CleanValue)
    End Select
    
End Sub

' Checks if Number passes a Modulus 11 check.
' Returns True if passed.
'
' Note that non-standard legacy weights are applied for
' old Danish production unit numbers.
'
Private Function IsCvrModulus11( _
    ByVal Number As Double) _
    As Boolean

    Const Modulus11         As Integer = 11
    Const WeightCheck       As Integer = 1
    
    Dim Weights()           As Variant
    
    Dim Weight              As Integer
    Dim WeightSum           As Integer
    Dim Position            As Integer
    Dim ReverseNumber       As String
    Dim LengthNumber        As Integer
    Dim CheckDigit          As Integer
    Dim Result              As Boolean
    
    ' Integers only.
    If Number - Int(Number) = 0 Then
        ReverseNumber = StrReverse(CStr(Number))
        LengthNumber = Len(ReverseNumber)
        If LengthNumber = CvrPnoCodeLength.Dk And Number <= PNumberOldMax Then
            ' Legacy weights.
            Weights = Array(9, 8, 4, 6, 3, 7, 6, 5, 1)
        Else
            ' Standard weights.
            Weights = Array(2, 3, 4, 5, 6, 7, 2, 3, 4)
        End If
        
        CheckDigit = CInt(Mid(ReverseNumber, Position + 1, 1))
        WeightSum = CheckDigit * WeightCheck
        
        For Position = 2 To LengthNumber
            CheckDigit = CInt(Mid(ReverseNumber, Position, 1))
            Weight = CInt(Weights(Position - 2))
            WeightSum = WeightSum + CheckDigit * Weight
        Next
        Result = Not CBool(WeightSum Mod Modulus11)
    End If
    
    IsCvrModulus11 = Result
    
End Function

' Validates a set of search parameters to exclude those that are bound to fail.
' Returns True if validation succeeds.
' Returns False if validation fails. ErrorCode returns the error code.
' SearchValue is returned cleaned.
'
Private Function ValidateSearch( _
    ByVal CountryValue As CvrCountrySelect, _
    ByVal SearchKey As CvrSearchKey, _
    ByRef SearchValue As String, _
    ByRef ErrorCode As String) _
    As Boolean
    
    Const MinLengthSearch   As Integer = 3
    Const FixLengthPhone    As Integer = 8
    Dim Result              As Boolean
    
    Call CleanSearchValue(SearchKey, SearchValue)
    If SearchValue <> "" Then
        Select Case SearchKey
            Case CvrSearchKey.GeneralSearch, CvrSearchKey.CompanyName
                If Len(SearchValue) < MinLengthSearch Then
                    ErrorCode = "NOT_FOUND"
                End If
            Case CvrSearchKey.PhoneNumber
                If Len(SearchValue) <> FixLengthPhone Then
                    ErrorCode = "INVALID_PHN"
                End If
            Case CvrSearchKey.Productionunit
                If CountryValue = Denmark Then
                    If Len(SearchValue) <> CvrPnoCodeLength.Dk Then
                        ErrorCode = "INVALID_PNO"
                    ElseIf Val(Left(SearchValue, 1)) < CvrPnoFirstDigitMin.Dk Then
                        ErrorCode = "INVALID_PNO"
                    ElseIf Not IsCvrModulus11(CDbl(SearchValue)) Then
                        ErrorCode = "INVALID_PNO"
                    End If
                Else
                    If Len(SearchValue) <> CvrPnoCodeLength.No Then
                        ErrorCode = "INVALID_PNO"
                    ElseIf Val(Left(SearchValue, 1)) < CvrPnoFirstDigitMin.No Then
                        ErrorCode = "INVALID_PNO"
                    ElseIf Not IsCvrModulus11(CDbl(SearchValue)) Then
                        ErrorCode = "INVALID_PNO"
                    End If
                End If
            Case CvrSearchKey.VatNo
                If CountryValue = Denmark Then
                    If Len(SearchValue) <> CvrVatCodeLength.Dk Then
                        ErrorCode = "INVALID_VAT"
                    ElseIf Val(Left(SearchValue, 1)) < CvrVatFirstDigitMin.Dk Then
                        ErrorCode = "INVALID_VAT"
                    ElseIf Not IsCvrModulus11(CDbl(SearchValue)) Then
                        ErrorCode = "INVALID_VAT"
                    End If
                Else
                    If Len(SearchValue) <> CvrVatCodeLength.No Then
                        ErrorCode = "INVALID_VAT"
                    ElseIf Val(Left(SearchValue, 1)) < CvrVatFirstDigitMin.No Then
                        ErrorCode = "INVALID_VAT"
                    ElseIf Not IsCvrModulus11(CDbl(SearchValue)) Then
                        ErrorCode = "INVALID_VAT"
                    End If
                End If
        End Select
    End If
    
    If ErrorCode = "" Then
        Result = True
    Else
        SearchValue = ""
    End If
    
    ValidateSearch = Result

End Function

' Creates a collection with an error message as if this was received from CVRAPI
' using (optionally) the error code supplied in parameter ErrorCode.
'
Private Function CvrError( _
    Optional ByVal ErrorCode As String, _
    Optional ByVal VersionValue As CvrVersionSelect = CvrVersionSelect.Newest) _
    As Collection
    
    Const ErrorUnknown      As String = "UNKNOWN"
    Const TValue            As Integer = 0
    Const VersionCurrent    As Integer = CvrVersionSelect.Version6
    Const ResponseBase      As String = "{'error': '{0}','t': {1},'version': {2}}"
    
    Dim ResponseText        As String
    
    If ErrorCode = "" Then
        ErrorCode = ErrorUnknown
    End If
    If VersionValue = CvrVersionSelect.Newest Then
        VersionValue = VersionCurrent
    End If
    
    ' Build pseudo ResponseText to mimic an error response.
    ' Example:
    '   {"error": "INVALID_PNO","t": 0,"version": 6}
    ResponseText = Replace(ResponseBase, "'", Chr(34))
    ResponseText = Replace(ResponseText, "{0}", ErrorCode)
    ResponseText = Replace(ResponseText, "{1}", TValue)
    ResponseText = Replace(ResponseText, "{2}", VersionValue)
    
    ' Convert ResponseText to a collection.
    Set CvrError = CollectJson(ResponseText)
    
End Function

