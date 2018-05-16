Attribute VB_Name = "CvrDebug"
Option Compare Text
Option Explicit
'
' Functions for simple testing and listing of retrieved
' data collections from CVRAPI.

' Simple example print of items from a received data collection.
'
Public Sub ListCvr( _
    ByVal DataCollection As Collection)

    Const RootItem          As Integer = 1
    Const ItemsError        As Integer = 3
    
    Dim TypeCvrVat          As CvrVat
    Dim TypeError           As CvrError
    
    If DataCollection(RootItem)(CollectionItem.Data).Count = ItemsError Then
        ' Error message received.
        ' Fill user defined type.
        TypeError.Error = DataCollection(RootItem)(CollectionItem.Data)("error")(CollectionItem.Data)
        TypeError.T = DataCollection(RootItem)(CollectionItem.Data)("t")(CollectionItem.Data)
        TypeError.Version = DataCollection(RootItem)(CollectionItem.Data)("version")(CollectionItem.Data)
        ' List error message.
        Debug.Print TypeError.Error, TypeError.T, TypeError.Version
    Else
        ' Normal data collection received.
        ' Fill (partly) user defined type.
        TypeCvrVat.VAT = DataCollection(RootItem)(CollectionItem.Data)("vat")(CollectionItem.Data)
        TypeCvrVat.Name = DataCollection(RootItem)(CollectionItem.Data)("name")(CollectionItem.Data)
        ' List two basic fields.
        Debug.Print "VAT:", CStr(TypeCvrVat.VAT)
        Debug.Print "Name:", TypeCvrVat.Name
    End If
    
End Sub

' List all field names of a received data collection from CVRAPI.
' For "owners" and "productionunits", list the count of these as well as the content.
'
' 2018-05-15: Expanded to recursively list content of "owners" and "productionunits"
'             and accept null for "owners"'s collection.
'
Public Sub ListCvrFields( _
    ByVal DataCollection As Collection, _
    Optional ByVal SubIndex As Integer)
    
    Const RootItem          As Integer = 1
    
    Dim FieldName           As String
    Dim Item                As Integer
    Dim Items               As Integer
    Dim SubItem             As Integer
    Dim SubItems            As Integer
    
    Items = DataCollection(RootItem)(CollectionItem.Data).Count
    
    For Item = 1 To Items
        FieldName = DataCollection(RootItem)(CollectionItem.Data)(Item)(CollectionItem.Name)
        Debug.Print String(Sgn(SubIndex), vbTab) & Right(Str(Item), 2), FieldName, ;
        
        If IsObject(DataCollection(RootItem)(CollectionItem.Data)(Item)(CollectionItem.Data)) Then
            ' This is a collection, either "owners" or "productionunits".
            SubItems = DataCollection(RootItem)(CollectionItem.Data)(Item)(CollectionItem.Data).Count
            Debug.Print SubItems
            ' Call ListCvrFields recursively for each entry.
            For SubItem = 1 To SubItems
                ListCvrFields DataCollection(1)(CollectionItem.Data)(Item)(CollectionItem.Data), SubItem
            Next
        Else
            ' This is a field. Print its value.
            Debug.Print DataCollection(RootItem)(CollectionItem.Data)(Item)(CollectionItem.Data)
        End If
    Next

End Sub

' Test various calls to CvrLookup.
' Example listing of returned result.
'
' Returns True if success.
'
Public Function TestCvr() As Boolean

    Dim DataCollection      As Collection
    Dim Result              As Boolean
    Dim FullResult          As CvrVat
    Dim FullError           As CvrError
    
' Unmark one line:
'
'    Set DataCollection = CvrLookup(Result, CompanyName, "bergen", Norway)
'    Set DataCollection = CvrLookup(Result, ProductionUnit, "986 326 146 ", Norway)
'    Set DataCollection = CvrLookup(Result, VatNo, "886 300 352 ", Norway)
'    Set DataCollection = CvrLookup(Result, VatNo, "12002696", Denmark)
    Set DataCollection = CvrLookup(Result, Productionunit, "1000313698", Denmark)
'    Set DataCollection = CvrLookup(Result, CompanyName, "lagkage", Denmark)
'    Set DataCollection = CvrLookup(Result, CompanyName, "YELLOW ADVERTISING NORWAY AS", Norway)
'
'   ' Will fail:
'    Set DataCollection = CvrLookup(Result, PhoneNumber, "12002696", Denmark)
    
    ' First item is 1.
    ' First value is 0.
    
    ' Root element.
    Debug.Print DataCollection(1)(CollectionItem.Name)
    ' Items.
    Debug.Print DataCollection(1)(CollectionItem.Data).Count
    ' First field (vat or error).
    Debug.Print DataCollection(1)(CollectionItem.Data)(1)(CollectionItem.Name)
    Debug.Print CStr(DataCollection(1)(CollectionItem.Data)(1)(CollectionItem.Data))
    
    If Result = True Then
        ' Success.
        FullResult = FillType(DataCollection)
        Debug.Print FullResult.VAT, FullResult.Name
    Else
        ' Error.
        FullError = FillError(DataCollection)
        Debug.Print FullError.Error, CvrErrorText(FullError.Error)
    End If
    
    ' List found data.
    Call ListCvrFields(DataCollection)
    Call ListCvr(DataCollection)
    
    Set DataCollection = Nothing
    
    TestCvr = Result

End Function

