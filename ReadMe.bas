Attribute VB_Name = "ReadMe"
' VBA CVRAPI v1.2.2
' (c) Gustav Brock, Cactus Data ApS, CPH
' https://github.com/CactusData/VBA.CVRAPI
'
' Set of functions to retrieve data from CVRAPI.
'
' License: MIT (http://opensource.org/licenses/mit-license.php)
'
' As a minimum, these modules are needed:
'   CvrService
'   JsonCollection
'   JsonScript
'   JsonService
'
' Also, in function CvrLookup, don't forget to adjust the default
' application specific constants:
'
'   UserAgentOrg      As String = "Min organisation"  ' Your organisation.
'   UserAgentApp      As String = "Mit projekt"       ' Your app name.
'
' DON'T EVER USE this user-agent as it will at once BLOCK your IP-address at CVRAPI.
'
'   UserAgentOrg      As String = "Anonymous"         ' Your organisation.
'   UserAgentApp      As String = "Test"              ' Your app name.
'
' In most cases, an application will need functions like the
' example functions found in module CvrDemo.
'
' 2015-01-20.
' 2015-02-24 Enum CvrFormatKey.Format changed to CvrFormatKey.DataFormat to not collide with VBA function Format.
' 2015-04-02 CvrService.FormatCompany expanded to proper case company names like: "Company v/First Last".
'            Added demo functions GetCvrVat and RetrieveCvrVat.
' 2015-04-18 Default User-agent changed as described above.
'            Error "Service not available" (likely a IP address blocking) handled in CvrLookup.
'            GetCvrData added in CvrDemo. Returns a filled instance of UDT CvrVat.
' 2015-12-10 Version 0 is now allowed while Version as an empty string is not allowed.
'            Further, version 0 does not return the newest version.
'            Function CvrVersionValue modified to reflect this and validate version.
'
'            Format parameter cannot be mixed case.
'            Function CvrFormatValue modified to create lowercase format values.
'
'            Function RetrieveDataResponse, DefaultUserAgent changed to: "Min organisation - Mit projekt".
'
'            Option Compare Database changed to Option Compare Text for compatibility with Word/Excel.
'
'            Added module CvrUtil with function Nz for use in Word/Excel where Application.Nz is missing.
' 2016-04-13 CvrDemo.GetCvrDate expanded to take country code as an optional parameter.
' 2018-05-02 Module JsonScript updated to be able to run in 64-bit VBA as well.
'            For 64-bit, third-party script control must be installed separately from:
'               https://tablacus.github.io/scriptcontrol_en.html
'            Module CvrUtil renamed to Access, as it is for use in Excel only.
' 2018-05-15 CvrDebug: ListCvrFields expanded to recursively list content of "owners" and "productionsunits"
'            which also corrected a bug that caused an error if "owners" was null.
'            CvrService: Wrapped TypeVat.Creditbankrupt value in Nz() to prevent error if value was null.

