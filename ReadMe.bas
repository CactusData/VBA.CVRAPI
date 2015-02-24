Attribute VB_Name = "ReadMe"
' VBA CVRAPI v1.0.1
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
' Also, in function CvrLookup, don't forget to adjust the
' application specific constants:
'
'   UserAgentOrg      As String = "Anonymous"         ' Your organisation.
'   UserAgentApp      As String = "Test"              ' Your app name.
'
' In most cases, an application will need functions like the
' example functions found in module CvrDemo.
'
' 2015-01-20.
' 2015-02-24 Enum CvrFormatKey.Format changed to CvrFormatKey.DataFormat to not collide with VBA function Format.
