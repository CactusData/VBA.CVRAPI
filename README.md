# VBA.CVRAPI
A complete collection of VBA modules and functions to call the CVRAPI using JavaScript.

As a minimum, these modules are needed:
   CvrService
   JsonCollection
   JsonScript
   JsonService

Also, in function CvrLookup, don't forget to adjust the application specific constants:

   UserAgentOrg      As String = "Anonymous"         ' Your organisation.
   UserAgentApp      As String = "Test"              ' Your app name

In most cases, an application will need functions like the example functions found in module CvrDemo.
2015-01-20.
