# VBA.CVRAPI #
![General](https://raw.githubusercontent.com/CactusData/VBA.CVRAPI/master/images/cvrapi.png)

### Introduction ###
A complete collection of VBA modules and functions to call the CVR API using JavaScript at:

   [CVR API](http://cvrapi.dk)
      
### Documentation ###
The official documentation for the API can be found here:

   [CVR API Documentation](http://cvrapi.dk/documentation)

### Usage ###
Tested with Microsoft Access/Excel 2016 and 365, 32- and 64-bit, but should work with any version from 2007 and forward.

As a minimum, these modules are needed:

*    CvrService 
*    JsonBase
*    JsonCollection 
*    JsonScript 
*    JsonService
   
and, in Word or Excel, for functions only found in Access:

*    Access

Also, in function CvrLookup, *don't forget to fill in the application specific constants* UserAgentOrg and UserAgentApp with the information about your application. If you don't, a messagebox will pop to remind you to do so, and the function will exit:

```
    ' Specify company name and project name in UserAgentOrg and UserAgentApp before calling
    ' the CVR API service.
    ' Build a UserAgent string that holds this info and, optionally, contact name and phone/e-mail:
    '
    '   Company - Project [- contact person [- contact phone or e-mail]]
    '
    ' Example:
    '
    '   "Contoso - CRM-system - Martin Mikkelsen +45 42424242"
    '
    ' Application specific constants.
    Const UserAgentOrg      As String = ""                  ' MUST fill in: Your organisation.
    Const UserAgentApp      As String = ""                  ' MUST fill in: Your app name.
```

#### 64-bit VBA ####
For *64-bit VBA*, a third-party dll, [*Tablacus Script Control 64*](https://tablacus.github.io/scriptcontrol_en.html), must be installed to replace the *Microsoft Script Control 1.0* as this runs with *32-bit VBA* only.


### Implementation ###
In most cases, an application will need functions like the example functions found in module CvrDemo.

2025-03-25.
