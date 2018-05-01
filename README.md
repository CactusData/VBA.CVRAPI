# VBA.CVRAPI #
![General](https://raw.githubusercontent.com/CactusData/VBA.CVRAPI/master/images/cvrapi.png)

### Introduction ###
A complete collection of VBA modules and functions to call the CVR API using JavaScript at:

   [CVR API](http://cvrapi.dk)
      
### Documentation ###
The official documentation for the API can be found here:

   [CVR API Documentation](http://cvrapi.dk/documentation)

### Usage ###
Tested with Microsoft Access/Excel 2013/2016, 32- and 64-bit, but should work with any version from 2007 and forward.

As a minimum, these modules are needed:

*    CvrService 
*    JsonCollection 
*    JsonScript 
*    JsonService
   
and, in Word or Excel, for functions only found in Access:

*    Access

Also, in function CvrLookup, *don't forget to adjust the application specific constants* UserAgentOrg and UserAgentApp.

### Implementation ###
In most cases, an application will need functions like the example functions found in module CvrDemo.

2018-05-01.
