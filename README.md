excelsoapgateway
================

This is a verry simple gateway between win32::OLE and SOAP. It gives SOAP clients access to basic functions of MS Excel. All you need is a windows PC running Perl and MS Office.

Using this snippet of perl code you can:
- open and close
- read or write the content of a cell of some worksheet of
an Excel file over the network using SOAP.

I do have some linux boxes doing automated mesurement tasks and wantet to write measurement results into Excel files on windows PC.

Yes, there are other solutions (sharepoint, ...) but i only needed basic functionality and wanted a solution with a small footprint.

BE WARNED:

there is no such thing like:
- session management
- user handling
- security
however, feel free to implement it :)

kaback



