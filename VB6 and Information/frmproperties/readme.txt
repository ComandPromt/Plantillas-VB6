Shell Extensions - Property Sheet Handler
-----------------------------------------

This project creates a property sheet for .frm files.

How to install the extension:
------------------------------

1)drag and drop the 2 dlls  to regsvr32.exe  to register the dll

or run the commands	regsvr32 psheet.dll
and			regsvr32 psadd.dll

2) In the explorer right click the .inf file and select Install

thats all now right click on any vbasic form file and see in the proprties there are all the form's code-dlls/objects and more



How to uninstall the extensions:
--------------------------------

To uninstall the sample use the "Add/Remove Programs" icon from 
control panel.

What is PSADD.DLL?
------------------
as i know..
This library is made in C (not by me) and is used to call the AddPage function
pointer that is passed to IShellPropSheetExt::AddPages. 

The psadd.dll must be either in the windows\system directory or
in the same directory as the compiled dll.

If you want to make your on dll the code for AddPage is:

#include <windows.h>

typedef BOOL (CALLBACK FAR * LPFNADDPROPSHEETPAGE)(LONG, LONG);

int WINAPI __export AddPage(LPFNADDPROPSHEETPAGE lpfnAddPage, LONG hPage, LONG lParam) {

    return lpfnAddPage(hPage, lParam);

}

Psadd.dll can be freely distributed with your application.


for any help contact me here: megalos@freemail.gr
if you like this just send me an email and we can trade source codes mine with yours...or if you dont have any 
..i may give it to you free just to make it better if you can...c u 

Megalos
