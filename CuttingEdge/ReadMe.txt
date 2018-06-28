MyTracerExe.zip 
contains the MyTracer executables. Put them into a separate folder for use with VS.NET. It represents the installed app.

TestTracer.zip
sample ASP.NET project already configured to use MyTracer

MyTracer.zip
source code of the tool


Instructions for a sample project:

1) Create the VS.NET ASP.NET project
2) Copy mydebugtool.ascx and mydebug.asmx in the root virtual folder 
3) Modify any page you want to monitor as follows
	* Register the MyDebugTool.ascx
	* Insert an instance of the control wherever in the page (preferably at the top)
	° If needed, configure the control
	* If needed, bind to the view state programmatically (see the article)
4) Configure the VS.NET project to use MyTracer (Start External Program)
5) Set the command line argument textbox to the page (i.e., the default page)

NB: Using MyTracer will cause the debugger to fail. No error or exception is thrown but no symbols are found to serve any breakpoints.
To restore the ASP.NET debugging functionality, click on Start Project which is an alternative to Start External Program
