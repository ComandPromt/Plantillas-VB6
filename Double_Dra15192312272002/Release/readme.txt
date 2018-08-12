Lots of people are wondering what are these typelibs and where can they find 
the source idl files.

The only reason I'm distributing these files is because the vbp files are set 
to project compatibility using these typelibs as target files. This way the 
samples included in the zip does not lose references upon recompilation of the 
Outlook Bar control.

These typelibs are produced by Visual Basic. Open Project Properties dialog 
and navigate to Component tab. Find the Remote Server Files check box. This 
option effectively turns typelib compilation on and off for an ActiveX dll 
or ocx project.