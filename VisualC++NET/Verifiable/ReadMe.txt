Verifiable Library Assembly
===========================

(c) Richard Grimes 2003

This shows how to make an assembly verifiable. The example 
creates the assembly and then runs peverify

1) Open the project's properties and for all configurations
go to C/C++ Optimization and change the Optimization property 
to 'Disabled /Od'

2) Then go to the Linker pages and on the Advanced page change
the Resource Only DLL to 'Yes (/NOENTRY)'

3) On the Linker Input page add nochkclr.obj to the Additional 
Dependencies property.

4) Go to the main cpp file of your project and add the following:

using namespace System::Security::Permissions;
[assembly: SecurityPermissionAttribute(
              SecurityAction::RequestMinimum, 
              SkipVerification=false)];
extern "C" 
{ 
   int _dummy = 1;
}

5) Finally, you have to turn the on the COMIMAGE_FLAGS_ILONLY flag by 
running the silo tool on the library. This is part of a project called
SetILOnly which is provided as part of the MSDN library (search for this
library in the MSDN index).

To do this I have added a post-build event that runs a batch file 
verify.cmd which will call silo and then peverify to ensure that the
assembly is verifiable.
