# (c) Richard Grimes 2003
#
# Makefile to create an assembly that illustrates
# the app domain re-entrancy problem.
# If you are running this on v1.1 of the framework
# the assembly will be compiled with /clr. If you 
# want to see the .0 behavior then define the 
# INITDOMAIN when you invoke nmake, for example:
#
# nmake INITDOMAIN=

# Determine the framework version and pick the right 
# compiler switch
!IF "$(FRAMEWORKVERSION)" == "v1.1.4322"
CLRSWITCH=/clr
!ELSE
CLRSWITCH=/clr:initialAppDomain
!ENDIF

# Override the last decision for v1.1 if v1.0 behavior 
# is requested 
!IFDEF INITDOMAIN
CLRSWITCH=/clr:initialAppDomain
!ENDIF

all : clean Managed.exe

clean :
   -@erase *.obj
   -@erase *.lib
   -@erase *.exp
   -@erase *.dll
   -@erase *.exe

# Dependency and build rules for the managed process
Managed.exe : Native.lib Native.dll Managed.cpp
   cl $(CLRSWITCH) Managed.cpp

# Dependency and build rules for the native DLL   
Native.dll Native.lib : Native.cpp
   cl /LD Native.cpp
