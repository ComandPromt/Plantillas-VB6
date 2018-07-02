rem Make sure that you build the SetILOnly project and add the
rem silo.exe file to the path (or add the path in the call below)
@echo ----------
@echo Setting the ILOnly flag
silo -s %1
@echo Checking that the library is verifiable
peverify %1
@echo ----------
