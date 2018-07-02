#include <windows.h>
#include <srv.h>

///////////////////////////////////////////////////////////////////////
// Entry point

extern "C" BOOL WINAPI DllMain (HINSTANCE hInstance, DWORD dwReason,
    LPVOID lpReserved)
{
    return TRUE;
}

///////////////////////////////////////////////////////////////////////
// Exported functions

extern "C" __declspec (dllexport)
ULONG __GetXpVersion ()
{
   return ODS_VERSION;
}

extern "C" __declspec (dllexport)
SRVRETCODE xsp_UpdateSignalFile (SRV_PROC *srvproc)
{
	//
	// Make sure an input parameter is present.
	//
	if (srv_rpcparams (srvproc) == 0)
		return -1;

	//
	// Extract the file name from the input parameter.
	//
	BYTE bType;
	char file[256];
	ULONG ulMaxLen = sizeof (file);
	ULONG ulActualLen;
	BOOL fNull;

	if (srv_paraminfo (srvproc, 1, &bType, &ulMaxLen, &ulActualLen,
		(BYTE*) file, &fNull) == FAIL)
		return -1;

	if (bType != SRVBIGCHAR && bType != SRVBIGVARCHAR)
		return -1;

	file[ulActualLen] = 0;

	//
	// Update the file's time stamp.
	//
	char path[288] = "C:\\AspNetSql\\";
	lstrcat (path, file);

	HANDLE hFile = CreateFile (path, GENERIC_WRITE, 0, NULL,
		CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

	if (hFile != INVALID_HANDLE_VALUE)
		CloseHandle (hFile);

	return 0;
}
