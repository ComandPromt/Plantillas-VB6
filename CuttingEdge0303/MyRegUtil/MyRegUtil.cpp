// MyRegUtil.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    return TRUE;
}


HKEY APIENTRY InitializeRegistry()
{
	HKEY hkMyCU;
	RegCreateKey(HKEY_CURRENT_USER, "Dino", &hkMyCU);
	RegOverridePredefKey(HKEY_CURRENT_USER, hkMyCU);
	return hkMyCU;
}


void ResetRegistry(HKEY hkMyCU)
{
	RegOverridePredefKey(HKEY_CURRENT_USER, NULL);
	RegCloseKey(hkMyCU);
	return;
}