// This is the main DLL file.

#include "stdafx.h"

#include "Verifiable.h"

using namespace System::Security::Permissions;
[assembly: SecurityPermissionAttribute(
              SecurityAction::RequestMinimum, 
              SkipVerification=false)];
extern "C" 
{
   int _dummy = 1; 
}

