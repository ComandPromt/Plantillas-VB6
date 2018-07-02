// (c) Richard Grimes 2003
// Illustrating solving some of the managed resource issues with
// managed C++. See ReadMe.txt for detailed instructions.

#include "stdafx.h"

#using <mscorlib.dll>

using namespace System;
using namespace System::Resources;
using namespace System::Reflection;

void _tmain()
{
   String* argv[] = Environment::GetCommandLineArgs();

   // We want to access the default resources for this assembly
   ResourceManager* rm = new ResourceManager(
      S"ResourceExample", Assembly::GetExecutingAssembly());

   Console::WriteLine(rm->GetString(S"MSG_LOGO"));
   if (argv->Length == 1)
   {
      Console::WriteLine(rm->GetString(S"MSG_FAIL_TOO_FEW_PARAMETERS"));
   }
   else
   {
      // We want to access the other resources embedded in this assembly
      ResourceManager* rmOther = new ResourceManager(
         S"OtherResources", Assembly::GetExecutingAssembly());

      int data = Int32::Parse(argv[1]);

      if (data > 100)
      {
         Console::WriteLine(rmOther->GetString(S"MSG_TOO_HIGH"));
      }

      if (data < 50)
      {
         Console::WriteLine(rmOther->GetString(S"MSG_TOO_LOW"));
      }
   }

   ResourceManager* rmLinked = new ResourceManager(
      S"LinkedResource", Assembly::GetExecutingAssembly());
   Console::WriteLine(rmLinked->GetString(S"MSG_DATA"));

   System::IO::Stream* mri;
   mri = Assembly::GetExecutingAssembly()->GetManifestResourceStream(S"ResourceExample.ResourceFiles.resources");
   System::Text::StringBuilder* sb = new System::Text::StringBuilder;
   for (int x = 0; x < mri->Length; x++)
   {
      Char b = static_cast<Char>(mri->ReadByte());
      if (b > 31 && b < 127)
      {
         sb->Append(b);
      }
      else
      {
         sb->Append('.');
      }
   }
   Console::WriteLine(sb->ToString());
}