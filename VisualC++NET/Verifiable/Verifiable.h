// Verifiable.h

#pragma once

using namespace System;
using namespace System::Runtime::InteropServices;

// Don't use unmanaged types
// __nogc class Unmanaged
// {
// public:
//    Unmanaged()
//    {
//    }
// };

public __gc class Managed
{
public:
   Managed()
   {
      // Don't use unmanaged types
      // Unmanaged u;

      // Don't use unmanaged pointers
      // int i = 0;
      // int __nogc* p = &i;

      String* s = S"a string";
      Object* o = s;
      // Don't use static_cast<> for downcasts
      // String* s1 = static_cast<String*>(o);
      String* s1 = dynamic_cast<String*>(o);

      // Interior pointers are OK
      Int32 arr[] = {0, 1, 2, 3};
      Int32* p = &arr[0];
      // But arithmetic on interior pointers is not
      // p++;
   }
};