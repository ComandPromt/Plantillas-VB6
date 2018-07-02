// (c) Richard Grimes 2003
// Illustrating the affects of /clr:initialAppDomain
// 
// I have not provided a VS.NET solution because you
// will not be able to compile it on v1.0 of the 
// framework. On v1.0 the output will be "Managed.exe"
// If you compile this on v1.1 you will get 
// "new domain" unless you specify that you want v1.0
// behavior by defining the nmake macro INITDOMAIN

#using <mscorlib.dll>
using namespace System;

#pragma comment (lib, "Native.lib")

typedef void (*FUNC) (void);
extern "C" void ExternalFunc(FUNC f);

void ShowAppDomain()
{
   AppDomain* ad = AppDomain::CurrentDomain;
   Console::WriteLine(ad->FriendlyName);
}

public __gc class ADClass : public MarshalByRefObject
{
public:
   void Proc()
   {
      ExternalFunc(::ShowAppDomain);
   }
};


void main()
{
   AppDomain* ad = AppDomain::CreateDomain(S"new domain");
   System::Reflection::Assembly* assem = System::Reflection::Assembly::GetExecutingAssembly();
   ADClass* a = static_cast<ADClass*>(ad->CreateInstanceAndUnwrap(assem->FullName, S"ADClass"));
   a->Proc();
}
