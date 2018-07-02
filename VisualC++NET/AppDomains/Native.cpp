// (c) Richard Grimes 2003
// Unmanaged DLL, used to provide a function that calls
// back into managed code through a function pointer.

typedef void (*FUNC)(void);

extern "C" __declspec(dllexport)
void ExternalFunc(FUNC f)
{
   f();
}
