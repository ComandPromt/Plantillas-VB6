
#pragma warning( disable: 4049 )  /* more than 64k source lines */

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 5.03.0280 */
/* at Thu Oct 03 00:24:13 2002
 */
/* Compiler settings for C:\Scratch\TriggerServer\TriggerServer.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32 (32b run), ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __TriggerServer_h__
#define __TriggerServer_h__

/* Forward Declarations */ 

#ifndef __ITrigger_FWD_DEFINED__
#define __ITrigger_FWD_DEFINED__
typedef interface ITrigger ITrigger;
#endif 	/* __ITrigger_FWD_DEFINED__ */


#ifndef __IDataChangedEvents_FWD_DEFINED__
#define __IDataChangedEvents_FWD_DEFINED__
typedef interface IDataChangedEvents IDataChangedEvents;
#endif 	/* __IDataChangedEvents_FWD_DEFINED__ */


#ifndef __Trigger_FWD_DEFINED__
#define __Trigger_FWD_DEFINED__

#ifdef __cplusplus
typedef class Trigger Trigger;
#else
typedef struct Trigger Trigger;
#endif /* __cplusplus */

#endif 	/* __Trigger_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __ITrigger_INTERFACE_DEFINED__
#define __ITrigger_INTERFACE_DEFINED__

/* interface ITrigger */
/* [unique][helpstring][uuid][object] */ 


EXTERN_C const IID IID_ITrigger;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("C1AB316B-2A13-406D-8C26-CC1858AE7942")
    ITrigger : public IUnknown
    {
    public:
        virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE DataChanged( 
            /* [string][in] */ wchar_t __RPC_FAR *strTableName) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ITriggerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            ITrigger __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            ITrigger __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            ITrigger __RPC_FAR * This);
        
        /* [helpstring] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DataChanged )( 
            ITrigger __RPC_FAR * This,
            /* [string][in] */ wchar_t __RPC_FAR *strTableName);
        
        END_INTERFACE
    } ITriggerVtbl;

    interface ITrigger
    {
        CONST_VTBL struct ITriggerVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITrigger_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define ITrigger_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define ITrigger_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define ITrigger_DataChanged(This,strTableName)	\
    (This)->lpVtbl -> DataChanged(This,strTableName)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring] */ HRESULT STDMETHODCALLTYPE ITrigger_DataChanged_Proxy( 
    ITrigger __RPC_FAR * This,
    /* [string][in] */ wchar_t __RPC_FAR *strTableName);


void __RPC_STUB ITrigger_DataChanged_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __ITrigger_INTERFACE_DEFINED__ */


#ifndef __IDataChangedEvents_INTERFACE_DEFINED__
#define __IDataChangedEvents_INTERFACE_DEFINED__

/* interface IDataChangedEvents */
/* [unique][helpstring][uuid][object] */ 


EXTERN_C const IID IID_IDataChangedEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("EC50A8A7-E7B2-4734-8822-AD6ED818782B")
    IDataChangedEvents : public IUnknown
    {
    public:
        virtual /* [helpstring] */ HRESULT STDMETHODCALLTYPE DataChanged( 
            /* [string][in] */ wchar_t __RPC_FAR *strTableName) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IDataChangedEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IDataChangedEvents __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IDataChangedEvents __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IDataChangedEvents __RPC_FAR * This);
        
        /* [helpstring] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *DataChanged )( 
            IDataChangedEvents __RPC_FAR * This,
            /* [string][in] */ wchar_t __RPC_FAR *strTableName);
        
        END_INTERFACE
    } IDataChangedEventsVtbl;

    interface IDataChangedEvents
    {
        CONST_VTBL struct IDataChangedEventsVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IDataChangedEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IDataChangedEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IDataChangedEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IDataChangedEvents_DataChanged(This,strTableName)	\
    (This)->lpVtbl -> DataChanged(This,strTableName)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring] */ HRESULT STDMETHODCALLTYPE IDataChangedEvents_DataChanged_Proxy( 
    IDataChangedEvents __RPC_FAR * This,
    /* [string][in] */ wchar_t __RPC_FAR *strTableName);


void __RPC_STUB IDataChangedEvents_DataChanged_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IDataChangedEvents_INTERFACE_DEFINED__ */



#ifndef __TRIGGERSERVERLib_LIBRARY_DEFINED__
#define __TRIGGERSERVERLib_LIBRARY_DEFINED__

/* library TRIGGERSERVERLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_TRIGGERSERVERLib;

EXTERN_C const CLSID CLSID_Trigger;

#ifdef __cplusplus

class DECLSPEC_UUID("40AFBBB2-C389-412B-AA2A-3483366C6835")
Trigger;
#endif
#endif /* __TRIGGERSERVERLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


