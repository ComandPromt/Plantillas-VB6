﻿//------------------------------------------------------------------------------
// <autogenerated>
//     This code was generated by a tool.
//     Runtime Version: 1.0.3705.209
//
//     Changes to this file may cause incorrect behavior and will be lost if 
//     the code is regenerated.
// </autogenerated>
//------------------------------------------------------------------------------

// 
// This source code was auto-generated by Microsoft.VSDesigner, Version 1.0.3705.209.
// 
namespace MyTracer.MyDebug {
    using System.Diagnostics;
    using System.Xml.Serialization;
    using System;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Web.Services;
    
    
    /// <remarks/>
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="MyDebugToolSoap", Namespace="MsdnMag.CuttingEdge")]
    public class MyDebugTool : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        /// <remarks/>
        public MyDebugTool() {
            this.Url = "http://localhost/mydebug/mydebug.asmx";
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("MsdnMag.CuttingEdge/GetInfo", RequestNamespace="MsdnMag.CuttingEdge", ResponseNamespace="MsdnMag.CuttingEdge", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public System.Data.DataSet GetInfo(string connString, string userKey) {
            object[] results = this.Invoke("GetInfo", new object[] {
                        connString,
                        userKey});
            return ((System.Data.DataSet)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginGetInfo(string connString, string userKey, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("GetInfo", new object[] {
                        connString,
                        userKey}, callback, asyncState);
        }
        
        /// <remarks/>
        public System.Data.DataSet EndGetInfo(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((System.Data.DataSet)(results[0]));
        }
    }
}
