﻿//------------------------------------------------------------------------------
// <autogenerated>
//     This code was generated by a tool.
//     Runtime Version: 1.0.3705.288
//
//     Changes to this file may cause incorrect behavior and will be lost if 
//     the code is regenerated.
// </autogenerated>
//------------------------------------------------------------------------------

// 
// This source code was auto-generated by Microsoft.VSDesigner, Version 1.0.3705.288.
// 
namespace localhost {
    using System.Diagnostics;
    using System.Xml.Serialization;
    using System;
    using System.Web.Services.Protocols;
    using System.ComponentModel;
    using System.Web.Services;
    
    
    /// <remarks/>
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="Arithmetic", Namespace="urn:msdn-microsoft-com:hows")]
    public class Arithmetic : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        /// <remarks/>
        public Arithmetic() {
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("urn:msdn-microsoft-com:hows/Add", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
        [return: System.Xml.Serialization.XmlElementAttribute("AddResponse", Namespace="urn:msdn-microsoft-com:hows")]
        public AddResponse Add([System.Xml.Serialization.XmlElementAttribute("Add", Namespace="urn:msdn-microsoft-com:hows")] Add Add1) {
            object[] results = this.Invoke("Add", new object[] {
                        Add1});
            return ((AddResponse)(results[0]));
        }
        
        /// <remarks/>
        public System.IAsyncResult BeginAdd(Add Add1, System.AsyncCallback callback, object asyncState) {
            return this.BeginInvoke("Add", new object[] {
                        Add1}, callback, asyncState);
        }
        
        /// <remarks/>
        public AddResponse EndAdd(System.IAsyncResult asyncResult) {
            object[] results = this.EndInvoke(asyncResult);
            return ((AddResponse)(results[0]));
        }
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="urn:msdn-microsoft-com:hows")]
    public class Add {
        
        /// <remarks/>
        public int n1;
        
        /// <remarks/>
        public int n2;
    }
    
    /// <remarks/>
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="urn:msdn-microsoft-com:hows")]
    public class AddResponse {
        
        /// <remarks/>
        public int sum;
    }
}