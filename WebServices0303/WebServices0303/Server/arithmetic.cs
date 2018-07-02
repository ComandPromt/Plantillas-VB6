using System.Diagnostics;
using System.Xml;
using System.Xml.Serialization;
using System;
using System.Web.Services.Protocols;
using System.ComponentModel;
using System.Web;
using System.Web.Services;
using System.Text;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Xml.Schema;
using Server;

[System.Web.Services.WebServiceBindingAttribute(Name="Arithmetic", Namespace="urn:msdn-microsoft-com:hows")]
public class Arithmetic : System.Web.Services.WebService {
    
    private const string soapURI = "http://schemas.xmlsoap.org/soap/envelope/";
    private const string nsURI = "urn:msdn-microsoft-com:hows";

    [XmlStreamSoapExtension]
    [WebMethod()]
    [SoapDocumentMethod("urn:msdn-microsoft-com:hows/Add", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Bare)]
    [return: XmlElement("AddResponse", Namespace="urn:msdn-microsoft-com:hows")]
    public /* AddResponse */ void Add(/* Add Add1 */)
    {
        XmlValidatingReader valid = new XmlValidatingReader(new XmlTextReader(SoapStreams.InputMessage));
        valid.Schemas.Add(XmlSchema.Read(new XmlTextReader(HttpContext.Current.Server.MapPath("server.xsd")), null));
        XmlDocument doc = new XmlDocument();
        doc.Load(valid);
        XslTransform transform = new XslTransform();
        transform.Load(new XmlTextReader(HttpContext.Current.Server.MapPath("add.xslt")));
        transform.Transform(doc, null, new XmlTextWriter(SoapStreams.OutputMessage, Encoding.UTF8));
    }
}



