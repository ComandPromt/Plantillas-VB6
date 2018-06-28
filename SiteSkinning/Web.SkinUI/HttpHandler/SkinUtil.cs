namespace DevHawk.Web.SkinUI
{
    using System;
	using System.Xml;
	using System.Xml.Xsl;
	using System.Web;
	using System.Collections.Specialized;
	using System.Data;
	using System.Data.SqlClient;

public class SkinUtil
{
	public const string SkinUINamespace = "urn:schemas-DevHawk-net:webskinui";

	public static void AppendNode(XmlNode parent, XmlNode node)
	{
		if (node == null)
			return;

		if (node.OwnerDocument != parent.OwnerDocument)
			node = parent.OwnerDocument.ImportNode(node, true);

		parent.AppendChild(node);
	}

	public static XmlNode NameValCol(XmlNode parent, NameValueCollection col, string val, string nsName)
	{
		if (nsName == null)
			nsName = parent.NamespaceURI;

		if (val != null)
			return NameValColWorkhorse(parent, val, col.GetValues(val), nsName);

		XmlDocumentFragment frag = parent.OwnerDocument.CreateDocumentFragment();
		for (int x = 0; x < col.Count; x++)
			AppendNode(frag, NameValColWorkhorse(parent, col.GetKey(x), col.GetValues(x), nsName));

		return frag;
	}

	private static XmlNode NameValColWorkhorse(XmlNode parent, string name, string[] arVal, string nsName)
	{
		if (arVal == null)
			return null;

		XmlDocumentFragment frag = parent.OwnerDocument.CreateDocumentFragment();

		for (int x = 0; x < arVal.Length; x++)
		{
			XmlElement e = parent.OwnerDocument.CreateElement(name, nsName);
			e.InnerText = arVal[x];
			frag.AppendChild(e);
		}

		return frag;
	}

	public static XmlNode CookieCol(XmlNode parent, HttpCookieCollection col, string val, string nsName)
	{
		if (nsName == null)
			nsName = parent.NamespaceURI;

		if (val != null)
			return CookieColWorkhorse(parent, col[val], nsName);

		XmlDocumentFragment frag = parent.OwnerDocument.CreateDocumentFragment();
		for (int x = 0; x < col.Count; x++)
			AppendNode(frag, CookieColWorkhorse(parent, col[x], nsName));

		return frag;
	}

	private static XmlNode CookieColWorkhorse(XmlNode parent, HttpCookie cookie, string nsName)
	{
		if (cookie == null)
			return null;

		XmlElement e = parent.OwnerDocument.CreateElement(cookie.Name, nsName);
		e.SetAttribute("domain", cookie.Domain);
		e.SetAttribute("expires", cookie.Expires.ToString());
		e.SetAttribute("path", cookie.Path);
		e.SetAttribute("secure", cookie.Secure.ToString());

		if (cookie.HasKeys)
		{
			for (int x = 0; x < cookie.Values.Count; x++)
			{
				XmlElement sub = parent.OwnerDocument.CreateElement(cookie.Values.GetKey(x), parent.NamespaceURI);
				sub.InnerText = cookie[cookie.Values.GetKey(x)];
				e.AppendChild(sub);
			}
		}
		else
			e.InnerText = cookie.Value;

		return e;

	}

	public static SqlCommand Command(string sql, SqlConnection con, string type)
	{
		//Create a new cmd for the connection
		SqlCommand cmd = new SqlCommand(sql.Trim(), con);

		//if the type attribute is specified, use it to set the command type property
		if (type != null)
			cmd.CommandType = (CommandType)Enum.Parse(cmd.CommandType.GetType(), type, true);

		return cmd;
	}

	public static SqlParameter Parameter(string name, string datatype, string size, string val, 
										 string def, string key, string collection, HttpContext context)
	{
		//convert the datatype attribute into a SqlDataType enum value
		if (datatype == null)
			throw new Exception("Missing datatype attribute on parameter element");

		SqlDbType dt = new SqlDbType();
		dt = (SqlDbType)Enum.Parse(dt.GetType(), datatype, true);

		//Create a parameter w/ name and type as specified in the XmlNode
		SqlParameter p; 

		//If the optional size attribute is specified, pass it to the parameter constructor 
		if (name == null)
			throw new Exception("Missing name attribute on parameter element");

		if (size != null)
			p = new SqlParameter(name, dt, Convert.ToInt32(size));
		else 
			p = new SqlParameter(name, dt);

		//If the XmlNode specified a value attribute, use it as the parameter value
		if (val != null)
			p.Value = val;
		else
		{
			if (collection == null || key == null)
				throw new Exception("you must specify either a value or both collection and key attributes for a query/parameter tag");

			switch (collection.ToLower())
			{
				case "querystring":
					p.Value = context.Request.QueryString[key];
					break;
				case "form":
					p.Value = context.Request.Form[key];
					break;
				case "params":
					p.Value = context.Request.Params[key];
					break;
				case "servervariables":
					p.Value = context.Request.ServerVariables[key];
					break;
				case "cookies":
					p.Value = context.Request.Cookies[key].Value;
					break;
				default:
					throw new Exception("Invalid collection \"" + collection + "\" attribute on parameter tag");
			}

			if (p.Value == null && def != null)
				p.Value = def;

		}
		return p;
		
	}

	public static XmlNode Execute(XmlDocument dom, SqlCommand cmd) 
	{
		//Retrieve the result into an XML 
		XmlReader xr = cmd.ExecuteXmlReader();

		XmlDocumentFragment xdf = dom.CreateDocumentFragment();
		XmlNode xn;

		while ((xn = dom.ReadNode(xr)) != null)
		{
			xdf.AppendChild(xn);
		}

		xr.Close();

		return xdf;
	}

	public static void Transform(XmlDocument dom, string xsltFile, HttpContext context)
	{
		if ((xsltFile == null) || (xsltFile == ""))
		{
		    dom.Save(context.Response.Output);
		}
		else 
		{
			XslTransform xslt = (XslTransform)(context.Cache[xsltFile]);
		    if (xslt == null) 
			{
		        xslt = new XslTransform();
		        xslt.Load(xsltFile);
		        context.Cache.Insert(xsltFile, xslt, new System.Web.Caching.CacheDependency(xsltFile));
		    }
		    xslt.Transform(dom.DocumentElement, null, context.Response.Output);
		}
	}
	
}

}
