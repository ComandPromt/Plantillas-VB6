namespace DevHawk.Web.SkinUI
{
using System;
using System.Collections;
using System.Collections.Specialized;
using System.Web;
using System.Xml;
using System.Reflection;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// HTTP Handler for SkinUI Controller XML Templates
/// </summary>
public class SkinHandler : IHttpHandler
{
	/// <summary>
	/// helper class for storing the name of the business object and method for
	/// retrieving the XSLT transform for this handler
	/// </summary>
	protected class TransformInfo
	{
		public string variable;
		public string method;
		public string defaultTransform;

		/// <summary>
		/// constructor that reads the variable and method information from 
		/// the XmlNode parameter
		/// </summary>
		/// <param name="n">The current xml:transform node</param>
		public TransformInfo(XmlNode n)
		{
			XmlElement e = (XmlElement)n;
			variable = e.GetAttribute("var");
			method = e.GetAttribute("method");
			defaultTransform = e.GetAttribute("default");
		}
	}

	/// <summary>
	/// Member of IHttpHandler interface. 
	/// Since there is no object state, the object instance can always be 
	/// reused, so this method always returns true.
	/// </summary>
	public bool IsReusable
	{
		get
		{
			return true;
		}
	}


	/// <summary>
	/// Member of the IHttpHandler interface
	/// Process a dynamic web page request
	/// </summary>
	/// <param name="context">represents the current web page request</param>
	public void ProcessRequest(HttpContext context)
	{
#if DEBUG
		if (context.Request["__admin"] != null)
		{
			if (context.Request["__admin"].ToUpper() == "showXml".ToUpper())
			{
				System.Xml.XmlDocument dom = new System.Xml.XmlDocument();
				dom.Load(context.Request.PhysicalPath);
				dom.Save(context.Response.Output);

				return;
			}
		}
#endif

		ListDictionary objDict = null;

		try
		{
			//Load the XML document indicated in the request
			XmlDocument dom = new XmlDocument();
			dom.Load(context.Request.PhysicalPath);

			//get a list of all the nodes in the SkinUI Namespace
			XmlNodeList l = dom.GetElementsByTagName("*", SkinUtil.SkinUINamespace);

			//Initalize the XSLT variable
			//Create a new object dictionary to store variable declarations
			TransformInfo xsltInfo = null;
			objDict = new ListDictionary();

			//process each of the SkinUI Nodes. Note, since we delete each WebSkin
			//node as we process it, we just need to keep processing the top element
			//of the list until the list is empty.
			while (l.Count > 0)
			{
				XmlNode n = l[0];
				switch (n.LocalName)
				{
					case "class": 
						ClassDecl(context, n, objDict); 
						break;
					case "methodcall":
						MethodCall(context, n, objDict);
						break;
					case "database":
						DatabaseDecl(n, objDict); 
						break;
					case "query":
						Query(context, n, objDict);
						break;
					case "requestvar":
						RequestVar(context.Request, n);
						break;
					case "transform":
						if (xsltInfo != null)
							throw new Exception("Multiple SkinUI Transform tags found");
						xsltInfo = new TransformInfo(n);
						n.ParentNode.RemoveChild(n);
						break;
					default:
						throw new Exception("Unrecognized SkinUI tag name: " + n.LocalName);
				}
			}

			//Delete any namespace declarations off the root node for the SkinUI namespace
			for (int i = dom.DocumentElement.Attributes.Count - 1; i >= 0; i--)
			{
				XmlAttribute a = dom.DocumentElement.Attributes[i];
				if (a.Value == SkinUtil.SkinUINamespace && a.Prefix == "xmlns")
					a.OwnerElement.Attributes.Remove(a);
			}

			//Get the XSLT file to use in transforms from the controller class
			string xsltFile = "";

			if (xsltInfo != null)
			{
				if (xsltInfo.defaultTransform != "")
					xsltFile = context.Request.PhysicalApplicationPath + xsltInfo.defaultTransform;
				else
				{
					Object o = objDict[xsltInfo.variable];
					xsltFile = GetXSLTemplate(o, xsltInfo.method, context);
				}
			}
		
			//Call the helper function to perform the actual XSLT Transform
			SkinUtil.Transform(dom, xsltFile, context);
		}
		catch (Exception ex)
		{
			//Catch any exception and add it to the context as an HTTP exception
			HttpException h = new HttpException(500, ex.Message, ex);
			context.AddError(h);
		}
		finally
		{
			if (objDict != null)
			{
				//Close and release all entries in the object dictionary
				foreach(DictionaryEntry entry in objDict)
				{
					if (entry.Value is SqlConnection)
						((SqlConnection)entry.Value).Close();
				}
			}
		}
	}

	/// <summary>
	/// Function to process the "requestvar" skin tag. Determines the specified 
	/// collection and calls out to a helper function in SkinUtil to process
	/// the directive from the specified Http collection.
	/// </summary>
	/// <param name="req">The current HTTP Request</param>
	/// <param name="n">The current skin directive node</param>
	protected void RequestVar(HttpRequest req, XmlNode n)
	{
		//Get the specified http request collection
		string colName = GetAttribNull(n, "collection");
	
		//if the collection is not specified, default to the "params" collection
		if (colName == null) colName = "Params";

		switch (colName)
		{
			case "QueryString":
				//Since the querystring, form, servervariables and params collections
				//are simple name value collections and all of the same type, the 
				//code to process the requestvar directive has been factored into a 
				//helper function.
				NameValCol(req.QueryString, n);
				break;
			case "Form":
				NameValCol(req.Form, n);
				break;
			case "ServerVariables":
				NameValCol(req.ServerVariables, n);
				break;
			case "Params":
				NameValCol(req.Params, n);
				break;
			case "Cookies":
				//since cookies is handled seperately, there's no point in factoring
				//the code into a seperate function. 
				XmlNode newNode = SkinUtil.CookieCol(n.ParentNode, req.Cookies, GetAttribNull(n, "name"), GetAttribNull(n, "namespace"));
				ReplaceNode(n.ParentNode, newNode, n);
				break;
			default:
				throw new Exception("Unrecognized Request Var collection name: " + colName);
		}
	}

	/// <summary>
	/// helper function to process generic HTTP request name value collections
	/// </summary>
	/// <param name="col">the collection to retrieve the value(s) from</param>
	/// <param name="n">The current skin directive node</param>
	protected void NameValCol(NameValueCollection col, XmlNode n)
	{
		XmlNode newNode = SkinUtil.NameValCol(n.ParentNode, col, GetAttribNull(n, "name"), GetAttribNull(n, "namespace"));
		ReplaceNode(n.ParentNode, newNode, n);
	}

	/// <summary>
	/// helper method to call the methodName method via reflection
	/// </summary>
	/// <param name="request">HttpRequest object from the current HttpContext</param>
	/// <returns>XSLT file name to use in transform</returns>
	protected string GetXSLTemplate(object o, string methodName, HttpContext context)
	{
		Type t = o.GetType();
		MethodInfo m = t.GetMethod(methodName);

		if (m == null)
			return "";

		Object ret = m.Invoke(o, null); 
		return (string)ret;
	}
	
	/// <summary>
	/// Create an instance of the object described in the XmlNode and store it 
	/// in the object dictionary
	/// </summary>
	/// <param name="context">the http context, used to get the current app path</param>
	/// <param name="n">the class node</param>
	/// <param name="objDict">the object dictionary that stores the instances</param>
	protected void ClassDecl(HttpContext context, XmlNode n, IDictionary objDict)
	{
		//Create class instance via reflection
		string filename = context.Request.PhysicalApplicationPath + @"bin\" + GetAttribException(n, "assembly") + ".dll";
		Assembly a = Assembly.LoadFrom(filename);
		Object o = a.CreateInstance(GetAttribException(n, "class"), true);

		//Add class instance to object dictionary
		objDict.Add(GetAttribException(n, "var"), o);

		//remove class element from XML document
		n.ParentNode.RemoveChild(n);
	}

	/// <summary>
	/// Create an instance of a SQL connection to the database described in
	/// the XmlNode and store it in the object dictionary
	/// </summary>
	/// <param name="n">The database xml node</param>
	/// <param name="objDict">the object dictionary that stores the 
	/// SQL connection instance</param>
	protected void DatabaseDecl(XmlNode n, IDictionary objDict)
	{
		//Create a database connection object
		SqlConnection con = new SqlConnection(GetAttribException(n, "connectionstring"));
		con.Open();

		//Add database connection instance to object dictionary
		objDict.Add(GetAttribException(n, "var"), con);

		//remove database element from XML document
		n.ParentNode.RemoveChild(n);
	}


	/// <summary>
	/// Call the method described in the XmlNode, passing in the context as a 
	/// parameter and adding the XmlNode it returns (if there is one) to the DOM
	/// </summary>
	/// <param name="context">The HTTP context that is passed to the method</param>
	/// <param name="n">The XmlNode that describes the instance and method to call</param>
	/// <param name="objDict">The object dictionary that stores all the object instances</param>
	protected void MethodCall(HttpContext context, XmlNode n, IDictionary objDict)
	{
		//Retrieve the method object via reflection
		Object o = objDict[GetAttribException(n, "var")];
		Type t = o.GetType();
		MethodInfo m = t.GetMethod(GetAttribException(n, "method"));

		//Invoke the desired method and Replace the directive node w/ the 
		//return value of the invoked method
		XmlNode node = (XmlNode)m.Invoke(o, null); 
		ReplaceNode(n.ParentNode, node, n);
	}

	/// <summary>
	/// Execute the text of the node against the database indicated in the XmlNode
	/// </summary>
	/// <param name="context">The HTTP context, used for parameter variables</param>
	/// <param name="n">The XmlNode describing the database to run the query against</param>
	/// <param name="objDict">The object dictionary that stores all the connection instances</param>
	protected void Query(HttpContext context, XmlNode n, IDictionary objDict)
	{
		//Get the connection from the object dictionary
		SqlConnection con = (SqlConnection)objDict[GetAttribException(n, "var")];

		if (con == null)
			throw new Exception("There is no SQL connection with the var name " + GetAttribException(n, "var"));

		//Create a new cmd for the connection
		SqlCommand cmd = SkinUtil.Command(n.InnerText, con, GetAttribNull(n, "type"));

		//Iterate through the childern of the query node
		XmlNode child = n.FirstChild;
		while (child != null)
		{
			//look for parameter elements
			//Each parameter element represents a parameter to pass to the SQL command
			if (child.NodeType == XmlNodeType.Element && child.LocalName == "parameter")
				cmd.Parameters.Add(Parameter(child, context));

			//Move to next query child node
			child = child.NextSibling;
		}

		//Execute the query and replace the current node w/ the result of the query.
		XmlNode newNode = SkinUtil.Execute(n.OwnerDocument, cmd);
		ReplaceNode(n.ParentNode, newNode, n);
	}

	/// <summary>
	/// helper function to process parameter tags
	/// </summary>
	/// <param name="n">the parameter skin directive tag</param>
	/// <param name="context">the current HTTP context</param>
	/// <returns>a SQL parameter obejct that has the name specified in the 
	/// directive tag and a value either from the tag or one of the 
	/// http request collections</returns>
	protected SqlParameter Parameter(XmlNode n, HttpContext context)
	{
		return SkinUtil.Parameter(GetAttribNull(n, "name"), GetAttribNull(n, "datatype"),
							  GetAttribNull(n, "size"), GetAttribNull(n, "value"),
							  GetAttribNull(n, "default"), GetAttribNull(n, "key"), 
							  GetAttribNull(n, "collection"), context);
	}

	/// <summary>
	/// Helper method to return the specified attribute of a node, throwing an 
	/// exception if it doesn't exist
	/// </summary>
	/// <param name="n">the node to read attribute from</param>
	/// <param name="attribute">the name of the attribute to retrieve</param>
	/// <returns>the attribute value as a string
	/// (Throws exception if attribute doesn't exist)</returns>
	protected string GetAttribException(XmlNode n, string attribute)
	{
		if (n.Attributes[attribute] == null)
			throw new Exception("XML " + n.Name + " Node missing required " + attribute + " attribute");

		return n.Attributes[attribute].Value;
	}

	/// <summary>
	/// helper function to return a specified attribute of a node, returning NULL if 
	/// the attribute doesn't exist.
	/// </summary>
	/// <param name="n">the node to read the attribute from</param>
	/// <param name="a">the name of the attribute to retrieve </param>
	/// <returns>the attribute value as a string 
	/// (null if the attribute doesn't exist)</returns>
	protected string GetAttribNull(XmlNode n, string a)
	{
		string val = null;
		
		if (n.Attributes[a] != null)
			val = n.Attributes[a].Value;

		return val;
	}

	/// <summary>
	/// Helper function to replace a given XML node in the DOM (typically a 
	/// skin directive) with a new node. If the new node is null, just remove
	/// the old one. If the new node is not from the same DOM tree, import it
	/// into the DOM first before doing the replacement.
	/// </summary>
	/// <param name="parent">The node that is the parent to the node to be replaced</param>
	/// <param name="node">The node to insert into the DOM (may be null)</param>
	/// <param name="refNode">The node to remove from the DOM</param>
	protected void ReplaceNode(XmlNode parent, XmlNode node, XmlNode refNode)
	{
		if (node == null)
			parent.RemoveChild(refNode);
		else
		{
			if (parent.OwnerDocument != node.OwnerDocument)
				node = parent.OwnerDocument.ImportNode(node, true);

			parent.ReplaceChild(node, refNode);
		}			
	}
}
}
