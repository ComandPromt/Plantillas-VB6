<%@ WebService class="EchoClass" language="C#"%>

using System;
using System.Web;
using System.Web.Services;

[WebService(Namespace="http://example.org/echo")]
public class EchoClass : System.Web.Services.WebService
{
	[WebMethod]
	public string Echo(string input)
	{
		return String.Format("Endpoint location: {0}\nInput: {1}",
			HttpContext.Current.Request.Url, input);
	}
}

