using System;
using System.Web;
using System.Xml;
using System.Data;
using System.Data.SqlClient;

namespace DevHawkBooks
{

	public class Controller
	{
		public string GetTransform()
		{
			HttpContext ctx = HttpContext.Current;
			
			if (ctx.Request.QueryString["xslt"] == "blank")
				return "";

			HttpCookie c = ctx.Request.Cookies.Get("skin");

			if (c == null)
				return HttpContext.Current.Request.PhysicalApplicationPath + "simple.xslt";

			return HttpContext.Current.Request.PhysicalApplicationPath + c.Value;
		}

		public XmlNode SetSkin()
		{
			HttpContext ctx = HttpContext.Current;

			if (ctx.Request.QueryString["skin"] == "clear")
			{
				HttpCookie c = ctx.Request.Cookies.Get("skin");
				if (c != null)
				{
					c.Expires = DateTime.Now;
					ctx.Response.Cookies.Add(c);
				}
				ctx.Request.Cookies.Remove("skin");
			}
			else if (ctx.Request.QueryString["skin"] != null)
			{
				ctx.Response.Cookies.Add(new HttpCookie("skin", ctx.Request.QueryString["skin"]));
			}

			return null;
		}
	}
}
