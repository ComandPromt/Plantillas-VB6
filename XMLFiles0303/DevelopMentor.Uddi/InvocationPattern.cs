using System;
using System.Net;
using System.Xml;
using System.Configuration;
using Microsoft.Uddi;

namespace DevelopMentor.Uddi
{
	public class InvocationPattern
	{
		internal static string GetUddiServerLocation(string keyName)
		{
			string url = ConfigurationSettings.AppSettings[keyName];
			return (url != null) ? url : "http://test.uddi.microsoft.com/inquire";
		}

		internal static bool NotFound(WebException we)
		{
			if (we.Status == WebExceptionStatus.ConnectFailure)
				return true;
			else if (we.Response != null)
			{
				HttpWebResponse r = we.Response as HttpWebResponse;
				switch (r.StatusCode)
				{
					case HttpStatusCode.Gone:
					case HttpStatusCode.NotFound:
					case HttpStatusCode.Moved:
					case HttpStatusCode.Redirect:
						return true;
					default:
						return false;
				}
			}
			return false;
		}

		public static string FindCurrentLocation(WebException we, string bindingKeyName, string uddiKeyName)
		{
			if (NotFound(we))
			{
				Inquire.Url = GetUddiServerLocation(uddiKeyName);
				GetBindingDetail gbd = new GetBindingDetail();
				gbd.BindingKeys.Add(ConfigurationSettings.AppSettings[bindingKeyName]);
				BindingDetail bd = gbd.Send();
				if (bd != null && bd.BindingTemplates.Count > 0)
					return bd.BindingTemplates[0].AccessPoint.Text;
			}
			return "";
		}

		public static void UpdateLocalConfiguration(string configFileName, string keyName, string keyValue)
		{
			XmlDocument doc = new XmlDocument();
			doc.Load(configFileName);
			string query = string.Format("/configuration/appSettings/add[@key='{0}']", keyName);
			XmlNode key = doc.SelectSingleNode(query);
			if (key != null)
			{
				XmlElement keyElement = key as XmlElement;
				keyElement.SetAttribute("value", keyValue);
				doc.Save(configFileName);
			}
		}
	}
}
