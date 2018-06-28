using System;
using System.Data;
using System.Xml;
using System.Text;
using System.IO;

namespace WSWeb
{
	/// <summary>
	/// Utility Library for MSDN SQLXML 3.0 Article
	/// </summary>
	public class WSLib
	{
		public WSLib()
		{
		}

		/// <summary>
		/// Hydrates a dataset from the any XML fragment
		/// </summary>
		/// <param name="oXml"></param>
		/// <returns></returns>
		public static DataSet GetDataSetFromXmlFragment(XmlElement oXml)
		{
			// now that we have xml we can create a dataset with it or whatever
			DataSet ds = new DataSet();
			XmlTextReader oReader = new XmlTextReader(oXml.OuterXml, XmlNodeType.Element, new XmlParserContext(null, null, null, XmlSpace.None)); 
			// now lets create a schema off of the instance data
			ds.ReadXml(oReader, XmlReadMode.InferSchema);
			return ds;		
		}

		/// <summary>
		/// Retruns XML elements from an object and processs any error using SEH
		/// </summary>
		/// <param name="oaData"></param>
		/// <returns></returns>
		public static XmlElement[] GetXmlFromObjectArray(object[] oaData)
		{
			System.Xml.XmlElement[] oXmlResult = new System.Xml.XmlElement[oaData.Length];
			etier3.SqlMessage oErrorMessage; 

			for (int i = 0; i < oaData.Length; i++)
			{
				switch (oaData[i].GetType().ToString())
				{
					case "System.Xml.XmlElement":
						oXmlResult[i] = (System.Xml.XmlElement)oaData[i];
						break;
					case "WSClient.etier3.SqlMessage":
						oErrorMessage = (etier3.SqlMessage)oaData[i];
						throw new Exception("Error - Source: " + oErrorMessage.Source + " - Message: " + oErrorMessage.Message); 
					default:
						break;
				}
			}
			return oXmlResult;
		}

		/// <summary>
		/// Simply formats the output into a string builder so that its contents can be dumped to screen
		/// </summary>
		/// <param name="sProcName"></param>
		/// <param name="oXmlResult"></param>
		/// <param name="sOutput"></param>
		/// <returns></returns>
		public static DataSet FormatDataToBuffer(string sProcName, XmlElement[] oXmlResult, StringBuilder sOutput)
		{
			DataSet ds = null;
			StringWriter oWriter;
			
			for (int i=0; i < oXmlResult.Length; i++)
			{
				sOutput.Append(sProcName + " Results Schema -----------");
					
				// now hydrate our dataset... schema and all
				if (oXmlResult[i] != null)
				{
					ds = WSLib.GetDataSetFromXmlFragment(oXmlResult[i]);
					oWriter = new StringWriter();
					ds.WriteXml(oWriter, XmlWriteMode.WriteSchema);
					// send the schema to screen
					sOutput.Append(oWriter.ToString());
				}
			}

			return ds;
		}
	}
}
