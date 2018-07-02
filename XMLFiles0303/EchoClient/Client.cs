using System;
using System.Net;
using System.Configuration;
using DevelopMentor.Uddi;

using System.Xml;

namespace EchoClient
{
	class Client
	{
		static void Main(string[] args)
		{
			if (args.Length < 1)
			{
				Console.WriteLine("usage: echoclient input-string");
				return;
			}

			EchoClass proxy = new EchoClass();

			try
			{
				Console.WriteLine(proxy.Echo(args[0]));
			}
			catch(Exception e)
			{
				if (e is WebException)
				{
					// contact UDDI server for current location
					Console.WriteLine("Contacting UDDI server for current location...");
					string newLoc = InvocationPattern.FindCurrentLocation((WebException)e, "EchoServiceUddiBindingKey", "UddiServerLocation");
					if (newLoc != "")
					{
						proxy.Url = newLoc;
						Console.WriteLine(proxy.Echo(args[0]));
						// if it succeeds, update local configuration file
						Console.WriteLine("Updating client configuration...");
						InvocationPattern.UpdateLocalConfiguration("echoclient.exe.config", "EchoServiceLocation", newLoc);
						return;
					}
				}
				Console.WriteLine(e.Message);
			}
		}
	}
}
