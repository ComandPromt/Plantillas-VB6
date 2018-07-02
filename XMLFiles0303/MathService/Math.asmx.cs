using System;
using System.Web;
using System.Web.Services;

namespace MathService
{
	[WebService(Namespace="http://example.org/math")]
	public class MathClass : WebService
	{
		[WebMethod]
		public double Add(double x, double y)
		{
			return x+y;
		}

		[WebMethod]
		public double Sub(double x, double y)
		{
			return x-y;
		}

		[WebMethod]
		public double Mul(double x, double y)
		{
			return x*y;
		}

		[WebMethod]
		public double Div(double x, double y)
		{
			return x/y;
		}

		[WebMethod]
		public double Mod(double x, double y)
		{
			return x%y;
		}

		[WebMethod]
		public double CalcDistance(Point orig, Point dest)
		{
			return Math.Sqrt(Math.Pow(orig.x-dest.x, 2) + Math.Pow(orig.y-dest.y, 2));
		}

		[WebMethod]
		public Point[] GetPoints()
		{
			return new Point[] { new Point(23, 44), new Point(11, 19), new Point(98, 82) };
		}
	
	}

	public class Point
	{
		public double x;
		public double y;

		public Point() {}
		public Point(double x, double y) { this.x = x; this.y = y; }
	}
}
