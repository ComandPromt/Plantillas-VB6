using System;
using localhost;

namespace Client
{
	class Test
	{
		static void Main(string[] args)
		{
            Arithmetic proxy = new Arithmetic();
            proxy.Url = "http://localhost/server/arithmetic.asmx";
            Add add = new Add();
            add.n1 = 10;
            add.n2 = 20;
            AddResponse resp = proxy.Add(add);
            Console.WriteLine("10 + 20 = " + resp.sum);
		}
	}
}
