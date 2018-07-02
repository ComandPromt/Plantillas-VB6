using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Web;
using System.Web.Services;
using System.Xml.Serialization;

namespace MortgageCalc
{
[WebService(Namespace="urn:mortgage-tools")]
public class MortgageService : WebService
{
	[WebMethod]
	public MortgagePayments CalculateMortgage(MortgageInfo minfo)
	{
		// TODO: calc mortgage and return MortgagePayments object
		return new MortgagePayments();
	}
	[WebMethod]
	public object GetSomething()
	{
		return "foo";
	}
	[WebMethod]
	public DataSet GetAuthors()
	{
		return new DataSet();
	}
	//[WebMethod]
	public Hashtable GetSomeHashTable()
	{
		return new Hashtable();
	}

}
public class MortgageInfo
{
	public double amount;
	public double years;
	public double interest;
	public double annualTax;
	public double annualInsurance;
}

public class MortgagePayments
{
	public double MonthlyPI;
	public double MonthlyTax;
	public double MonthlyInsurance;
	public double MonthlyTotal;
}

}
