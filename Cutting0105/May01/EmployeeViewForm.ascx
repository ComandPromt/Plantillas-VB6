<!-- No html, body, form, or style tags allowed -->

<% @ Control Language="C#" %>

<script runat="server">
    public void ClearAll()
    {
	lblEmployeeID.Text = "";
	txtTitleOfCourtesy.Text = "";
	txtFirstName.Text = "";
	txtLastName.Text = "";
	txtTitle.Text = "";
	txtHireDate.Text = "";
	txtNotes.Text = "";
    }

    public String EmployeeID {
        get { return lblEmployeeID.Text; }	
        set { lblEmployeeID.Text = value; }
    }	

    public String TitleOfCourtesy {
        get { return txtTitleOfCourtesy.Text; }	
        set { txtTitleOfCourtesy.Text = value; }
    }	

    public String Title {
        get { return txtTitle.Text; }	
        set { txtTitle.Text = value; }
    }

    public String FirstName {
        get { return txtFirstName.Text; }	
        set { txtFirstName.Text = value; }
    }	

    public String LastName {
        get { return txtLastName.Text; }	
        set { txtLastName.Text = value; }
    }	

    public String HireDate {
        get { return txtHireDate.Text; }	
        set { txtHireDate.Text = value; }
    }	

    public String Notes {
        get { return txtNotes.Text; }	
        set { txtNotes.Text = value; }
    }	
</script>

<table style="border:solid 1px black;" bgcolor="lightyellow">
   <tr>
     <td><asp:label runat="server" text="First Name" Font-Bold="true" /></td>
     <td><asp:textbox runat="server" ID="lblEmployeeID" readonly="true" Font-Bold="true" 
		style="font-family:verdana;font-size:x-small;background-color=beige;border:solid 1px black;width=40px;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');"/>

         <asp:textbox runat="server" id="txtTitleOfCourtesy" 
		style="font-family:verdana;font-size:x-small;border:solid 1px black;width=50px;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');" />
         <asp:textbox runat="server" id="txtFirstName" 
		style="font-family:verdana;font-size:x-small;border:solid 1px black;width=200px;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');" /></td></tr>
   <tr>
     <td><asp:label runat="server" text="Last Name" Font-Bold="true" /></td>
     <td><asp:textbox runat="server" id="txtLastName" 
		style="font-family:verdana;font-size:x-small;border:solid 1px black;width=150px;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');" />
	 <asp:textbox runat="server" id="txtHireDate" 
		style="font-family:verdana;font-size:x-small;border:solid 1px black;width=145px;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');" /></td></tr>
   <tr>
     <td><asp:label runat="server" text="Title" Font-Bold="true" /></td>
     <td><asp:textbox runat="server" id="txtTitle" 
		style="font-family:verdana;font-size:x-small;border:solid 1px black;width=300px;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');" /></td></tr>
  <tr>
     <td><asp:label runat="server" text="Notes" Font-Bold="true" /></td>
     <td><asp:textbox runat="server" id="txtNotes" 
		textmode="multiline" rows="4"
		style="font-family:verdana;font-size:xx-small;border:solid 1px black;width=300px;filter:progid:DXImageTransform.Microsoft.dropshadow(OffX=2, OffY=2, Color='gray', Positive='true');" /></td></tr>

</table>

