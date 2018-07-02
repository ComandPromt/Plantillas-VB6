<%@ Import Namespace="System.Data" %>

<html>
  <body>
    <h3><asp:Label ID="Quotation" RunAt="server" /></h3>
    <i><asp:Label ID="Author" RunAt="server" /></i>
  </body>
</html>

<script language="C#" runat="server">
void Page_Load (Object sender, EventArgs e)
{
    DataSet ds = (DataSet) Cache["Quotes"];

    if (ds != null) {
        // Display a randomly selected quotation
        DataTable table = ds.Tables["Quotations"];

        Random rand = new Random ();
        int index = rand.Next (0, table.Rows.Count);
        DataRow row = table.Rows[index];

        Quotation.Text = (string) row["Quotation"];
        Author.Text = (string) row["Author"];
    }
    else {
        // If quotes is null, this request arrived after the
        // DataSet was removed from the cache and before a new
        // DataSet was inserted. Tell the user the server is
        // busy; a page refresh should solve the problem.
        Quotation.Text = "Server busy";
    }
}
</script>