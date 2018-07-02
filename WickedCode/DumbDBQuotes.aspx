<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>

<html>
  <body>
    <h3><asp:Label ID="Quotation" RunAt="server" /></h3>
    <i><asp:Label ID="Author" RunAt="server" /></i>
  </body>
</html>

<script language="C#" runat="server">
void Page_Load (Object sender, EventArgs e)
{
    SqlDataAdapter adapter = new SqlDataAdapter (
        "SELECT * FROM Quotations",
        "server=localhost;database=quotes;uid=sa;pwd="
    );

    DataSet ds = new DataSet ();
    adapter.Fill (ds, "Quotations");
    DataTable table = ds.Tables["Quotations"];

    Random rand = new Random ();
    int index = rand.Next (0, table.Rows.Count);
    DataRow row = table.Rows[index];

    Quotation.Text = (string) row["Quotation"];
    Author.Text = (string) row["Author"];
}
</script>