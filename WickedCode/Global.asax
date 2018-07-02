<%@ Import NameSpace="System.Data" %>
<%@ Import NameSpace="System.Data.SqlClient" %>

<script language="C#" runat="server">
static Cache _cache = null;

void Application_Start ()
{
    _cache = Context.Cache; // Save reference for later

    //
    // Query the database and cache the resulting DataSet.
    //
    RefreshCache (null, null, 0);
}

static void RefreshCache (string key, object item,
    CacheItemRemovedReason reason)
{
    //
    // Query the database.
    //
    SqlDataAdapter adapter = new SqlDataAdapter (
        "SELECT * FROM Quotations",
        "server=localhost;database=quotes;uid=sa;pwd="
    );

    DataSet ds = new DataSet ();
    adapter.Fill (ds, "Quotations");

    //
    // Add the DataSet to the application cache.
    //
    _cache.Insert (
        "Quotes",
        ds,
        new CacheDependency ("C:\\AspNetSql\\Quotes.Quotations"),
        Cache.NoAbsoluteExpiration,
        Cache.NoSlidingExpiration,
        CacheItemPriority.Default,
        new CacheItemRemovedCallback (RefreshCache)
    );
}
</script>