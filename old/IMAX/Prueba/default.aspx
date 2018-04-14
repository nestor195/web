<%@ Page Language="VB" %>
<%@ Register TagPrefix="wmx" Namespace="Microsoft.Matrix.Framework.Web.UI" Assembly="Microsoft.Matrix.Framework, Version=0.6.0.0, Culture=neutral, PublicKeyToken=6f763c9966660626" %>
<script runat="server">

    Sub Button1_Click(sender As Object, e As EventArgs)
        AccessDataSourceControl1.SelectCommand="SELECT * FROM [Estados]"
    End Sub

</script>
<html>
<head runat="server">
    <title>GridView Bound to SqlDataSource</title>
</head>
<body>
    <form id="form1" runat="server">
        <p>
            <wmx:AccessDataSourceControl id="AccessDataSourceControl1" runat="server" ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0; Ole DB Services=-4; Data Source=D:\Sistema\base\Copia de Sistema.mdb" SelectCommand="SELECT * FROM [Clientes]"></wmx:AccessDataSourceControl>
            <wmx:MxDataGrid id="MxDataGrid1" runat="server" BorderStyle="None" BorderWidth="1px" BorderColor="#CCCCCC" BackColor="White" DataMember="Clientes" DataSourceControlID="AccessDataSourceControl1" DataKeyField="Id" CellPadding="3" AllowSorting="True" AllowPaging="True">
                <FooterStyle backcolor="White" forecolor="#000066"></FooterStyle>
                <HeaderStyle backcolor="#006699" font-bold="True" forecolor="White"></HeaderStyle>
                <ItemStyle forecolor="#000066"></ItemStyle>
                <PagerStyle mode="NumericPages" horizontalalign="Center" backcolor="White" forecolor="#000066"></PagerStyle>
                <SelectedItemStyle backcolor="#669999" font-bold="True" forecolor="White"></SelectedItemStyle>
            </wmx:MxDataGrid>
        </p>
        <p>
            <asp:Button id="Button1" onclick="Button1_Click" runat="server" Text="Button"></asp:Button>
        </p>
    </form>
</body>
</html>
