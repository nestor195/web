<%@ Page Language="VB" %>
<html>
<head runat="server">
    <title>Filtering Data In A GridView Using a DropDownList</title>
</head>
<body>
    <form id="form1" runat="server">
        <b>Choose a state:</b>
        <asp:DropDownList id="DropDownList1" Runat="server" DataTextField="state" AutoPostBack="true" DataSourceID="SqlDataSource2"></asp:DropDownList>
        <asp:SqlDataSource id="SqlDataSource2" Runat="server" ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0; Ole DB Services=-4; Data Source=D:\Sistema\base\Copia de Sistema.mdb" SelectCommand="SELECT DISTINCT [Estado] FROM [Ordenes]"></asp:SqlDataSource>
        <br />
        <br />
        <asp:GridView id="GridView1" Runat="server" DataSourceID="SqlDataSource1" AutoGenerateColumns="False" DataKeyNames="au_id" AutoGenerateEditButton="True" AllowPaging="True" AllowSorting="True">
            <Columns>
                <asp:BoundField ReadOnly="true" HeaderText="ID" DataField="Id" SortExpression="Id" />
                <asp:BoundField HeaderText="Cliente" DataField="Cliente" SortExpression="Cliente" />
                <asp:BoundField HeaderText="FechaEstado" DataField="FechaEstado" SortExpression="FechaEstado" />
                <asp:CheckBoxField HeaderText="Contract" SortExpression="contract" DataField="contract" />
            </Columns>
        </asp:GridView>
        <asp:SqlDataSource id="SqlDataSource1" Runat="server" ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0; Ole DB Services=-4; Data Source=D:\Sistema\base\Copia de Sistema.mdb" SelectCommand="SELECT [ID], [Cliente], [Estado] FROM [Ordenes]" UpdateCommand="UPDATE [authors] SET [au_lname] = @au_lname, [au_fname] = @au_fname, [phone] = @phone, [address] = @address, [city] = @city, [state] = @state, [zip] = @zip, [contract] = @contract WHERE [au_id] = @au_id">
            <SelectParameters>
                <asp:ControlParameter Name="state" ControlID="DropDownList1" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="au_lname" />
                <asp:Parameter Name="au_fname" />
                <asp:Parameter Name="phone" />
                <asp:Parameter Name="address" />
                <asp:Parameter Name="city" />
                <asp:Parameter Name="state" />
                <asp:Parameter Name="zip" />
                <asp:Parameter Name="contract" />
                <asp:Parameter Name="au_id" />
            </UpdateParameters>
        </asp:SqlDataSource>
    </form>
</body>
</html>
