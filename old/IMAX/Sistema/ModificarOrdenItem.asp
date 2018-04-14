<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<%
IF Request.Form = "" THEN
%>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT OrdenItem.Id, OrdenItem.Orden, OrdenItem.Cantidad, Items.Descripcion, OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, Items.Id "
SQL = SQL & "FROM Items INNER JOIN OrdenItem ON Items.Id = OrdenItem.Item "
SQL = SQL & "Where OrdenItem.Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn
%>
<b>Modificar de Ítem de la Orden</b><form method="Post" action="ModificarOrdenItem.asp" webbot-action="--WEBBOT-SELF--">
	<p><select size="1" name="Id">
	<option selected value="<%Response.Write ObRs (0)%>"><%Response.Write ObRs (0)%></option>
	</select></p>
	<p>Orden:<select size="1" name="Orden">
	<option selected value="<%Response.Write ObRs ("Orden")%>"><%Response.Write ObRs ("Orden")%></option>
	</select></p>
	<p>Ítem:<%Response.Write ObRs ("Descripcion")%><a href="ModificarItem.asp?IdItem=<%Response.Write ObRs (6)%>"><img border="0" src="images/Editar.gif" width="32" height="29"></a></p>
	<p>Cantidad:<input type="text" name="Cantidad" size="12" value="<%Response.Write ObRs ("Cantidad")%>"></p>
	<p>Precio Unitario $<input type="text" name="PrecioUnitario" size="20" value="<%Response.Write ObRs ("PrecioUnitario")%>"></p>
	<p>Precio de Costo $<input type="text" name="PrecioCosto" size="20" value="<%Response.Write ObRs ("PrecioCosto")%>"></p>
	<p><input type="submit" value="Enviar" name="B1"><input type="reset" value="Restablecer" name="B2"></p>
</form>
<p>&nbsp;
<%
ObRs.Close
ObConn.Close
%>
</p>
<p>

<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From OrdenItem Where Id = " & Request.Form ("Id")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Cantidad") = Request.Form ("Cantidad")
ObRs ("PrecioUnitario") = Request.Form ("PrecioUnitario")
ObRs ("PrecioCosto") = Request.Form ("PrecioCosto")
ObRs.Update

ObRs.Close
ObConn.Close
Response.Redirect ("ConsultaDeOrden.asp?Id=" & Request.Form ("Orden"))
END IF
%>






</body>
</html>