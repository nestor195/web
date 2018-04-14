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
<b>Eliminar Item de la Orden</b><form method="POST" action="EliminarVentaItem.asp" webbot-action="--WEBBOT-SELF--">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Id = " & Request.QueryString("Id")
ObRs.Open SQL,ObConn
%>
Eliminar
<table border="1" width="47%" id="table1">
	<tr>
		<td width="66"><%Response.Write ObRs ("Codigo")%>&nbsp;</td>
		<td width="112"><%Response.Write ObRs ("Descripcion")%>&nbsp;</td>
		<td width="83"><%Response.Write ObRs ("Cantidad")%>&nbsp;</td>
		<td><%Response.Write ObRs ("PrecioUnitario")%>&nbsp;</td>
	</tr>
</table>
de la <span lang="es">del carrito de Venta</span><br>
<select size="1" name="Id">
<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></option>
</select><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
<%
ObRs.Close
ObConn.Close
%>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From OrdenItem Where Id = " & Request.Form("Id")
ObRs.Open SQL,ObConn, 3, 3

ObRs.Delete

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
END IF
%>

</body>
</html>