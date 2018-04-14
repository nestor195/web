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
<b>Ingreso De Ítem a Una <span lang="es">Venta</span></b><form method="POST" action="IngresoVentaItem.asp" webbot-action="--WEBBOT-SELF--">
	<p>Ítem: <select size="1" name="Item">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Items Where Venta = 1 Order By Descripcion"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Descripcion")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select> <a target="_parent" href="IngresoItem.asp">Nuevo</a><Br>
	Precio: <input type="text" name="PrecioUnitario" size="8"><Br>
	Cantidad: <input type="text" name="Cantidad" size="3" value="1"><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "OrdenItem",ObConn, 3, 3

ObRs.AddNew
ObRs ("Orden") = 1
ObRs ("Item") = Request.Form ("Item")
ObRs ("PrecioUnitario") = Request.Form ("PrecioUnitario")
ObRs ("Cantidad") = Request.Form ("Cantidad")
ObRs.Update

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
END IF
%>

</body>
</html>