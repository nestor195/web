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
<b>Ingreso De Item a Una Orden</b><form method="POST" action="IngresoOrdenItem.asp" webbot-action="--WEBBOT-SELF--">
	<p>Orden: <select size="1" name="Orden">
	<option value="<%Response.Write Request.QueryString("Id")%>">
	<%Response.Write Request.QueryString("Id")%></option>
	</select><br>
	Item: <select size="1" name="Item">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL,ObConn
Equipo = ObRs ("Equipo")
ObRs.Close
ObConn.Close
%>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaEquipoItem Where Equipo = " & Equipo & " Order By Descripcion"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Item")%>"><%Response.Write ObRs ("Descripcion")%></option>
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
ObRs ("Orden") = Request.Form ("Orden")
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