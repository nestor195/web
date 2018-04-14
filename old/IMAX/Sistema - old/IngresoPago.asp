<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<SCRIPT language=javascript type="text/Jscript">
function ingresocliente(){
newWindow = window.open('IngresoCliente.asp','IngresoCliente','width=500,height=300')
}
function ingresoequipo(){
newWindow = window.open('IngresoEquipo.asp','IngresoEquipo','width=250,height=180')
}
</SCRIPT>
</head>

<body>
<%
IF Request.Form = "" THEN
%>
<b>Ingreso de Orden<span lang="es"> de Pago</span></b><form method="POST" action="IngresoPago.asp">
	<p>Cliente: <select size="1" name="Cliente">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Clientes Order By Nombre"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
IF Request.QueryString("Cliente") = ObRs ("Nombre") THEN
%>
	<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Nombre")%></option>
<%
ELSE
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Nombre")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select> <a href="javascript:ingresocliente()">Nuevo</a><br>
	Equipo: <select size="1" name="Equipo">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Equipos Where Modelo ='pago'"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Modelo")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select> <a href="javascript:ingresoequipo()">Nuevo</a><br>
	Usuario de Ingreso: <select size="1" name="UsuarioIngreso">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Usuarios",ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Nick")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select><br>
	Observaciones de Ingreso:<br>
	<textarea rows="3" name="ObservacionIngreso" cols="39"></textarea><br>
	</p>
	<p>Ítem: <select size="1" name="Item">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Items Where Id = 183"
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

	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Ordenes",ObConn, 3, 3

ObRs.AddNew
ObRs ("Cliente") = Request.Form ("Cliente")
ObRs ("Equipo") = Request.Form ("Equipo")
ObRs ("Serie") = UCase(Request.Form ("Serie"))
ObRs ("Estado") = 20
ObRs ("Accesorios") = Request.Form ("Accesorios")
ObRs ("UsuarioIngreso") = Request.Form ("UsuarioIngreso")
ObRs ("UsuarioEstado") = Request.Form ("UsuarioIngreso")
ObRs ("FechaIngreso") = DATE
ObRs ("FechaEstado") = DATE
ObRs ("ObservacionIngreso") = Request.Form ("ObservacionIngreso")
ObRs.Update
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Order By Id"
ObRs.Open SQL, ObConn
DO WHILE NOT ObRs.Eof
Orden = ObRs("Id")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "OrdenItem",ObConn, 3, 3

ObRs.AddNew
ObRs ("Orden") = orden
ObRs ("Item") = Request.Form ("Item")
ObRs ("PrecioUnitario") = Request.Form ("PrecioUnitario") * -1
ObRs ("Cantidad") = 1
ObRs.Update

ObRs.Close
ObConn.Close

%>
<b>Datos Ingresados</b><p><b>
<a target="_blank" href="fpdf/Orden.asp?Id=<%Response.Write orden%>">Imprimir</a></b>
</p>
<%

END IF
%>

</body>
</html>