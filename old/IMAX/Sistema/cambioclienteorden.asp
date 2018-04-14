<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Cambia Equipo Orden</title>
</head>
<%
If Request.form("Cliente") <> "" Then

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * from Ordenes where Id = " & Request.Form("Orden")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Cliente") = Request.form("Cliente")

ObRs.Update

ObRs.Close
ObConn.Close

Response.Redirect ("ConsultaDeOrden.asp?Id=" & Request.Form("Orden"))

End If
%>

<body>
<%
Orden = Request.QueryString("Orden")

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes where Id = " & Orden
ObRs.Open SQL, ObConn

Actual = ObRs("Cliente")

ObRs.Close
ObConn.Close


SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Clientes Order by Nombre"
ObRs.Open SQL, ObConn
%>

<form method="post" action="cambioclienteorden.asp">
<p>
<input type="text" name="Orden" size="20" value="<%Response.Write Orden%>"></p>
<p>
Cliente<select size="1" name="Cliente">
<%
DO WHILE NOT ObRs.Eof
IdEquipo = ObRs("Id")
Nombre = ObRs("Nombre")
If Actual = IdEquipo then
%>
<option selected value="<%Response.Write IdEquipo%>"><%Response.Write Nombre%></option>
<%
else
%>
<option value="<%Response.Write IdEquipo%>"><%Response.Write Nombre%></option>
<%

end if
ObRs.MoveNext
LOOP
%>
<%
ObRs.Close
ObConn.Close
%>
</select> <a href="IngresoCliente.asp">Nuevo</a></p>
	<p>
<input type="submit" value="Enviar" name="B1"></p>
</form>


</body>

</html>