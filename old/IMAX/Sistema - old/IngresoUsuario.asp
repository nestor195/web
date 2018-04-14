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
<b>Ingreso de Usuario</b><form method="POST" action="IngresoUsuario.asp" webbot-action="--WEBBOT-SELF--">
	<p>Nick: <input type="text" name="Nick" size="30"><br>
	Password: <input type="text" name="Password" size="37"><br>
	Area: <select size="1" name="Area">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Areas",ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Area")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select><br>
	Cliente:<select size="1" name="Cliente">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Clientes",ObConn
DO WHILE NOT ObRs.Eof
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Nombre")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select><br>
	&nbsp;<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Usuarios",ObConn, 3, 3

ObRs.AddNew
ObRs ("Nick") = Request.Form ("Nick")
ObRs ("Password") = Request.Form ("Password")
ObRs ("Area") = Request.Form ("Area")
ObRs ("Cliente") = Request.Form ("Cliente")
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