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
SQL = "Select * From Usuarios Where Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn
%>

<b>Modificar Usuario</b><form method="POST" action="ModificarUsuario.asp" webbot-action="--WEBBOT-SELF--">
	<p>Id: <select size="1" name="Id">
    <option value="<%Response.Write ObRs("Id")%>" selected><%Response.Write ObRs("Id")%></option>
    </select><br>
	Nick: 
	<input type="text" name="Nick" size="30" value="<%Response.Write ObRs("Nick")%>"><br>	
	Password: 
	<input type="text" name="Password" size="37" value="<%Response.Write ObRs("Password")%>"><br>

	Area: <select size="1" name="Area">
<%
SET ObConn1 = Server.CreateObject ("ADODB.Connection")
SET ObRs1 = Server.CreateObject ("ADODB.RecordSet")
ObConn1.Open "Sistema"
SQL = "Select * From Areas Order by Area"
ObRs1.Open SQL, ObConn1
DO WHILE NOT ObRs1.Eof
If ObRs1("Id") = ObRs("area") then
%>
	<option selected value="<%Response.Write ObRs1 ("Id")%>"><%Response.Write ObRs1 ("Area")%></option>
<%
else
%>
	<option value="<%Response.Write ObRs1 ("Id")%>"><%Response.Write ObRs1 ("Area")%></option>
<%
End If
ObRs1.MoveNext
LOOP
%>
	</select><br>
<%
ObRs1.Close
ObConn1.Close
%>
	Cliente: <select size="1" name="Cliente">
<%
SET ObConn1 = Server.CreateObject ("ADODB.Connection")
SET ObRs1 = Server.CreateObject ("ADODB.RecordSet")
ObConn1.Open "Sistema"
SQL = "Select * From Clientes Order by Nombre"
ObRs1.Open SQL, ObConn1
DO WHILE NOT ObRs1.Eof
If ObRs1("Id") = ObRs("Cliente") then
%>
	<option selected value="<%Response.Write ObRs1 ("Id")%>"><%Response.Write ObRs1 ("Nombre")%></option>
<%
else
%>
	<option value="<%Response.Write ObRs1 ("Id")%>"><%Response.Write ObRs1 ("Nombre")%></option>
<%
End If
ObRs1.MoveNext
LOOP
%>
	</select><br>
<%
ObRs1.Close
ObConn1.Close
%>

	</select>Habilitado:
	&nbsp;<input type="checkbox" name="habilitado" value="true" <%if ObRs ("habilitado")= true then Response.Write "checked"%>><br>
	<input type="submit" value="Enviar" name="B1">
	</p>
</form>
<%
ObRs.Close
ObConn.Close
%>

<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios Where Id = " & Request.Form ("Id")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Nick") = Request.Form ("Nick")
ObRs ("Password") = Request.Form ("Password")
ObRs ("Area") = Request.Form ("Area")
ObRs ("Cliente") = Request.Form ("Cliente")
if Request.Form ("habilitado") = "" then
ObRs ("habilitado") = false
Else
ObRs ("habilitado") = true
End if
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