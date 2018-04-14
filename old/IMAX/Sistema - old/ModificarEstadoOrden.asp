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

<b>Modificar Estado de la Orden</b><form method="POST" action="ModificarEstadoOrden.asp" webbot-action="--WEBBOT-SELF--">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL,ObConn
%>
	<p><b>Orden: <select size="1" name="Id">
	<option selected value="<%Response.Write ObRs ("Id")%>">
	<%Response.Write ObRs ("Id")%></option>
	</select></b></p>
<%
UsuarioEstado = ObRs ("UsuarioEstado")
Estado = ObRs ("Estado")
ObRs.Close
ObConn.Close
%>
<p><b>Técnico: <select size="1" name="UsuarioEstado">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios  Where habilitado = true"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
IF UsuarioEstado = ObRs ("Id") Then
%>
<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Nick")%></option>
<%
ELSE
%>
<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Nick")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
</select></b><Br>
<b>Estado: <select size="1" name="Estado">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Estados",ObConn
DO WHILE NOT ObRs.Eof
IF Estado = ObRs ("Id") Then
%>
<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Estado")%></option>
<%
Else
%>
<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Estado")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>

</select></b><Br>
<span lang="es"><b>Fecha:</b> 
<input type="text" name="FechaEstado" size="13" value="<%Response.Write Date%>"><Br>

</span><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.Form("Id")
ObRs.Open SQL, ObConn, 3, 3

ObRs ("Estado") = Request.Form ("Estado")
ObRs ("UsuarioEstado") = Request.Form ("UsuarioEstado")
ObRs ("FechaEstado") = Request.Form ("FechaEstado")
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