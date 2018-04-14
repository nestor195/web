<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="es-ar" http-equiv="Content-Language" />
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Usuarios</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
</head>

<body>

<p>Crear Nuevo</p>
<table cellpadding="2" cellspacing="0" class="tablas" style="width: 100%">
	<tr class="celdaazul">
		<td class="celdaazul">Responsabre</td>
		<td class="celdaazul">Inhabilitado</td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "ACYP"
SQL = "SELECT Id, Responsable, Inhabilitado FROM Responsables"

Select Case Request.Querystring ("Campo")
Case "Id"
If IsNumeric(Request.QueryString ("Dato")) Then
SQL = SQL & " Where Planilla.Id = " & Request.QueryString ("Dato")
End If
Case "Estado"
SQL = SQL & " Where Estados.Estado Like '%" & Request.QueryString ("Dato") & "%'"
End Select

SQL = SQL & " Order By Responsable ASC"

ObRs.Open SQL,ObConn
Do While ObRs.EOF = false
%>

	<tr class="celdablanca">
		<td  class="celdablanca">
		<a href="consultaresponsable.asp?AC=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></a></td>
<%
Inhabilitado = ""
If ObRs("Inhabilitado") = True Then
Inhabilitado = "X"
End If
%>
		<td  class="celdablanca"><%Response.Write Inhabilitado%></td>
	</tr>
<%
ObRs.MoveNext
Loop
ObRs.Close
ObConn.Close
%>

</table>

</body>

</html>
