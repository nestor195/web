<!DOCTYPE html PUBLIC
          "-//W3C//DTD XHTML 1.0 Transitional//EN"
          "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Cantidad de equipos ingresados</title>
    <script src="src/js/jscal2.js"></script>
    <script src="src/js/lang/es.js"></script>
    <link rel="stylesheet" type="text/css" href="src/css/jscal2.css" />
    <link rel="stylesheet" type="text/css" href="src/css/border-radius.css" />
    <link rel="stylesheet" type="text/css" href="src/css/steel/steel.css" />
</head>

<body>

<p><font face="Arial"><b>Cantidad de equipos ingresados</b></font></p>
<form method="POST" action="Cantidad_de_equipos_ingresados.asp" webbot-action="--WEBBOT-SELF--">
	<!--webbot bot="SaveResults" U-File="_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" startspan --><input NAME="VTI-GROUP" TYPE="hidden" VALUE="0"><!--webbot bot="SaveResults" i-checksum="37496" endspan -->
	<p>Desde:
    <input size="11" id="f_date1" name="desde" /><button id="f_btn1">...</button><br />
    Hasta:&nbsp;
    <input size="11" id="f_date2" name="hasta" /><button id="f_btn2">...</button>
	</p>
	<p><input type="submit" value="Enviar" name="B1"></p>
</form>
    <script type="text/javascript">//<![CDATA[

      var cal = Calendar.setup({
          onSelect: function(cal) { cal.hide() }
      });
      cal.manageFields("f_btn1", "f_date1", "%Y/%m/%d");
      cal.manageFields("f_btn2", "f_date2", "%Y/%m/%d");
    //]]></script>
<%
Response.Write "desde " & Request.Form("desde") & " hasta "
Response.Write Request.Form("hasta")
%>
<table border="1" width="100%">
	<tr>
		<td>&nbsp;</td>
		<td>Ingresados</td>
		<td>Presupuestados</td>
		<td>Confirmados</td>
		<td>Listas</td>
		<td>Entregadas</td>
	</tr>
	<tr>
		<td>PC</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 3 OR Equipos.Tipo = 10) "
SQL = SQL & "AND Ordenes.Estado <> 10"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 3 OR Equipos.Tipo = 10) "
SQL = SQL & "AND Ordenes.Estado = 2"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 3 OR Equipos.Tipo = 10) "
SQL = SQL & "AND Ordenes.Estado = 3"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 3 OR Equipos.Tipo = 10) "
SQL = SQL & "AND Ordenes.Estado = 4"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 3 OR Equipos.Tipo = 10) "
SQL = SQL & "AND (Ordenes.Estado = 5 OR Ordenes.Estado = 9 OR Ordenes.Estado = 27 OR Ordenes.Estado = 12 OR Ordenes.Estado = 17 OR Ordenes.Estado = 19 OR Ordenes.Estado = 11)"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
	</tr>
	<tr>
		<td>Notebook</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 7 OR Equipos.Tipo = 32 OR Equipos.Tipo = 42 OR Equipos.Tipo = 47) "
SQL = SQL & "AND Ordenes.Estado <> 10"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 7 OR Equipos.Tipo = 32 OR Equipos.Tipo = 42 OR Equipos.Tipo = 47) "
SQL = SQL & "AND Ordenes.Estado = 2"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 7 OR Equipos.Tipo = 32 OR Equipos.Tipo = 42 OR Equipos.Tipo = 47) "
SQL = SQL & "AND Ordenes.Estado = 3"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 7 OR Equipos.Tipo = 32 OR Equipos.Tipo = 42 OR Equipos.Tipo = 47) "
SQL = SQL & "AND Ordenes.Estado = 4"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 7 OR Equipos.Tipo = 32 OR Equipos.Tipo = 42 OR Equipos.Tipo = 47) "
SQL = SQL & "AND (Ordenes.Estado = 5 OR Ordenes.Estado = 9 OR Ordenes.Estado = 27 OR Ordenes.Estado = 12 OR Ordenes.Estado = 17 OR Ordenes.Estado = 19 OR Ordenes.Estado = 11)"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
	</tr>
	<tr>
		<td>Impresoras</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 1 OR Equipos.Tipo = 2 OR Equipos.Tipo = 4 OR Equipos.Tipo = 5 OR Equipos.Tipo = 6 OR Equipos.Tipo = 13 OR Equipos.Tipo = 31 OR Equipos.Tipo = 33 OR Equipos.Tipo = 44 OR Equipos.Tipo = 54) "
SQL = SQL & "AND Ordenes.Estado <> 10 "
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 1 OR Equipos.Tipo = 2 OR Equipos.Tipo = 4 OR Equipos.Tipo = 5 OR Equipos.Tipo = 6 OR Equipos.Tipo = 13 OR Equipos.Tipo = 31 OR Equipos.Tipo = 33 OR Equipos.Tipo = 44 OR Equipos.Tipo = 54) "
SQL = SQL & "AND Ordenes.Estado = 2"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 1 OR Equipos.Tipo = 2 OR Equipos.Tipo = 4 OR Equipos.Tipo = 5 OR Equipos.Tipo = 6 OR Equipos.Tipo = 13 OR Equipos.Tipo = 31 OR Equipos.Tipo = 33 OR Equipos.Tipo = 44 OR Equipos.Tipo = 54) "
SQL = SQL & "AND Ordenes.Estado = 3"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 1 OR Equipos.Tipo = 2 OR Equipos.Tipo = 4 OR Equipos.Tipo = 5 OR Equipos.Tipo = 6 OR Equipos.Tipo = 13 OR Equipos.Tipo = 31 OR Equipos.Tipo = 33 OR Equipos.Tipo = 44 OR Equipos.Tipo = 54) "
SQL = SQL & "AND (Ordenes.Estado = 4)"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 1 OR Equipos.Tipo = 2 OR Equipos.Tipo = 4 OR Equipos.Tipo = 5 OR Equipos.Tipo = 6 OR Equipos.Tipo = 13 OR Equipos.Tipo = 31 OR Equipos.Tipo = 33 OR Equipos.Tipo = 44 OR Equipos.Tipo = 54) "
SQL = SQL & "AND (Ordenes.Estado = 5 OR Ordenes.Estado = 9 OR Ordenes.Estado = 27 OR Ordenes.Estado = 12 OR Ordenes.Estado = 17 OR Ordenes.Estado = 19 OR Ordenes.Estado = 11)"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%>
</td>
	</tr>
	<tr>
		<td>Toners</td>
		<td>&nbsp;
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 9 OR Equipos.Tipo = 29 OR Equipos.Tipo = 27 OR Equipos.Tipo = 30) "
SQL = SQL & "AND Ordenes.Estado <> 10"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 9 OR Equipos.Tipo = 29 OR Equipos.Tipo = 27 OR Equipos.Tipo = 30) "
SQL = SQL & "AND Ordenes.Estado = 2"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 9 OR Equipos.Tipo = 29 OR Equipos.Tipo = 27 OR Equipos.Tipo = 30) "
SQL = SQL & "AND Ordenes.Estado = 3"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 9 OR Equipos.Tipo = 29 OR Equipos.Tipo = 27 OR Equipos.Tipo = 30) "
SQL = SQL & "AND Ordenes.Estado = 4"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 9 OR Equipos.Tipo = 29 OR Equipos.Tipo = 27 OR Equipos.Tipo = 30) "
SQL = SQL & "AND (Ordenes.Estado = 5 OR Ordenes.Estado = 9 OR Ordenes.Estado = 27 OR Ordenes.Estado = 12 OR Ordenes.Estado = 17 OR Ordenes.Estado = 19 OR Ordenes.Estado = 11)"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
	</tr>
	<tr>
		<td>Ploters</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 46) "
SQL = SQL & "AND Ordenes.Estado <> 10"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 46) "
SQL = SQL & "AND Ordenes.Estado = 2"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 46) "
SQL = SQL & "AND Ordenes.Estado = 3"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 46) "
SQL = SQL & "AND Ordenes.Estado = 4"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 46) "
SQL = SQL & "AND (Ordenes.Estado = 5 OR Ordenes.Estado = 9 OR Ordenes.Estado = 27 OR Ordenes.Estado = 12 OR Ordenes.Estado = 17 OR Ordenes.Estado = 19 OR Ordenes.Estado = 11)"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
	</tr>
	<tr>
		<td>Otros</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 14 OR Equipos.Tipo = 15 OR Equipos.Tipo = 16 OR Equipos.Tipo = 17 OR Equipos.Tipo = 18 OR Equipos.Tipo = 24 OR Equipos.Tipo = 25 OR Equipos.Tipo = 26 OR Equipos.Tipo = 28 OR Equipos.Tipo = 34 OR Equipos.Tipo = 35 OR Equipos.Tipo = 36 OR Equipos.Tipo = 37 OR Equipos.Tipo = 38 OR Equipos.Tipo = 39 OR Equipos.Tipo = 40 OR Equipos.Tipo = 41 OR Equipos.Tipo = 43 OR Equipos.Tipo = 45 OR Equipos.Tipo = 48 OR Equipos.Tipo = 49 OR Equipos.Tipo = 50 OR Equipos.Tipo = 51 OR Equipos.Tipo = 52 OR Equipos.Tipo = 53 OR Equipos.Tipo = 55 OR Equipos.Tipo = 56 OR Equipos.Tipo = 57) "
SQL = SQL & "AND Ordenes.Estado <> 10"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 14 OR Equipos.Tipo = 15 OR Equipos.Tipo = 16 OR Equipos.Tipo = 17 OR Equipos.Tipo = 18 OR Equipos.Tipo = 24 OR Equipos.Tipo = 25 OR Equipos.Tipo = 26 OR Equipos.Tipo = 28 OR Equipos.Tipo = 34 OR Equipos.Tipo = 35 OR Equipos.Tipo = 36 OR Equipos.Tipo = 37 OR Equipos.Tipo = 38 OR Equipos.Tipo = 39 OR Equipos.Tipo = 40 OR Equipos.Tipo = 41 OR Equipos.Tipo = 43 OR Equipos.Tipo = 45 OR Equipos.Tipo = 48 OR Equipos.Tipo = 49 OR Equipos.Tipo = 50 OR Equipos.Tipo = 51 OR Equipos.Tipo = 52 OR Equipos.Tipo = 53 OR Equipos.Tipo = 55 OR Equipos.Tipo = 56 OR Equipos.Tipo = 57) "
SQL = SQL & "AND Ordenes.Estado = 2"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 14 OR Equipos.Tipo = 15 OR Equipos.Tipo = 16 OR Equipos.Tipo = 17 OR Equipos.Tipo = 18 OR Equipos.Tipo = 24 OR Equipos.Tipo = 25 OR Equipos.Tipo = 26 OR Equipos.Tipo = 28 OR Equipos.Tipo = 34 OR Equipos.Tipo = 35 OR Equipos.Tipo = 36 OR Equipos.Tipo = 37 OR Equipos.Tipo = 38 OR Equipos.Tipo = 39 OR Equipos.Tipo = 40 OR Equipos.Tipo = 41 OR Equipos.Tipo = 43 OR Equipos.Tipo = 45 OR Equipos.Tipo = 48 OR Equipos.Tipo = 49 OR Equipos.Tipo = 50 OR Equipos.Tipo = 51 OR Equipos.Tipo = 52 OR Equipos.Tipo = 53 OR Equipos.Tipo = 55 OR Equipos.Tipo = 56 OR Equipos.Tipo = 57) "
SQL = SQL & "AND Ordenes.Estado = 3"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 14 OR Equipos.Tipo = 15 OR Equipos.Tipo = 16 OR Equipos.Tipo = 17 OR Equipos.Tipo = 18 OR Equipos.Tipo = 24 OR Equipos.Tipo = 25 OR Equipos.Tipo = 26 OR Equipos.Tipo = 28 OR Equipos.Tipo = 34 OR Equipos.Tipo = 35 OR Equipos.Tipo = 36 OR Equipos.Tipo = 37 OR Equipos.Tipo = 38 OR Equipos.Tipo = 39 OR Equipos.Tipo = 40 OR Equipos.Tipo = 41 OR Equipos.Tipo = 43 OR Equipos.Tipo = 45 OR Equipos.Tipo = 48 OR Equipos.Tipo = 49 OR Equipos.Tipo = 50 OR Equipos.Tipo = 51 OR Equipos.Tipo = 52 OR Equipos.Tipo = 53 OR Equipos.Tipo = 55 OR Equipos.Tipo = 56 OR Equipos.Tipo = 57) "
SQL = SQL & "AND Ordenes.Estado = 4"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 14 OR Equipos.Tipo = 15 OR Equipos.Tipo = 16 OR Equipos.Tipo = 17 OR Equipos.Tipo = 18 OR Equipos.Tipo = 24 OR Equipos.Tipo = 25 OR Equipos.Tipo = 26 OR Equipos.Tipo = 28 OR Equipos.Tipo = 34 OR Equipos.Tipo = 35 OR Equipos.Tipo = 36 OR Equipos.Tipo = 37 OR Equipos.Tipo = 38 OR Equipos.Tipo = 39 OR Equipos.Tipo = 40 OR Equipos.Tipo = 41 OR Equipos.Tipo = 43 OR Equipos.Tipo = 45 OR Equipos.Tipo = 48 OR Equipos.Tipo = 49 OR Equipos.Tipo = 50 OR Equipos.Tipo = 51 OR Equipos.Tipo = 52 OR Equipos.Tipo = 53 OR Equipos.Tipo = 55 OR Equipos.Tipo = 56 OR Equipos.Tipo = 57) "
SQL = SQL & "AND (Ordenes.Estado = 5 OR Ordenes.Estado = 9 OR Ordenes.Estado = 27 OR Ordenes.Estado = 12 OR Ordenes.Estado = 17 OR Ordenes.Estado = 19 OR Ordenes.Estado = 11)"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
	</tr>
	<tr>
		<td>Totales</td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "#"
SQL = SQL & "AND (Equipos.Tipo <> 8 AND Equipos.Tipo <> 19 AND Equipos.Tipo <> 20 AND Equipos.Tipo <> 21 AND Equipos.Tipo <> 22 AND Equipos.Tipo <> 23) "
SQL = SQL & "AND Ordenes.Estado <> 10"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "#"
SQL = SQL & "AND (Equipos.Tipo <> 8 AND Equipos.Tipo <> 19 AND Equipos.Tipo <> 20 AND Equipos.Tipo <> 21 AND Equipos.Tipo <> 22 AND Equipos.Tipo <> 23) "
SQL = SQL & "AND Ordenes.Estado = 2"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "#"
SQL = SQL & "AND (Equipos.Tipo <> 8 AND Equipos.Tipo <> 19 AND Equipos.Tipo <> 20 AND Equipos.Tipo <> 21 AND Equipos.Tipo <> 22 AND Equipos.Tipo <> 23) "
SQL = SQL & "AND Ordenes.Estado = 3"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "#"
SQL = SQL & "AND (Equipos.Tipo <> 8 AND Equipos.Tipo <> 19 AND Equipos.Tipo <> 20 AND Equipos.Tipo <> 21 AND Equipos.Tipo <> 22 AND Equipos.Tipo <> 23) "
SQL = SQL & "AND Ordenes.Estado = 4"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
		<td>&nbsp;<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.FechaIngreso, Ordenes.FechaEstado, Ordenes.Estado, "
SQL = SQL & "Equipos.Tipo FROM Equipos INNER JOIN Ordenes ON "
SQL = SQL & "Equipos.Id = Ordenes.Equipo "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaEstado >= #" & Request.Form("desde") & "# AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "#"
SQL = SQL & "AND (Equipos.Tipo <> 8 AND Equipos.Tipo <> 19 AND Equipos.Tipo <> 20 AND Equipos.Tipo <> 21 AND Equipos.Tipo <> 22 AND Equipos.Tipo <> 23) "
SQL = SQL & "AND (Ordenes.Estado = 5 OR Ordenes.Estado = 9 OR Ordenes.Estado = 27 OR Ordenes.Estado = 12 OR Ordenes.Estado = 17 OR Ordenes.Estado = 19 OR Ordenes.Estado = 11)"
end if
ObRs.CursorType=1
ObRs.Open SQL,ObConn
total=ObRs.RecordCount
Response.Write total
%></td>
	</tr>
</table>

</body>

</html>