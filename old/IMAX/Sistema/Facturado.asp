<!DOCTYPE html PUBLIC
          "-//W3C//DTD XHTML 1.0 Transitional//EN"
          "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Facturado</title>
    <script src="src/js/jscal2.js"></script>
    <script src="src/js/lang/es.js"></script>
    <link rel="stylesheet" type="text/css" href="src/css/jscal2.css" />
    <link rel="stylesheet" type="text/css" href="src/css/border-radius.css" />
    <link rel="stylesheet" type="text/css" href="src/css/steel/steel.css" />
</head>

<body>

<p><b><font face="Arial">Facturado</font></b></p>
<form method="POST" action="Facturado.asp" webbot-action="--WEBBOT-SELF--">
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
		<td width="81">&nbsp;</td>
		<td>Total</td>
		<td>Costo</td>
		<td>Ganancia</td>
		<td>%</td>
	</tr>
	<tr>
		<td width="81">PC</td>
		<td>&nbsp;$<%
total = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 3 OR Equipos.Tipo = 10) "
SQL = SQL & "AND (Ordenes.Estado = 6 OR Ordenes.Estado = 9)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total = total + ObRs ("Cantidad") * ObRs ("PrecioUnitario") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total,2)
%>
</td>
		<td>&nbsp;$<%
total1 = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 3 OR Equipos.Tipo = 10) "
SQL = SQL & "AND (Ordenes.Estado = 6 OR Ordenes.Estado = 9)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total1 = total1 + ObRs ("Cantidad") * ObRs ("PrecioCosto") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total1,2)
%>
</td>
		<td>&nbsp;$<%
Response.Write formatnumber(total - total1,2)
		%></td>
		<td>&nbsp;<%
if total > 0 then
Response.Write formatnumber(100-(total1/total)*100,2)
end if
		%>%</td>
	</tr>
	<tr>
		<td width="81">Notebook</td>
		<td>&nbsp;$<%
total = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 7 OR Equipos.Tipo = 32 OR Equipos.Tipo = 42 OR Equipos.Tipo = 47) "
SQL = SQL & "AND (Ordenes.Estado = 6 OR Ordenes.Estado = 9)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total = total + ObRs ("Cantidad") * ObRs ("PrecioUnitario") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total,2)
%>
</td>
		<td>&nbsp;$<%
total1 = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 7 OR Equipos.Tipo = 32 OR Equipos.Tipo = 42 OR Equipos.Tipo = 47) "
SQL = SQL & "AND (Ordenes.Estado = 6 OR Ordenes.Estado = 9)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total1 = total1 + ObRs ("Cantidad") * ObRs ("PrecioCosto") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total1,2)
%>
</td>
		<td>&nbsp;$<%
Response.Write formatnumber(total - total1,2)
		%></td>
		<td>&nbsp;<%
if total > 0 then
Response.Write formatnumber(100-(total1/total)*100,2)
end if
		%>%</td>
	</tr>
	<tr>
		<td width="81">Impresoras</td>
		<td>&nbsp;$<%
total = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 1 OR Equipos.Tipo = 2 OR Equipos.Tipo = 4 OR Equipos.Tipo = 5 OR Equipos.Tipo = 6 OR Equipos.Tipo = 13 OR Equipos.Tipo = 31 OR Equipos.Tipo = 33 OR Equipos.Tipo = 44 OR Equipos.Tipo = 54) "
SQL = SQL & "AND (Ordenes.Estado = 6 OR Ordenes.Estado = 9)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total = total + ObRs ("Cantidad") * ObRs ("PrecioUnitario") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total,2)
%>
</td>
		<td>&nbsp;$<%
total1 = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Equipos.Tipo = 1 OR Equipos.Tipo = 2 OR Equipos.Tipo = 4 OR Equipos.Tipo = 5 OR Equipos.Tipo = 6 OR Equipos.Tipo = 13 OR Equipos.Tipo = 31 OR Equipos.Tipo = 33 OR Equipos.Tipo = 44 OR Equipos.Tipo = 54) "
SQL = SQL & "AND (Ordenes.Estado = 6 OR Ordenes.Estado = 9)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total1 = total1 + ObRs ("Cantidad") * ObRs ("PrecioCosto") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total1,2)
%>
</td>
		<td>&nbsp;$<%
Response.Write formatnumber(total - total1,2)
		%></td>
		<td>&nbsp;<%
if total > 0 then
Response.Write formatnumber(100-(total1/total)*100,2)
end if
		%>%</td>
	</tr>
	<tr>
		<td width="81">Venta</td>
		<td>&nbsp;$<%
total = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Ordenes.Estado = 19)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total = total + ObRs ("Cantidad") * ObRs ("PrecioUnitario") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total,2)
%>
</td>
		<td>&nbsp;$<%
total1 = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Ordenes.Estado = 19)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total1 = total1 + ObRs ("Cantidad") * ObRs ("PrecioCosto") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total1,2)
%>
</td>
		<td>&nbsp;$<%
Response.Write formatnumber(total - total1,2)
		%></td>
		<td>&nbsp;<%
if total > 0 then
Response.Write formatnumber(100-(total1/total)*100,2)
end if
		%>%</td>
	</tr>
	<tr>
		<td width="81">Total</td>
		<td>&nbsp;$<%
total = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Ordenes.Estado = 6 or Ordenes.Estado = 9 or Ordenes.Estado = 19)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total = total + ObRs ("Cantidad") * ObRs ("PrecioUnitario") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total,2)
%>
</td>
		<td>&nbsp;$<%
total1 = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Equipos.Tipo, Ordenes.Estado, OrdenItem.Cantidad, "
SQL = SQL & "OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "Ordenes.FechaIngreso, Ordenes.FechaEstado "
SQL = SQL & "FROM (Equipos INNER JOIN Ordenes ON Equipos.Id = Ordenes.Equipo) "
SQL = SQL & "INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden "
if Request.Form("desde") = "" or Request.Form("hasta") = "" then
SQL = SQL & "WHERE Ordenes.Estado = 0"
else
SQL = SQL & "WHERE Ordenes.FechaIngreso >= #" & Request.Form("desde") & "# AND Ordenes.FechaIngreso <= #" & Request.Form("hasta") & "# "
SQL = SQL & "AND (Ordenes.Estado = 6 or Ordenes.Estado = 9 or Ordenes.Estado = 19)"
end if
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
total1 = total1 + ObRs ("Cantidad") * ObRs ("PrecioCosto") 
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
Response.Write formatnumber(total1,2)
%>
</td>
		<td>&nbsp;$<%
Response.Write formatnumber(total - total1,2)
		%></td>
		<td>&nbsp;<%
if total > 0 then
Response.Write formatnumber(100-(total1/total)*100,2)
end if
		%>%</td>
	</tr>
</table>

</body>

</html>
