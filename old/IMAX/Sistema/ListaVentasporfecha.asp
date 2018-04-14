<!DOCTYPE html PUBLIC
          "-//W3C//DTD XHTML 1.0 Transitional//EN"
          "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Cantidad de equipos ingresados</title>
    <script src="src/js/jscal2.js"></script>
    <script src="src/js/lang/es.js"></script>
    <link rel="stylesheet" type="text/css" href="src/css/jscal2.css" />
    <link rel="stylesheet" type="text/css" href="src/css/border-radius.css" />
    <link rel="stylesheet" type="text/css" href="src/css/steel/steel.css" />
</head>

<body>
<p><b><font face="Arial">Lista de ventas</font></b></p>
<form method="POST" action="ListaVentasporfecha.asp" webbot-action="--WEBBOT-SELF--">
	<!--webbot bot="SaveResults" U-File="_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" startspan --><input NAME="VTI-GROUP" TYPE="hidden" VALUE="0"><!--webbot bot="SaveResults" i-checksum="37496" endspan -->
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td width="160">
	<p>Desde:
    <input size="11" id="f_date1" name="desde" /><button id="f_btn1">...</button><br />
    Hasta:&nbsp;
    <input size="11" id="f_date2" name="hasta" /><button id="f_btn2">...</button>
	</p>
	<p><input type="submit" value="Enviar" name="B1"></p>

			</td>
			<td valign="top">Estado:<select size="1" name="Estado">
	<option value="0">Seleccionar</option>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Estados",ObConn
DO WHILE NOT ObRs.Eof
If int(Request.Form("Estado")) = ObRs ("Id") THEN
%>
	<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Estado")%></option>
<%
SQLEstado = " Estado = '" & ObRs ("Estado") & "'"

ELSE
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Estado")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select></td>
		</tr>
	</table>
</form>
    <script type="text/javascript">//<![CDATA[

      var cal = Calendar.setup({
          onSelect: function(cal) { cal.hide() }
      });
      cal.manageFields("f_btn1", "f_date1", "%Y/%m/%d");
      cal.manageFields("f_btn2", "f_date2", "%Y/%m/%d");
    //]]></script>
<table border="1" width="738" id="table1" cellspacing="0" cellpadding="0" bordercolor="#000000">
	<tr>
		<td width="63" bgcolor="#3399FF"><font face="Arial" size="2"><b>Nº Orden</b></font></td>
		<td width="76" bgcolor="#3399FF"><b><font face="Arial" size="2">Cliente</font></b></td>
		<td bgcolor="#3399FF" width="165"><b><font face="Arial" size="2">Equipo</font></b></td>
		<td bgcolor="#3399FF" width="144"><font face="Arial" size="2"><b>Estado</b></font></td>
		<td width="107" bgcolor="#3399FF"><b><font face="Arial" size="2">Fecha de Estado</font></b></td>
		<td width="81" bgcolor="#3399FF"><b><font face="Arial" size="2">Precio Venta</font></b></td>
		<td width="86" bgcolor="#3399FF"><b><font face="Arial" size="2">Precio Costo</font></b></td>
	</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.Id, Clientes.Nombre, Equipos.Modelo, Estados.Estado, "
SQL = SQL & "Ordenes.FechaEstado, OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto, "
SQL = SQL & "OrdenItem.Cantidad FROM Estados INNER JOIN ((Equipos INNER JOIN "
SQL = SQL & "(Clientes INNER JOIN Ordenes ON Clientes.Id = Ordenes.Cliente) "
SQL = SQL & "ON Equipos.Id = Ordenes.Equipo) INNER JOIN OrdenItem ON "
SQL = SQL & "Ordenes.Id = OrdenItem.Orden) ON Estados.Id = Ordenes.Estado "

SQL = SQL & "Where (Ordenes.Id > 19) "
if Request.Form ("desde") <> "" then
SQL = SQL & "AND Ordenes.FechaEstado >= #" & Request.Form("desde") & "# "
end if
if Request.Form ("hasta") <> "" then
SQL = SQL & "AND Ordenes.FechaEstado <= #" & Request.Form("hasta") & "# "
end if
SQL = SQL & "AND Ordenes.Estado = " & Request.Form("Estado") & " "


SQL = SQL & "Order by Ordenes.FechaEstado"

if (Request.Form ("desde") = "" and Request.Form ("hasta") = "") then
SQL = "select * from Ordenes where Id = 0"
end if

ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<tr>
		<td width="63" bgcolor="#FFFFFF"><font face="Arial" size="2"><a href="ConsultaDeOrden.asp?Id=<%Response.Write ObRs ("id")%>">&nbsp;<%Response.Write ObRs ("id")%></a></font></td>
		<td width="76" bgcolor="#FFFFFF"><font face="Arial" size="2">&nbsp;<%Response.Write ObRs ("Nombre")%></font></td>
		<td bgcolor="#FFFFFF" width="165"><font face="Arial" size="2">&nbsp;<%Response.Write ObRs ("Modelo")%></font></td>
		<td bgcolor="#FFFFFF" width="144"><font face="Arial" size="2">&nbsp;<%Response.Write ObRs ("Estado")%></font></td>
		<td width="107" bgcolor="#FFFFFF"><font face="Arial" size="2">&nbsp;<%Response.Write ObRs ("FechaEstado")%></font></td>
		<td width="81" bgcolor="#FFFFFF"><font face="Arial" size="2">&nbsp;$<%Response.Write ObRs ("PrecioUnitario") * ObRs ("Cantidad")%></font></td>
		<td width="86" bgcolor="#FFFFFF"><font face="Arial" size="2">&nbsp;$<%Response.Write ObRs ("PrecioCosto") * ObRs ("Cantidad")%></font></td>
	</tr>

<%

PrecioVenta = ObRs ("PrecioUnitario") * ObRs ("Cantidad") + PrecioVenta
PrecioCosto = ObRs ("PrecioCosto") * ObRs ("Cantidad") + Preciocosto

ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	<tr>
		<td bgcolor="#FFFFFF" colspan="5" bordercolor="#000000">&nbsp;</td>
		<td width="81" bgcolor="#FFFFFF"><font face="Arial" size="2">&nbsp;$<%Response.Write PrecioVenta%></font></td>
		<td width="86" bgcolor="#FFFFFF"><font face="Arial" size="2">&nbsp;$<%Response.Write PrecioCosto%></font></td>
	</tr>

	<tr>
		<td bgcolor="#FFFFFF" colspan="5" bordercolor="#000000"><b>
		<font face="Arial" size="2">Utilidad</font></b></td>
		<td width="167" bgcolor="#FFCC99" colspan="2" align="center">
		<font face="Arial" size="2">$<%Response.Write PrecioVenta - PrecioCosto%></font></td>
	</tr>

</table>
</body>

</html>
