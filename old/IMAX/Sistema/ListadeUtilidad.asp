<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>
</head>

<body>
<%
if Session("IMAX") = False then
Response.Redirect ("administrativo.asp")
End if
%>

<form method="Post" action="ListadeUtilidad.asp" webbot-action="--WEBBOT-SELF--">
	<p>Técnico
	<select size="1" name="Tecnico">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select Id, Nick From Usuarios Where Habilitado = true Order By Nick",ObConn
DO WHILE NOT ObRs.Eof
If int(Request.Form("Tecnico")) = int(ObRs("Id")) THEN
%>
	<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Nick")%></option>
<%
FiltroUsuario = "And Ordenes.UsuarioEstado = " & ObRs ("Id") & " "
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
<%If DatePart("m", Date) = 1 then Response.Write "Selected"%>
	</select>&nbsp;&nbsp;&nbsp;&nbsp; Mes:<select size="1" name="Mes">
	<option <%If (DatePart("m", Date) = 1 and Request.Form("Mes") = "") or Request.Form("Mes") = 1 then Response.Write "Selected "%>value="01">Enero</option>
	<option <%If (DatePart("m", Date) = 2 and Request.Form("Mes") = "") or Request.Form("Mes") = 2 then Response.Write "Selected "%>value="02">Febrero</option>
	<option <%If (DatePart("m", Date) = 3 and Request.Form("Mes") = "") or Request.Form("Mes") = 3 then Response.Write "Selected "%>value="03">Marzo</option>
	<option <%If (DatePart("m", Date) = 4 and Request.Form("Mes") = "") or Request.Form("Mes") = 4 then Response.Write "Selected "%>value="04">Abril</option>
	<option <%If (DatePart("m", Date) = 5 and Request.Form("Mes") = "") or Request.Form("Mes") = 5 then Response.Write "Selected "%>value="05">Mayo</option>
	<option <%If (DatePart("m", Date) = 6 and Request.Form("Mes") = "") or Request.Form("Mes") = 6 then Response.Write "Selected "%>value="06">Junio</option>
	<option <%If (DatePart("m", Date) = 7 and Request.Form("Mes") = "") or Request.Form("Mes") = 7 then Response.Write "Selected "%>value="07">Julio</option>
	<option <%If (DatePart("m", Date) = 8 and Request.Form("Mes") = "") or Request.Form("Mes") = 8 then Response.Write "Selected "%>value="08">Agosto</option>
	<option <%If (DatePart("m", Date) = 9 and Request.Form("Mes") = "") or Request.Form("Mes") = 9 then Response.Write "Selected "%>value="09">Septiembre</option>
	<option <%If (DatePart("m", Date) = 10 and Request.Form("Mes") = "") or Request.Form("Mes") = 10 then Response.Write "Selected "%>value="10">Octubre</option>
	<option <%If (DatePart("m", Date) = 11 and Request.Form("Mes") = "") or Request.Form("Mes") = 11 then Response.Write "Selected "%>value="11">Noviembre</option>
	<option <%If (DatePart("m", Date) = 12 and Request.Form("Mes") = "") or Request.Form("Mes") = 12 then Response.Write "Selected "%>value="12">Diciembre</option>
	</select> Año:<input type="text" name="anio" size="7" value="<%If Request.Form("anio") = "" Then Response.Write CStr(DatePart("yyyy", Date)) Else Response.Write Request.Form("anio")%>"><input type="submit" value="Enviar" name="B1"></p>
</form>
<table border="1" width="738" id="table1" cellspacing="0" cellpadding="0" bordercolor="#000000">
	<tr>
		<td width="63" bgcolor="#3399FF"><font face="Arial" size="2"><b>Nº Orden</b></font></td>
		<td width="76" bgcolor="#3399FF"><font face="Arial" size="2"><b>Usuario</b></font></td>
		<td bgcolor="#3399FF" width="165"><b><font face="Arial" size="2">Equipo</font></b></td>
		<td bgcolor="#3399FF" width="144"><font face="Arial" size="2"><b>Estado</b></font></td>
		<td width="107" bgcolor="#3399FF"><b><font face="Arial" size="2">Fecha de Estado</font></b></td>
		<td width="81" bgcolor="#3399FF"><b><font face="Arial" size="2">Precio Venta</font></b></td>
		<td width="86" bgcolor="#3399FF"><b><font face="Arial" size="2">Precio Costo</font></b></td>
	</tr>
<%

mes2 = ""

If Request.Form("Anio") = "" then
anio = DatePart("yyyy", Date)
Else
anio = Request.Form("Anio")
end if
anio2 = anio

If Request.Form("Mes") = "" then
mes = DatePart("m", Date)
Else
mes = Request.Form("Mes")
end if

mes2 = mes + 1
if mes2 = 13 then
mes2 = 1
anio2 = anio + 1
end if


PrecioVenta = 0
PrecioCosto = 0
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Ordenes.Id, Ordenes.FechaEstado, Equipos.Modelo, Estados.Estado, Usuarios.Nick, OrdenItem.Item, OrdenItem.Cantidad, OrdenItem.PrecioUnitario, OrdenItem.PrecioCosto "
SQL = SQL & "FROM Estados INNER JOIN (Equipos INNER JOIN (Usuarios INNER JOIN (Ordenes INNER JOIN OrdenItem ON Ordenes.Id = OrdenItem.Orden) ON Usuarios.Id = Ordenes.UsuarioEstado) ON Equipos.Id = Ordenes.Equipo) ON Estados.Id = Ordenes.Estado "
SQL = SQL & "Where (Ordenes.Estado = 9 or Ordenes.Estado = 6) and Ordenes.FechaEstado between #"& mes &"/01/"& anio &"# and #"& mes2 &"/01/"& anio2 &"# - 1 " & FiltroUsuario
SQL = SQL & "Order by Ordenes.FechaEstado"
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<tr>
		<td width="63" bgcolor="#FFFFFF"><font face="Arial" size="2"><a href="ConsultaDeOrden.asp?Id=<%Response.Write ObRs ("id")%>">&nbsp;<%Response.Write ObRs ("id")%></a></font></td>
		<td width="76" bgcolor="#FFFFFF"><font face="Arial" size="2">&nbsp;<%Response.Write ObRs ("Nick")%></font></td>
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