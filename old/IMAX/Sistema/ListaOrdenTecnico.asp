<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Language" content="es-ar" />
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
<title>Lista de Ordenes por técnico</title>
<style type="text/css">
.style1 {
	font-size: large;
}
.style2 {
	background-color: #0000FF;
}
</style>
</head>

<body>

<p class="style1">Lista de Ordenes por técnico</p>
<form method="post" action="">
		Tecnico Asignado<table style="width: 100%">
			<tr>
				<td style="width: 160px">
<%Response.Write Request.Form("TecnicoAsignado")%>
<select name="TecnicoAsignado" style="width: 125px">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios Where habilitado = true and Area = 1 order by Nick"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof

IF Request.Form("TecnicoAsignado") = ObRs ("Id") Then
untecnico = 1
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

IF untecnico = 1 Then
%>
<option value="0">Ninguno</option>
<%
ELSE
%>
<option selected value="0">Ninguno</option>
<%
END IF

%>
</select>	
<br />
		<input name="Button1" type="submit" value="Enviar" /></td>
				<td>Nota<br />
				Solo se muestran los equipos con estado No Visto, En Banco y 
				Confirmado</td>
			</tr>
		</table>
		<br />
</form>
<table style="width: 100%">
	<tr>
		<td class="style2"><strong>Orden</strong></td>
		<td class="style2"><strong>Cliente</strong></td>
		<td class="style2" style="width: 54px"><strong>Equipo</strong></td>
		<td class="style2"><strong>Serie</strong></td>
		<td class="style2"><strong>Estado</strong></td>
		<td class="style2" style="width: 701px"><strong>Observaciones de Ingreso</strong></td>
		<td class="style2"><strong>Fecha de Estado</strong></td>
		<td class="style2"><strong>Fecha de Ingreso</strong></td>
		<td class="style2"><strong>Tecnico Asignado</strong></td>
		<td class="style2"><strong>Dias Restantes</strong></td>
	</tr>
<%
if request.form ("TecnicoAsignado") = eof then
TecnicoAsignado = 0
else
TecnicoAsignado = request.form ("TecnicoAsignado")
end if

SQL = "SELECT Ordenes.Id, Clientes.Nombre, Equipos.Modelo, Equipos.Id, Ordenes.Serie,"
SQL = SQL & " Estados.Estado, Ordenes.FechaEstado, Ordenes.FechaIngreso, Usuarios.Nick,"
SQL = SQL & " Ordenes.FechaProgramada, Clientes.TipoCliente, Ordenes.ObservacionIngreso"
SQL = SQL & " FROM (Estados INNER JOIN (Equipos INNER JOIN"
SQL = SQL & " (Clientes INNER JOIN Ordenes ON Clientes.Id = Ordenes.Cliente)"
SQL = SQL & " ON Equipos.Id = Ordenes.Equipo) ON Estados.Id = Ordenes.Estado)"
SQL = SQL & " INNER JOIN Usuarios ON Ordenes.TecnicoAsignado = Usuarios.Id"
SQL = SQL & " Where (Ordenes.Estado = 1 or Ordenes.Estado = 3 or Ordenes.Estado = 13)"
SQL = SQL & " And Ordenes.TecnicoAsignado = " & TecnicoAsignado
SQL = SQL & " Order by Ordenes.FechaProgramada"

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open SQL,ObConn

DO WHILE NOT ObRs.Eof
Select Case ObRs ("TipoCliente")
Case 2
Color = "#FFFFAA"
Case 4
Color = "#FFCC66"
Case 5
Color = "#99FF66"
Case else
Color = "#FFFFFF"
End Select
%>
	<tr>
		<td width="63" bgcolor="<%Response.Write Color%>">&nbsp;<a href="ConsultaDeOrden.asp?Id=<%Response.Write ObRs (0)%>"><%Response.Write ObRs (0)%></a></td>
		<td width="217" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Nombre")%></td>
		<td bgcolor="<%Response.Write Color%>" style="width: 54px">&nbsp;<%Response.Write ObRs ("Modelo")%><img alt="" src="imagen.asp?Id=<%Response.Write ObRs (3)%>&Tabla=Equipos" width="45" height="36" /></td>
		<td width="96" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Serie")%></td>
		<td width="96" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Estado")%></td>
		<td bgcolor="<%Response.Write Color%>" style="width: 701px">&nbsp;<%Response.Write ObRs("ObservacionIngreso")%></td>
		<td width="72" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("FechaEstado")%></td>
		<td width="65" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("FechaIngreso")%></td>
		<td bgcolor="<%Response.Write Color%>" style="width: 70px">&nbsp;<%Response.Write ObRs ("Nick")%></td>
		<td bgcolor="<%Response.Write Color%>" style="width: 70px">
		<%
diashabiles = 0
i = date
do while i <= ObRs("FechaProgramada") - 1
	i = i+1
	j = DatePart("w", i)
	if j <> 1 and j <> 7 then
		diashabiles = diashabiles + 1
	end if
loop  
response.write diashabiles
		%>
		</td>
	</tr>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close

%>
</table>

</body>

</html>
