<!DOCTYPE html PUBLIC
          "-//W3C//DTD XHTML 1.0 Transitional//ES"
          "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
    <script src="src/js/jscal2.js"></script>
    <script src="src/js/lang/es.js"></script>
    <link rel="stylesheet" type="text/css" href="src/css/jscal2.css" />
    <link rel="stylesheet" type="text/css" href="src/css/border-radius.css" />
    <link rel="stylesheet" type="text/css" href="src/css/steel/steel.css" />
</head>

<body>
<form method="Get" action="ListaOrdenEstado.asp" webbot-action="--WEBBOT-SELF--">
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
      <tr>
        <td width="34%">Estado: <select size="1" name="Estado">
	<option value="">Todos</option>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Estados",ObConn
DO WHILE NOT ObRs.Eof
If Request.QueryString("Estado") = ObRs ("Estado") THEN
%>
	<option selected value="<%Response.Write ObRs ("Estado")%>"><%Response.Write ObRs ("Estado")%></option>
<%
SQLEstado = " Estado = '" & ObRs ("Estado") & "'"

ELSE
%>
	<option value="<%Response.Write ObRs ("Estado")%>"><%Response.Write ObRs ("Estado")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select><br>Cliente:
	<select size="1" name="Cliente">
	<option value="">Todos</option>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select * From Clientes Order By Nombre",ObConn
DO WHILE NOT ObRs.Eof
If Request.QueryString("Cliente") = ObRs ("Nombre") THEN
%>
	<option selected value="<%Response.Write ObRs ("Nombre")%>"><%Response.Write ObRs ("Nombre")%></option>
<%
SQLCliente = " Nombre = '" & ObRs ("Nombre") & "'"

ELSE
%>
	<option value="<%Response.Write ObRs ("Nombre")%>"><%Response.Write ObRs ("Nombre")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select>
</td>
        <td width="16%" valign="top">N° de referencia<br>
		<input type="text" name="Referencia" size="20" value="<%Response.Write Request.QueryString("Referencia")%>"></td>
        <td width="50%" valign="top">Filtrar<input type="checkbox" name="Filtro" value="1"><br>
        <input type="radio" value="note" checked name="solonote">Solamente 
        Notebooks<br>
        <input type="radio" name="solonote" value="nonote">Todo sin Notebooks</td>
      </tr>
      <tr>
        <td width="34%">Fecha de Ingreso<br>Desde:
    <input size="11" id="f_date1" name="desde" value="<%Response.Write Request.QueryString("desde")%>" /><button id="f_btn1">...</button><br />
    Hasta:&nbsp;
    <input size="11" id="f_date2" name="hasta" value="<%Response.Write Request.QueryString("hasta")%>" /><button id="f_btn2">...</button>
		<br>
		<input type="submit" value="Submit" name="B1"></td>
        <td width="16%" valign="top">
    Fecha de Estado<br>Desde:
    <input size="11" id="f_date3" name="desde1" value="<%Response.Write Request.QueryString("desde1")%>" /><button id="f_btn3">...</button><br />
    Hasta:&nbsp;
    <input size="11" id="f_date4" name="hasta1" value="<%Response.Write Request.QueryString("hasta1")%>" /><button id="f_btn4">...</button>
</td>
        <td width="50%" valign="top"><span lang="es-ar"><br />
		Tecnico Asignado<br />
<select name="TecnicoAsignado" style="width: 125px">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios Where habilitado = true and Area = 1 order by Nick"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
IF TecnicoAsignado = ObRs ("Id") Then
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
		</span></td>
      </tr>
    </table>
</form>
<table border="1" width="733" id="table1">
	<tr>
		<td width="63" bgcolor="#3399FF"><b>Nº Orden</b></td>
		<td width="217" bgcolor="#3399FF"><b>Cliente</b></td>
		<td bgcolor="#3399FF"><b>Equipo</b></td>
		<td width="96" bgcolor="#3399FF"><b>Serie</b></td>
		<td width="96" bgcolor="#3399FF"><b>Estado</b></td>
		<td width="72" bgcolor="#3399FF"><b>Observaciones de Ingreso</b></td>
		<td width="72" bgcolor="#3399FF"><b>Fecha de Estado</b></td>
		<td width="65" bgcolor="#3399FF"><b>Fecha de Ingreso</b></td>
		<td bgcolor="#3399FF" style="width: 70px"><b>Tecnico Asignado</b></td>
		<td bgcolor="#3399FF" style="width: 70px"><span lang="es-ar"><strong>
		Dias Restantes</strong></span></td>
	</tr>

<%
if Request.QueryString("solonote") <> "" then


SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
IF SQLCliente <> "" THEN
y1 = " And"
END IF
IF SQLEstado <> "" THEN
y2 = " And"
END IF
If Request.QueryString("Referencia") <> "" then
Referencia = " Referencia Like '%" & Request.QueryString("Referencia") & "%' and"
End If
If Request.QueryString("filtro") = false then
SQL = "Select * From ConsultaOrdenes where" & Referencia & " Estado <> 'Anulado' And Estado <> 'Vendido'" & y1 & SQLCliente & y2 & SQLEstado
else
Select case Request.QueryString("solonote")
case "note"
SQL = "Select * From ConsultaOrdenes where" & Referencia & " Estado <> 'Anulado' And Estado <> 'Vendido' And (Tipo = 'Notebook' or Tipo = 'Motherboard' or Tipo = 'Netbook' or Tipo = 'Fuente' or Tipo = 'MONITOR LCD')" & y1 & SQLCliente & y2 & SQLEstado
case "nonote"
SQL = "Select * From ConsultaOrdenes where" & Referencia & " Estado <> 'Anulado' And Estado <> 'Vendido' And (Tipo <> 'Notebook' and Tipo <> 'Motherboard' and Tipo <> 'Netbook' and Tipo <> 'Fuente' or Tipo <> 'MONITOR LCD')" & y1 & SQLCliente & y2 & SQLEstado
end select
end if

if Request.QueryString("TecnicoAsignado") <> 0 then
SQL = SQL & " AND TecnicoAsignado = " & Request.QueryString("TecnicoAsignado")
end if

If Request.QueryString("desde") <> "" then
SQL = SQL & " AND FechaIngreso >= #" & Request.QueryString("desde") & "#"
end if
If Request.QueryString("hasta") <> "" then
SQL = SQL & " AND FechaIngreso <= #" & Request.QueryString("hasta") & "#"
end if
If Request.QueryString("desde1") <> "" then
SQL = SQL & " AND FechaEstado >= #" & Request.QueryString("desde1") & "#"
end if
If Request.QueryString("hasta1") <> "" then
SQL = SQL & " AND FechaEstado <= #" & Request.QueryString("hasta1") & "#"
end if

SQL = SQL & " Order By Id"

ObRs.Open  SQL,ObConn
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
		<td width="63" bgcolor="<%Response.Write Color%>">&nbsp;<a href="ConsultaDeOrden.asp?Id=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a></td>
		<td width="217" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Nombre")%></td>
		<td bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Modelo")%></td>
		<td width="96" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Serie")%></td>
		<td width="96" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("Estado")%></td>
		<td width="72" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs("ObservacionIngreso")%></td>
		<td width="72" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("FechaEstado")%></td>
		<td width="65" bgcolor="<%Response.Write Color%>">&nbsp;<%Response.Write ObRs ("FechaIngreso")%></td>
		<td bgcolor="<%Response.Write Color%>" style="width: 70px">&nbsp;<%Response.Write ObRs ("TecnicoAsignado")%></td>
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

end if
%>

</table>


    <p><br />
</p>
<p>&nbsp;</p>
<p><span lang="es-ar">.</span></p>


    <script type="text/javascript">//<![CDATA[

      var cal = Calendar.setup({
          onSelect: function(cal) { cal.hide() }
      });
      cal.manageFields("f_btn1", "f_date1", "%Y/%m/%d");
      cal.manageFields("f_btn2", "f_date2", "%Y/%m/%d");
      cal.manageFields("f_btn3", "f_date3", "%Y/%m/%d");
      cal.manageFields("f_btn4", "f_date4", "%Y/%m/%d");

    //]]></script>

</body>

</html>