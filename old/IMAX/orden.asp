<html>

<head>
<meta http-equiv="Content-Language" content="es-ar">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Pagina nueva 1</title>

<script LANGUAGE="JavaScript">
var pagina="confirmacion.asp"
function redireccionar() 
{
 if (confirm("Desea Confirmar la orden"))
	{
		location.href=pagina
	}
} 
</script>

</head>

<body text="#FFFFFF" bgcolor="#000000">
<%
IF Session("Session") = "" THEN
Response.Redirect ("inicio.asp")
End If
Session ("Orden") = Request.QueryString("Orden")
%>


<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes where Id = "& Request.QueryString("Orden")
ObRs.Open  SQL,ObConn
Orden = ObRs("Id")
Equipo = ObRs("Equipo")
FechaIngreso = ObRs("FechaIngreso")
FechaEstado = ObRs("FechaEstado")
Estado = ObRs("Estado")
Serie = ObRs("Serie")
ObservacionIngreso = ObRs("ObservacionIngreso")
ObservacionTecnico = ObRs("ObservacionTecnico")
ObRs.Close
ObConn.Close
%>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Estados where Id = "& Estado
ObRs.Open  SQL,ObConn
Estado = ObRs("Estado")
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Equipos where Id = "& Equipo
ObRs.Open  SQL,ObConn
IdEquipo = ObRs("Id")
Marca = ObRs("Marca")
Tipo = ObRs("Tipo")
Modelo = ObRs("Modelo")
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Marcas where Id = "& Marca
ObRs.Open  SQL,ObConn
Marca = ObRs("Marca")
ObRs.Close
ObConn.Close
%>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From TiposdeEquipos where Id = "& Tipo
ObRs.Open  SQL,ObConn
Tipo = ObRs("Tipo")
ObRs.Close
ObConn.Close
%>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & Orden
ObRs.Open SQL,ObConn
Total = 0
DO WHILE NOT ObRs.Eof
Total = Total + ObRs ("Cantidad") * ObRs ("PrecioUnitario")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>

<p><b><font face="Arial">Orden: <%response.write Orden%></font></b></p>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1">
  <tr>
    <td width="50%"><font face="Arial" size="2"><b>Datos del Equipo:</b><br>
<b>Tipo:</b> <%response.write Tipo%> <b>&nbsp; Marca: </b><%response.write Marca%> <b>&nbsp; Modelo:</b> <%response.write Modelo%><br>
<b>Serie:</b> <%response.write Serie%></font></td>
    <td width="50%"><img border="0" src="imagen.asp?Id=<%Response.Write IdEquipo%>&Tabla=Equipos" height="86" width="91"></td>
  </tr>
</table>

<p><font face="Arial" size="2"><b>Fecha de Ingreso:</b> <%response.write FechaIngreso%><br>
<b>Estado:</b> <%response.write Estado%><b>

<br>Fecha de Estado:</b> <%response.write FechaEstado%></font></p>
<p><font face="Arial" size="2">
<b>Observaciones de Ingreso:</b></font></p>
<table style="BORDER-COLLAPSE: collapse" borderColor="#111111" height="47" cellSpacing="0" cellPadding="0" width="50%" border="1">
  <tr>
    <td vAlign="top" align="left" width="100%" bgColor="#ffffff" height="47">
    <font color="#000000"><%response.write ObservacionIngreso%></font></td>
  </tr>
</table>
<p><font face="Arial" size="2"><b>Observaciones del Técnico:</b></font></p>
<table style="BORDER-COLLAPSE: collapse" borderColor="#111111" height="45" cellSpacing="0" cellPadding="0" width="50%" border="1">
  <tr>
    <td vAlign="top" align="left" width="100%" bgColor="#ffffff" height="45">
    <font color="#000000"><%response.write ObservacionTecnico%></font></td>
  </tr>
</table>
<p><br></p>
<table border="1" width="100%" id="table1" cellspacing="1" bordercolor="#000000" style="border-collapse: collapse">
		<tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString("Orden")
ObRs.Open SQL, ObConn
%><font face="Arial" size="2"> </font>

			<td width="51">&nbsp;</td>
<%
ObRs.Close
ObConn.Close
%> </font>
			<td width="51"><font face="Arial" size="2"><b>Código</b></font></td>
			<td width="446"><font face="Arial" size="2"><b>Descripción</b></font></td>
			<td width="55"><font face="Arial" size="2"><b>Cantidad</b></font></td>
			<td width="91"><font face="Arial" size="2"><b>Precio Unitario</b></font></td>
			<td><font face="Arial" size="2"><b>Total</b></td>
		</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT OrdenItem.Id, OrdenItem.Orden, Items.Codigo, Items.Descripcion, OrdenItem.Cantidad, OrdenItem.PrecioUnitario, OrdenItem.Carrito, OrdenItem.PrecioCosto "
SQL = SQL & "FROM Items INNER JOIN OrdenItem ON Items.Id = OrdenItem.Item "
SQL = SQL & " Where Orden = " & Request.QueryString("Orden")

ObRs.Open SQL,ObConn
Total = 0
DO WHILE NOT ObRs.Eof
%> </font>
		<tr>
			<td width="51">
			</td>
			<td width="51"><%Response.Write ObRs ("Codigo")%>&nbsp;</td>
			<td width="446"><%Response.Write ObRs ("Descripcion")%>&nbsp;</td>
			<td width="55"><%Response.Write ObRs ("Cantidad")%>&nbsp;</td>
			<td width="91">$<%Response.Write ObRs ("PrecioUnitario")%>&nbsp;($<%Response.Write ObRs ("PrecioCosto")%>)</td>
			<td>$<%Response.Write ObRs ("Cantidad") * ObRs ("PrecioUnitario")%>&nbsp;</td>
		</tr>
<%
Total = Total + ObRs ("Cantidad") * ObRs ("PrecioUnitario")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
  <font face="Arial" size="2"> </font>
		<tr>
			<td width="643" colspan="5">
			<p align="right"><font face="Arial" size="2"><b>Total</b></font></td>
			<td>$<%Response.Write Total%>&nbsp;</td>
		</tr>
	</table>

<p><b><font face="Arial" size="2">Presupuesto: $<%response.write Total%>
<%
If Estado = "Presupuestado" then
%>
<font color="#0000FF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

<button onclick="redireccionar()" style="height: 25px; width:83px" name="boton">
<font face="Arial Black">Confirmar</font></button>


<%
End If
%> </font></b>
</p>

</body>

</html>