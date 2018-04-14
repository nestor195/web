<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<form method="Get" action="ListaItems.asp" webbot-action="--WEBBOT-SELF--">
<span lang="es">Búsqueda por Descripción</span>:<input type="text" name="Descripcion" size="20"> </p>
	<p><input type="submit" value="Enviar" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>
<table border="1" width="733" id="table1">
	<tr>
		<td width="63">Nº Ítem</td>
		<td width="58">Código</td>
		<td width="242">Descripción</td>
		<td width="136"><span lang="es">Fecha de Precio</span></td>
		<td width="136">Precio Costo</td>
		<td width="129">Precio Sugerido</td>
		<td width="129">Stock</td>
	</tr>
<%

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
If Request.QueryString("Descripcion") = "" then
SQL = "Select * From Items order by Descripcion"
else
SQL = "Select * From Items where Descripcion Like '%" & Request.QueryString("Descripcion") & "%' Order By Descripcion"
end if
ObRs.Open  SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
	<tr>
		<td width="63"><a href="ModificarItem.asp?IdItem=<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Id")%></a>&nbsp;</td>
		<td width="58">&nbsp;<%Response.Write ObRs ("Codigo")%></td>
		<td width="242">&nbsp;<%Response.Write ObRs ("Descripcion")%></td>
		<td width="136">&nbsp;<%Response.Write ObRs ("FechaPrecio")%></td>
		<td width="136">&nbsp;<%Response.Write ObRs ("PrecioCosto")%></td>
		<td width="129">&nbsp;<%Response.Write ObRs ("PrecioSugerido")%></td>
		<td width="129">&nbsp;<%Response.Write ObRs ("Stock")%></td>
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