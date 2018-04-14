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
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Items Where Id = " & Request.QueryString("IdItem")
ObRs.Open SQL, ObConn
%>
<b>Modificar de Item</b><form method="Post" action="ModificarItem.asp" webbot-action="--WEBBOT-SELF--">
	<p>Id: <select size="1" name="Id">
    <option value="<%Response.Write ObRs("Id")%>" selected><%Response.Write ObRs("Id")%></option>
    </select><br>
Código: <input type="text" name="Codigo" size="30" value="<%Response.Write ObRs("Codigo")%>"><br>
	Descripción: <input type="text" name="Descripcion" size="37" value="<%Response.Write ObRs("Descripcion")%>"><br>
	<span lang="es">Fecha Precio:
    <input type="text" name="FechaPrecio" size="20" value="<%Response.Write ObRs("FechaPrecio")%>"></span><br>
	Precio de Costo: 
    <input type="text" name="PrecioCosto" size="20" value="<%Response.Write ObRs("PrecioCosto")%>"><br>
	Precio Sugerido: 
    <input type="text" name="PrecioSugerido" size="20" value="<%Response.Write ObRs("PrecioSugerido")%>"><br>
	Stock:
	<input type="text" name="Stock" size="8" value="<%Response.Write ObRs("Stock")%>"><br>

	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ObRs.Close
ObConn.Close
%>
<form method="GET" action="IngresoItemEquipo.asp">
	<table border="1" width="685" id="table1">
		<tr>
			<td width="50">&nbsp;</td>
			<td width="47"><span lang="es"><b>Id</b></span></td>
			<td width="70"><span lang="es"><b>Tipo</b></span></td>
			<td width="181"><span lang="es"><b>Marca</b></span></td>
			<td width="261"><span lang="es"><b>Modelo</b></span></td>
		</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaItemEquipo Where Item = " & Request.QueryString("IdItem") & " Order By Modelo"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
		<tr>
			<td width="50">
			<a target="_parent" href="EliminarItemEquipo.asp?Id=<%Response.Write ObRs ("Id")%>">
			Eliminar</a></td>
			<td width="47"><%Response.Write ObRs ("Equipo")%>&nbsp;</td>
			<td width="70"><%Response.Write ObRs ("Tipo")%>&nbsp;</td>
			<td width="181"><%Response.Write ObRs ("Marca")%>&nbsp;</td>
			<td width="261"><%Response.Write ObRs ("Modelo")%>&nbsp;</td>
		</tr>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</table>
	<p><select size="1" name="IdItem">
	<option value="<%Response.Write Request.QueryString("IdItem")%>">
	<%Response.Write Request.QueryString("IdItem")%></option>
	</select><input type="submit" value="Submit" name="B5">
	<input type="reset" value="Reset" name="B6"></p>
</form>

<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Items Where Id = " & Request.Form ("Id")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Codigo") = Request.Form ("Codigo")
ObRs ("Descripcion") = Request.Form ("Descripcion")
ObRs ("FechaPrecio") = Request.Form ("FechaPrecio")
ObRs ("PrecioCosto") = Request.Form ("PrecioCosto")
ObRs ("PrecioSugerido") = Request.Form ("PrecioSugerido")
ObRs ("Stock") = Request.Form ("Stock")
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