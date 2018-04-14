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
SQL = "Select * From Equipos Where Id = " & Request.QueryString("IdEquipo")
ObRs.Open SQL, ObConn
%>
<b>Modificar de Equipo</b><form method="POST" action="ModificarEquipo.asp" webbot-action="--WEBBOT-SELF--">
	<p>Id: <select size="1" name="IdEquipo">
    <option value="<%Response.Write ObRs("Id")%>" selected><%Response.Write ObRs("Id")%></option>
    </select><br>

<%
tipoequipo = ObRs("Tipo")
marcaequipo = ObRs("Marca")
modelo = ObRs("Modelo")

ObRs.Close
ObConn.Close
%>Tipo: <select size="1" name="Tipo">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "TiposDeEquipos",ObConn
DO WHILE NOT ObRs.Eof
If ObRs("Id") = tipoequipo then
%>
	<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Tipo")%></option>
<%
else
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Tipo")%></option>
<%
End If
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select> <a target="_parent" href="IngresoTipoDeEquipo.asp">Nuevo</a><br>
	Marca: <select size="1" name="Marca">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Marcas",ObConn
DO WHILE NOT ObRs.Eof
If ObRs ("Id") = marcaequipo then
%>
	<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Marca")%></option>
<%
Else
%>
	<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Marca")%></option>
<%
End If
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</select> <a target="_parent" href="IngresoMarca.asp">Nuevo</a><br>
	Modelo: <input type="text" name="Modelo" size="20" value="<%Response.Write Modelo%>"><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<form method="GET" action="IngresoEquipoItem.asp">
	<table border="1" width="685" id="table1">
		<tr>
			<td width="57">&nbsp;</td>
			<td width="34"><b>Ítem</b></td>
			<td width="70"><b>Código</b></td>
			<td width="358"><b>Descripción</b></td>
			<td width="84"><b>Precio Costo</b></td>
			<td width="36"><b>Precio Sugerido</b></td>
		</tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From ConsultaEquipoItem Where Equipo = " & Request.QueryString("IdEquipo")
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
		<tr>
			<td width="57">
			<a href="EliminarEquipoItem.asp?Id=<%Response.Write ObRs ("Id")%>"><img border="0" src="images/Delete.GIF" width="23" height="22"></a>
			<a href="ModificarItem.asp?IdItem=<%Response.Write ObRs ("Item")%>"><img border="0" src="images/Editar.gif" width="25" height="25"></a></td>
			<td width="34"><%Response.Write ObRs ("Item")%>&nbsp;</td>
			<td width="70"><%Response.Write ObRs ("Codigo")%>&nbsp;</td>
			<td width="358"><%Response.Write ObRs ("Descripcion")%>&nbsp;</td>
			<td width="84"><%Response.Write ObRs ("PrecioCosto")%>&nbsp;</td>
			<td width="36"><%Response.Write ObRs ("PrecioSugerido")%>&nbsp;</td>
		</tr>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
	</table>
	<p><select size="1" name="IdEquipo">
	<option value="<%Response.Write Request.QueryString("IdEquipo")%>">
	<%Response.Write Request.QueryString("IdEquipo")%></option>
	</select><input type="submit" value="Submit" name="B5">
	<input type="reset" value="Reset" name="B6"></p>
</form>
<p>&nbsp;<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * from Equipos where Id = " & Request.Form("IdEquipo")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Tipo") = Request.Form ("Tipo")
ObRs ("Marca") = Request.Form ("Marca")
ObRs ("Modelo") = Request.Form ("Modelo")
ObRs.Update

ObRs.Close
ObConn.Close
%></p>
<p>
<b>Datos Ingresados</b>
<%
END IF
%> </p>

</body>
</html>