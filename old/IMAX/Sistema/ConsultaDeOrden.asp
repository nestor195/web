<%
If Request.Form ("TecnicoAsignado") <> "" then 
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.Querystring ("Id")
ObRs.Open SQL,ObConn, 3, 3

diasprogramados = Request.Form ("DiasProgramados")
fechaprogramada = date
for i = 1 to diasprogramados
	fechaprogramada = fechaprogramada + 1
	j = DatePart("w", fechaprogramada)
	select case j
		case 1
			fechaprogramada = fechaprogramada + 1
		case 7
			fechaprogramada = fechaprogramada + 2
	end select
next

ObRs ("TecnicoAsignado") = Request.Form ("TecnicoAsignado")
ObRs ("FechaProgramada") = FechaProgramada
ObRs.Update

ObRs.Close
ObConn.Close
end if
%>

<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Consulta de Orden</title>
 <script type="text/javascript">
         function Referencia1()
         {
			if(Referencia.style.visibility == 'visible')
				{
				Referencia.style.visibility = 'hidden';
				}
			else
				{
				Referencia.style.visibility = 'visible'; 
				}
          }
         function Serial1()
         {
			if(Serial.style.visibility == 'visible')
				{
				Serial.style.visibility = 'hidden';
				}
			else
				{
				Serial.style.visibility = 'visible'; 
				}
          }
         function Reballi()
         {
			if(Reball.style.visibility == 'visible')
				{
				Reball.style.visibility = 'hidden';
				ReballEdit.style.visibility = 'hidden';
				}
			else
				{
				Reball.style.visibility = 'visible'; 
				}
          }
         function Reballiedit()
         {
			if(ReballEdit.style.visibility == 'visible')
				{
				ReballEdit.style.visibility = 'hidden';
				}
			else
				{
				ReballEdit.style.visibility = 'visible'; 
				}
          }
 </script>
<style type="text/css">
.style1 {
	font-family: Aharoni;
}
.style2 {
	margin-top: 0px;
}
.style3 {
	font-family: Arial, Helvetica, sans-serif;
}
.style4 {
	font-size: x-small;
}
</style>
</head>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn
IdOrden = ObRs("Id")
IdEquipo = ObRs("Equipo")
IdCliente = ObRs("Cliente")
IdEstado = ObRs("Estado")
IdUsuarioIngreso = ObRs("UsuarioIngreso")
IdUsuarioEstado = ObRs("UsuarioEstado")
TecnicoAsignado = ObRs("TecnicoAsignado")
FechaIngreso = ObRs("FechaIngreso")
FechaEstado = ObRs("FechaEstado")
Accesorios = ObRs("Accesorios")
ObservacionIngreso = ObRs("ObservacionIngreso")
ObservacionTecnico = ObRs("ObservacionTecnico")
ObservacionInterna = ObRs("ObservacionInterna")
Serie = ObRs("Serie")
Referencia = ObRs("Referencia")
TipoTarjeta = ObRs("TipoTarjeta")
ModeloR = ObRs("Modelo")
Nbridge = ObRs("Nbridge")
GPU = ObRs("GPU")
Notas = ObRs("Notas")
Reballing = ObRs("Reballing")
FechaProgramada = ObRs("FechaProgramada")

ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Equipos Where Id = " & IdEquipo
ObRs.Open SQL, ObConn
IdMarca = ObRs("Marca")
IdTiposdeEquipos = ObRs("Tipo")
Modelo = ObRs("Modelo")
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Marcas Where Id = " & IdMarca
ObRs.Open SQL, ObConn
Marca = ObRs("Marca")
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From TiposdeEquipos Where Id = " & IdTiposdeEquipos
ObRs.Open SQL, ObConn
TipodeEquipo = ObRs("Tipo")
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & IdCliente
ObRs.Open SQL, ObConn
Nombre = ObRs("Nombre")
Direccion = ObRs("Direccion")
Telefono = ObRs("Telefono")
Email = ObRs("Email")
TipoCliente = ObRs("TipoCliente")
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Estados Where Id = " & IdEstado
ObRs.Open SQL, ObConn
Estado = ObRs("Estado")
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios Where Id = " & IdUsuarioIngreso
ObRs.Open SQL, ObConn
UsuarioIngreso = ObRs("Nick")
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios Where Id = " & IdUsuarioEstado
ObRs.Open SQL, ObConn
UsuarioEstado = ObRs("Nick")
ObRs.Close
ObConn.Close
%>
<%
If TipoTarjeta < 1 then TipoTarjeta = 1
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From TipoTarjeta Where Id = " & TipoTarjeta
ObRs.Open SQL, ObConn
NombreTipoTarjeta = ObRs("TipoTarjeta")
ObRs.Close
ObConn.Close
%>

<%
Select Case TipoCliente
Case 2
%>
<body bgcolor="#FFFFAA">
<%
Case 4
%>
<body bgcolor="#FFCC66">
<%
Case 5
%>
<body bgcolor="#99FF66">
<%
Case else
%>
<body bgcolor="#FFFFFF">
<%
End Select
%>

<div style="position: absolute; width: 295px; height: 45px; z-index: 5; left: 406px; top: 9px; right: 88px;" id="TecnicoAsignado">
	<span lang="es-ar">Tecnico Asignado</span><span class="style1"> </span><br>
	<form method="post" action="">
		<span class="style1">
		<table style="width: 100%">
			<tr>
				<td>
		<span class="style1">
	<select name="TecnicoAsignado" onchange="tecnicoasignado();" style="width: 125px">
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
</select><br>
		<input name="Enviar" type="submit" value="Enviar"></span></td>
				</span>
				<td><span lang="es-ar" class="style3"><strong>
				<span class="style4">Fecha Programada<br>
				<%Response.Write FechaProgramada%><br>
				</span></strong><strong><span class="style4">Dias restantes<br>
				<%
				diashabiles = 0
				i = date
				do while i <= FechaProgramada - 1
					i = i+1
					j = DatePart("w", i)
					if j <> 1 and j <> 7 then
						diashabiles = diashabiles + 1
					end if
				loop  
				response.write diashabiles
				%>
<select name="DiasProgramados">
<%
for i = 0 to 15
if i = diashabiles then
%>
<option selected value="<%Response.Write i%>"><%Response.Write i%></option>
<%
else
%>
<option value="<%Response.Write i%>"><%Response.Write i%></option>
<%
end if
next
%>
</select>
				</span></strong></span></td>
			</tr>
		</table>
		<span class="style1">
		<br>
		</span></form>
</div>
<%
ObRs.Close
ObConn.Close
%>


<div style="position: absolute; width: 356px; height: 100px; z-index: 3; left: 495px; top: 31px; visibility: hidden" id="ReballEdit">
	<table border="1" bgcolor="#00FFFF" bordercolor="#FF0000" width="359">
		<tr>
			<td>
			<form method="POST" action="RebalingEdit.asp?Id=<%Response.Write IdOrden%>">
			<b><font face="Arial">Modelo del equipo:</font></b><p>
				<input type="text" name="Modelo" size="20" value="<%Response.Write ModeloR%>"></p>
			<p><b><font face="Arial">Northbridge</font></b></p>
				<p><input type="text" name="Nbridge" size="20" value="<%Response.Write Nbridge%>"></p>
			<p><b><font face="Arial">GPU</font></b></p>
				<p><input type="text" name="GPU" size="20" value="<%Response.Write GPU%>"></p>
			<p><b><font face="Arial">Notas</font></b></p>
				<p><input type="text" name="Notas" size="46" value="<%Response.Write Notas%>"></p>
<%
if Reballing = true then
 Checkedsi = "checked"
 Checkedno = ""
else
 Checkedsi = ""
 Checkedno = "checked"
end if
%>
				<p><font face="Arial"><b>Reballing: <font color="#00CC00">SI<input type="radio" value="si" name="Reballing" <%Response.Write Checkedsi%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font> 
				<font color="#FF0000">NO</font><font color="#00CC00"><input type="radio" name="Reballing" value="no" <%Response.Write Checkedno%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></b></font>
				<input type="submit" value="Enviar" name="B8"></p>
			</form>
			</td>
		</tr>
	</table>
</div>
<div style="position: absolute; width: 356px; height: 401px; z-index: 2; left: 445px; top: 14px; visibility: hidden" id="Reball">
	<table border="1" bgcolor="#00FFFF" bordercolor="#FF0000" width="359" height="432" class="style2">
		<tr>
			<td><b><font face="Arial">Modelo del equipo:</font></b><table border="1" bgcolor="#FFFFFF" bordercolor="#000000" width="322" height="28">
				<tr>
					<td><font face="Arial"><%Response.Write ModeloR%></font></td>
				</tr>
			</table>
			<p><b><font face="Arial">Northbridge</font></b></p>
			<table border="1" bgcolor="#FFFFFF" bordercolor="#000000" width="242" height="27">
				<tr>
					<td><font face="Arial"><%Response.Write Nbridge%></font></td>
				</tr>
			</table>
			<p><b><font face="Arial">GPU</font></b></p>
			<table border="1" bgcolor="#FFFFFF" bordercolor="#000000" width="243" height="29">
				<tr>
					<td><font face="Arial"><%Response.Write GPU%></font></td>
				</tr>
			</table>
			<p><b><font face="Arial">Notas</font></b></p>
			<div align="left">
			<table border="1" bgcolor="#FFFFFF" bordercolor="#000000" width="325" height="81">
				<tr>
					<td align="left" valign="top"><font face="Arial"><%Response.Write Notas%></font></td>
				</tr>
			</table>
			</div>
			<p><font face="Arial"><b>Reballing:
			<%if Reballing = true then%>
			<font color="#00CC00">SI</font>
			<%else%>
			<font color="#CC0000">NO</font>
			<%end if%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			<a href="#" onclick="Reballiedit();">Editar</a>&nbsp;&nbsp; <a href="#" onclick="Reballi();">
			Cerrar</a></b></font></td>
		</tr>
	</table>
</div>
<div style="position: absolute; width: 194px; height: 106px; z-index: 1; left: 320px; top: 174px; visibility: hidden" id="Serial">
	<table border="1" width="100%">
		<tr>
			<td bgcolor="#FF0000"><font face="Arial"><b>Cambiar Nº de Serie</b></font><form method="POST" action="CambioSerie.asp?Id=<%Response.Write IdOrden%>">
		<p><input type="text" name="Serie" size="20" value="<%Response.Write Serie%>"></p>
		<p><input type="submit" value="Enviar" name="B7"></p>
	</form>
			<p>&nbsp;</td>
		</tr>
	</table>
</div>

<div style="position: absolute; width: 240px; height: 120px; z-index: 4; left: 437px; top: 358px; visibility: hidden" id="Referencia">
				<table border="1" width="100%">
				<tr>
				<td bgcolor="#008000"><font face="Arial"><b>Cambiar Nº de 
				Referencias </b></font>
				<form method="POST" action="CambioReferencia.asp?Id=<%Response.Write IdOrden%>">
					<p>
					<input type="text" name="Referencia" size="20" value="<%Response.Write Referencia%>"></p>
					<p><input type="submit" value="Enviar" name="B9"></p>
				</form>
				</td>
				</tr>
				</table>
</div>

		
<form method="GET" action="ConsultaDeOrden.asp" webbot-action="--WEBBOT-SELF--">
	<p><font face="Arial"><input type="text" name="Id" size="20"><font size="2"><br>
	</font>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></font></p>
</form>
<form method="POST" action="ConsultaDeOrden.asp" webbot-action="--WEBBOT-SELF--">
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="489">
      <tr>
        <td width="100%" colspan="2" height="21"><font face="Arial"><font size="2">Orden:</font><b><font color="#008000" size="5"><a href="ConsultaDeOrden.asp?Id=<%Response.Write IdOrden - 1%>"><span style="text-decoration: none"><font color="#008000">&lt;</font></span></a></font><font color="#008000" size="4"> <%Response.Write IdOrden%>
    	</font> 
    	<font color="#008000" size="5">
    	<a href="ConsultaDeOrden.asp?Id=<%Response.Write IdOrden + 1%>">
		<span style="text-decoration: none"><font color="#008000">&gt;</font></span></a></font></b> 
    <font size="2">&nbsp;&nbsp;&nbsp; <a href="fpdf/Orden.asp?Id=<%Response.Write IdOrden%>">Imprimir</a>
		<a href="http://server/sistema/Calendario/calendar.asp?action=addevent&date=<%Response.Write Year(DATE) & "-" & Month(DATE) & "-" & Day(DATE)%>">
		Agendar</a></font></font>
		</td>
      </tr>
      <tr>
        <td width="100%" colspan="2" height="86"><font face="Arial" size="2"><b>Datos del Cliente: </b>
	<a href="ModificarCliente.asp?IdCliente=<%Response.Write IdCliente%>">Modificar Cliente</a><span lang="es">
        <b>
        <a href="ListaOrdenEstado.asp?Estado=&Cliente=<%Response.Write Nombre%>&solonote=note">Otros equipos del cliente</a></b></span><br>
	<b>Nombre:</b> <%Response.Write Nombre%>&nbsp;&nbsp;
	<b>&nbsp;&nbsp;&nbsp;&nbsp; Dirección:</b> <%Response.Write Direccion%><br>
	<b>Teléfono:</b> <%Response.Write Telefono%>
	<b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Email:</b> <%Response.Write Email%></font><p>
		<a href="cambioclienteorden.asp?Orden=<%Response.Write IdOrden%>">.</a></td>
      </tr>
      <tr>
        <td width="63%" height="88"><font face="Arial" size="2"><b>Datos del Equipo:</b><br>
	<b>Tipo:</b> <%Response.Write TipodeEquipo%>&nbsp;&nbsp;
	<b>Marca: </b> <%Response.Write Marca%>&nbsp;&nbsp;&nbsp;
	<b>Modelo:</b> <%Response.Write Modelo%>
	
	<span lang="es">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </span></font>
		
		<p><font face="Arial" size="2">
	    <br>
	
	<b>Serie:</b> <%Response.Write Serie%>&nbsp;<a href="#" onclick="Serial1();">.</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <b> <a href="ListaOrdenSerie.asp?Serie=<%Response.Write Serie%>">Otros 
    Ingresos del Equipo</a></b></font></p>
		<p>
        <a href="CambiaEquipoOrden.asp?Orden=<%Response.Write IdOrden%>&Equipo=<%Response.Write IdEquipo%>">.</a></td>
        <td width="37%" height="88"><font face="Arial" size="2">
		<a title="Agregar Imagen al Equipo" href="ImagenEquipo.asp?equipo=<%Response.Write IdEquipo%>">
	<img border="0" src="imagen.asp?Id=<%Response.Write IdEquipo%>&Tabla=Equipos" height="86" width="91"></a><img border="0" src="imagen.asp?Id=<%Response.Write IdMarca%>&Tabla=Marcas" height="86" width="95"></font></td>
      </tr>
      <tr>
        <td width="63%" height="54"><font face="Arial" size="2"><b>Fecha de Ingreso:</b> <%Response.Write FechaIngreso%>&nbsp;&nbsp;&nbsp; 
	<b>Usuario de Ingreso:</b> <%Response.Write UsuarioIngreso%><br>
	<b>Estado:</b> <%Response.Write Estado%>&nbsp;&nbsp;&nbsp;
		<font color="#005500"<%
if Estado <> "Con Tarjeta" then
Response.Write " style='visibility:hidden'"
end if
		%>><b><%Response.Write NombreTipoTarjeta%></b></font>&nbsp;&nbsp;&nbsp;
	<b>Fecha de Estado:</b> <%Response.Write FechaEstado%> <b>
	<a target="_parent" href="ModificarEstadoOrden.asp?Id=<%Response.Write IdOrden%>">
	Modificar</a></b></font></td>
        <td width="37%" height="54">
<div id="Cheque" <%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select Estado from Ordenes where Id = " & Request.QueryString ("Id"),ObConn
if ObRs ("Estado") <> 28 then
%>style="visibility:hidden"<%
end if
ObRs.Close
ObConn.Close
%>>
<font face="Arial"><b><font size="2">Banco: 
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "SELECT Cheque.Id, BancoCheque.BancoCheque, Cheque.NCuenta, Cheque.NCheque,"
SQL = SQL & " Cheque.Fecha, Cheque.Cruzado, Cheque.anulado "
SQL = SQL & "FROM BancoCheque INNER JOIN Cheque ON BancoCheque.Id = Cheque.Banco"

SQL = SQL & " Where Orden = " & Request.QueryString("Id")
ObRs.Open SQL,ObConn
if not ObRs.Eof then
Banco = ObRs ("BancoCheque")
NCuenta = ObRs ("NCuenta")
NCheque = ObRs ("NCheque")
Fecha = ObRs ("Fecha")
Cruzado = ObRs ("Cruzado")
else
Banco = 1
end if
ObRs.Close
ObConn.Close
%>&nbsp;<%Response.Write Banco%><br>N° Cuenta:&nbsp;
<%Response.Write NCuenta%><br>N° Cheque:&nbsp;<%Response.Write NCheque%><br>
Fecha de Cobro:&nbsp;<%Response.Write Fecha%><br>
Cruzado</font></b><input type="checkbox" name="Cruzado" value="true"  onclick="return false;" style="font-weight: 700"<%
if Cruzado = True then
Response.Write "checked"
end if
%> ><b><font size="2"> </font>
</b></font>
</div></td>
      </tr>
      <tr>
        <td width="63%" height="236">
	<p><font face="Arial" size="2"><b>Técnico:</b> <%Response.Write UsuarioEstado%><b><Br>
	Accesorios:</b> <%Response.Write Accesorios%>.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<b>N° de Referencia:</b> <%Response.Write Referencia%> <a href="#" onclick="Referencia1();">.</a><br>
	<b>Observaciones de Ingreso:</b></font>
	<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="47">
      <tr>
        <td width="100%" bgcolor="#FFFFFF" height="47" align="left" valign="top">&nbsp;<%Response.Write ObservacionIngreso%></td>
      </tr>
    </table></p>
	<p><font face="Arial" size="2"><b>Observaciones del Técnico:</b> 
	<a target="_parent" href="ModificarObservacioTecnico.asp?Id=<%Response.Write IdOrden%>">
	<b>
	Modificar</b></a></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<font face="Arial" size="2"><a href="#" onclick="Reballi();">Reballing</a></font><table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="45">
          <tr>
            <td width="100%" bgcolor="#FFFFFF" height="45" align="left" valign="top">&nbsp;<%Response.Write ObservacionTecnico%></td>
          </tr>
    </table></p>
        <p></td>
        <td width="37%" height="236">
	<font face="Arial" size="2"><b>Observaciones Internas<br>(no comunicar al cliente)
        </b></font>
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="68%" height="45">
          <tr>
            <td width="100%" bgcolor="#FFFFFF" height="45" align="left" valign="top">&nbsp;<%Response.Write ObservacionInterna%></td>
          </tr>
    </table></td>
      </tr>
    </table>
	<p><font face="Arial" size="2"><br>
	</font></p>
</form>

<form method="GET" action="IngresoOrdenItem.asp">
	<table border="1" width="100%" id="table1" cellspacing="1" bordercolor="#000000" style="border-collapse: collapse">
		<tr>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn
%><font face="Arial" size="2"> </font>

			<td width="51"><font face="Arial" size="2"><a href="ModificarEquipo.asp?IdEquipo=<%Response.Write ObRs ("Equipo")%>">Listado</a></td>
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
SQL = SQL & " Where Orden = " & Request.QueryString("Id")

ObRs.Open SQL,ObConn
Total = 0
DO WHILE NOT ObRs.Eof
%> </font>
		<tr>
			<td width="51">
			<a href="EliminarOrdenItem.asp?Id=<%Response.Write ObRs ("Id")%>">
            <img border="0" src="images/Delete.GIF" width="26" height="27" alt="Eliminar Item"></a></td>
			<td width="51"><%Response.Write ObRs ("Codigo")%>&nbsp;</td>
			<td width="446"><%Response.Write ObRs ("Descripcion")%>&nbsp;<a href="ModificarOrdenItem.asp?Id=<%Response.Write ObRs ("Id")%>"><img border="0" src="images/Editar.gif" width="24" height="23"></a></td>
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
	<p><font face="Arial"><select size="1" name="Id">
	<option value="<%Response.Write Request.QueryString("Id")%>">
	<%Response.Write Request.QueryString("Id")%></option>
	</select><input type="submit" value="Submit" name="B5"><font size="2">
    </font>
	<input type="reset" value="Reset" name="B6"></font></p>
</form>

</body>

</html>