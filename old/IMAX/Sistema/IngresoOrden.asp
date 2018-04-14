<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Ingreso de Orden</title>
</head>
<script language="JavaScript">

function muestra_oculta(id){

var el = document.getElementById('contenido_a_mostrar'); //se define la variable "el" igual a nuestro div

if(id == 0 ){
el.style.display = 'none'; //damos un atributo display:none que oculta el div
            }
else {
el.style.display = 'block'; //damos un atributo display:none que oculta el div
}
							}


window.onload = function(){/*hace que se cargue la función lo que predetermina que div estará oculto hasta llamar a la función nuevamente*/

var el = document.getElementById('contenido_a_mostrar'); //se define la variable "el" igual a nuestro div
el.style.display = 'none'; //damos un atributo display:none que oculta el div

}
</script>


<%
Cliente = Request.QueryString ("Cliente")
Equipo = Request.QueryString ("Equipo")
%>
<%
Nombre = ""
Direccion = ""
Telefono = ""
Email = ""
If Cliente <> "" Then
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL="Select * From Clientes Where Id=" & Cliente
ObRs.Open SQL,ObConn
Nombre = ObRs("Nombre")
Direccion = ObRs("Direccion")
Telefono = ObRs("Telefono")
Email = ObRs("Email")
ObRs.Close
ObConn.Close
End If

Marca = ""
TipoEquipo = ""
Modelo = ""
If Equipo <> "" Then
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL="Select * From Equipos Where Id=" & Equipo
ObRs.Open SQL,ObConn
Marca = ObRs("Marca")
TipoEquipo = ObRs("Tipo")
Modelo = ObRs("Modelo")
ObRs.Close
ObConn.Close
End If

If Marca <> "" Then
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL="Select * From Marcas Where Id=" & Marca
ObRs.Open SQL,ObConn
Marca = ObRs("Marca")
ObRs.Close
ObConn.Close
End If

If TipoEquipo <> "" Then
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL="Select * From TiposDeEquipos Where Id=" & TipoEquipo
ObRs.Open SQL,ObConn
TipoEquipo = ObRs("Tipo")
ObRs.Close
ObConn.Close
End If
%>

<body>
<%
IF Request.Form = "" THEN
%>

<p>Ingreso de Orden</p>
<form method="POST" action="IngresoOrden.asp?Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>">
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="89%" id="AutoNumber1" height="137">
    <tr>
      <td width="100%" colspan="4" height="19">Usuario de Ingreso:<select size="1" name="UsuarioIngreso">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * from Usuarios Where habilitado = true"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
%>
		<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Nick")%></option>
<%
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
		</select></td>
    </tr>
    <tr>
      <td width="100%" colspan="4" height="19">Cliente: 
		<a href="SeleccionarCliente.asp?Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Pagina=<%Response.Write "IngresoOrden.asp"%>">Seleccionar</a></td>
    </tr>
    <tr>
      <td width="55%" colspan="3" height="19">Nombre: <%Response.Write Nombre%></td>
      <td width="43%" height="19">Dirección: <%Response.Write Direccion%></td>
    </tr>
    <tr>
      <td width="100%" colspan="4" height="19">Teléfono: <%Response.Write Telefono%></td>
    </tr>
    <tr>
      <td width="100%" height="19" colspan="4">Email: <%Response.Write Email%></td>
    </tr>
    <tr>
      <td width="100%" height="19" colspan="4">Equipo 
		<a href="SeleccionarEquipo.asp?Cliente=<%Response.Write Cliente%>&Equipo=<%Response.Write Equipo%>&Pagina=<%Response.Write "IngresoOrden.asp"%>">Seleccionar</a></td>
    </tr>
    <tr>
      <td width="39%" height="17">Tipo de Equipo <%Response.Write TipoEquipo%></td>
      <td width="59%" colspan="3" height="17">Marca: <%Response.Write Marca%></td>
    </tr>
    <tr>
      <td width="100%" height="19" colspan="4">Modelo: <%Response.Write Modelo%></td>
    </tr>
    <tr>
      <td width="100%" colspan="4" height="19">Serie:<input type="text" name="Serie" size="29"></td>
    </tr>
    <tr>
      <td width="100%" colspan="4" height="19">Accesorios:</td>
    </tr>
    <tr>
      <td width="50%" colspan="2" height="19">
      <input type="text" name="Accesorios" size="76">&nbsp;&nbsp;&nbsp;&nbsp; </td>
      <td width="50%" colspan="2" height="19">
      N° de Referencia: <input type="text" name="Referencia" size="20"></td>
    </tr>
    <tr>
      <td width="100%" colspan="4" height="19">Observaciones</td>
    </tr>
    <tr>
      <td height="19" style="width: 0%">
      <textarea rows="4" name="ObservacionIngreso" style="width: 290px"></textarea></td>
      <td width="100%" height="19" style="width: 50%">

<span lang="es-ar">Tecnico Asignado / dias programados<br>
</span>

<select id="tecnicoasignado" name="TecnicoAsignado" onchange="muestra_oculta(this.options[this.selectedIndex].value)" style="width: 125px">
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

%>
<option selected value="0">Elegir un Tecnico</option>

</select>
      
      <select name="DiasProgramados">
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7" selected="">7</option>
<option value="8">8</option>
<option value="9">9</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
<option value="13">13</option>
<option value="14">14</option>
<option value="15">15</option>
</select></td>
      <td width="100%" colspan="2" height="19" style="width: 50%">
      &nbsp;</td>
    </tr>
  </table>
  <div id="contenido_a_mostrar">
  <p><input type="submit" value="Enviar" name="B1"></p>
  </div>

</form>

<%

ELSE

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

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Ordenes",ObConn, 3, 3

ObRs.AddNew
ObRs ("Cliente") = Cliente
ObRs ("Equipo") = Equipo
ObRs ("Serie") = UCase(Request.Form ("Serie"))
ObRs ("Estado") = 1
ObRs ("Accesorios") = Request.Form ("Accesorios")
ObRs ("UsuarioIngreso") = Request.Form ("UsuarioIngreso")
ObRs ("UsuarioEstado") = Request.Form ("UsuarioIngreso")
ObRs ("FechaIngreso") = DATE
ObRs ("FechaEstado") = DATE
ObRs ("ObservacionIngreso") = Request.Form ("ObservacionIngreso")
ObRs ("Referencia") = Request.Form ("Referencia")
ObRs ("TecnicoAsignado") = Request.Form ("TecnicoAsignado")
ObRs ("FechaProgramada") = fechaprogramada

ObRs.Update
ObRs.Close
ObConn.Close
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Order By Id"
ObRs.Open SQL, ObConn
DO WHILE NOT ObRs.Eof
Orden = ObRs("Id")
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b><p><b>
<a target="_blank" href="fpdf/Orden.asp?Id=<%Response.Write orden%>">Imprimir</a></b>
</p>
<b>
<a href="ConsultaDeOrden.asp?Id=<%Response.Write orden%>">Consultar la Orden</a> </b>
<%
END IF
%>
</body>

</html>