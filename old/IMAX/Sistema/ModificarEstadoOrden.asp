<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>
<script language="JavaScript" type="text/javascript">

function Cambio(option)
{
if ( option.value == "27" )
        {
	document.getElementById("Campo").style.visibility = "visible";
	document.getElementById("imagen").style.visibility = "visible";
        }
else
        {
	document.getElementById("Campo").style.visibility = "hidden";
	document.getElementById("imagen").style.visibility = "hidden";
        }
if ( option.value == "28" )
        {
	document.getElementById("Cheque").style.visibility = "visible";
        }
else
        {
	document.getElementById("Cheque").style.visibility = "hidden";
        }

return
}

</script>
<body>
<div>
<%
IF Request.Form = "" THEN
%>

<b>Modificar Estado de la Orden</b><form method="POST" action="ModificarEstadoOrden.asp" webbot-action="--WEBBOT-SELF--">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.QueryString("Id")
ObRs.Open SQL,ObConn
%>
	<p><b>Orden: <select size="1" name="Id">
	<option selected value="<%Response.Write ObRs ("Id")%>">
	<%Response.Write ObRs ("Id")%></option>
	</select></b></p>
<%
UsuarioEstado = ObRs ("UsuarioEstado")
Estado = ObRs ("Estado")
TipoTarjeta = ObRs ("TipoTarjeta")
ObRs.Close
ObConn.Close
%>
<p><b>Técnico: <select size="1" name="UsuarioEstado">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Usuarios  Where habilitado = true"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
IF UsuarioEstado = ObRs ("Id") Then
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
%>
</select></b><Br>
<b>Estado: <select id="estado" onChange="Cambio(this)" size="1" name="Estado">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Estados",ObConn
DO WHILE NOT ObRs.Eof
IF Estado = ObRs ("Id") Then
%>
<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Estado")%></option>
<%
Else
%>
<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Estado")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>

</select></b>
<Select type="text" name="TipoTarjeta" id="Campo" <%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select Estado from Ordenes where Id = " & Request.QueryString ("Id"),ObConn
if ObRs ("Estado") <> 27 then
%>style="visibility:hidden"<%
end if
ObRs.Close
ObConn.Close
%>>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select Id, TipoTarjeta from TipoTarjeta Order By TipoTarjeta"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
IF TipoTarjeta = ObRs ("Id") Then
%>
<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("TipoTarjeta")%></option>
<%
Else
%>
<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("TipoTarjeta")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>


</select> <a href="IngresoTipoTarjeta.asp"> <img id= "imagen" border="0" src="images/Editar.gif" width="26" height="26" <%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Select Estado from Ordenes where Id = " & Request.QueryString ("Id"),ObConn
if ObRs ("Estado") <> 27 then
%>style="visibility:hidden"<%
end if
ObRs.Close
ObConn.Close
%>></a>

<Br>
<span lang="es"><b>Fecha:</b> 
<input type="text" name="FechaEstado" size="13" value="<%Response.Write Date%>" /><Br>

</span><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
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
<p><b>Banco: </b>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Cheque Where Orden = " & Request.QueryString("Id")
ObRs.Open SQL,ObConn
if not ObRs.Eof then
Banco = ObRs ("Banco")
NCuenta = ObRs ("NCuenta")
NCheque = ObRs ("NCheque")
Fecha = ObRs ("Fecha")
Cruzado = ObRs ("Cruzado")
else
Banco = 1
end if
ObRs.Close
ObConn.Close
%>

<select size="1" name="Banco">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From BancoCheque Order By BancoCheque"
ObRs.Open SQL,ObConn
DO WHILE NOT ObRs.Eof
IF Banco = ObRs ("Id") Then
%>
<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("BancoCheque")%></option>
<%
ELSE
%>
<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("BancoCheque")%></option>
<%
END IF
ObRs.MoveNext
LOOP
ObRs.Close
ObConn.Close
%>
</select>
<b><a href="IngresoBanco.asp"><img border="0" src="images/Editar.gif" width="31" height="27"></a><br>N° Cuenta:
</b>
<input name="NCuenta" size="20" value="<%Response.Write NCuenta%>"><b><br>N° Cheque: </b> 
<input name="NCheque" size="20" value="<%Response.Write NCheque%>"><b><br>
Fecha de Cobro: </b> 
<input name="Fecha" size="16" value="<%Response.Write Fecha%>"><b><br>
Cruzado</b><input type="checkbox" name="Cruzado" <%
if Cruzado then
Response.Write "checked"
end if
%> value="On">
</div></form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Request.Form("Id")
ObRs.Open SQL, ObConn, 3, 3

'************************************************************************************************  Estado En Cuenta
If ObRs ("Estado") <> 11 and Request.Form ("Estado") = 11 Then
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & ObRs ("Cliente")
ObRs2.Open SQL,ObConn, 3, 3

SET ObConn3 = Server.CreateObject ("ADODB.Connection")
SET ObRs3 = Server.CreateObject ("ADODB.RecordSet")
ObConn3.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & Request.Form("Id")
ObRs3.Open SQL,ObConn
DO WHILE NOT ObRs3.Eof
Total = Total + ObRs3 ("Cantidad") * ObRs3 ("PrecioUnitario")
ObRs3.MoveNext
LOOP
ObRs3.Close
ObConn3.Close


ObRs2 ("Cuenta") = ObRs2 ("Cuenta") + Total
ObRs2.Update

ObRs2.Close
ObConn2.Close
End If

If ObRs ("Estado") = 11 and Request.Form ("Estado") <> 11 Then
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & ObRs ("Cliente")
ObRs2.Open SQL,ObConn, 3, 3

SET ObConn3 = Server.CreateObject ("ADODB.Connection")
SET ObRs3 = Server.CreateObject ("ADODB.RecordSet")
ObConn3.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & Request.Form("Id")
ObRs3.Open SQL,ObConn
DO WHILE NOT ObRs3.Eof
Total = Total + ObRs3 ("Cantidad") * ObRs3 ("PrecioUnitario")
ObRs3.MoveNext
LOOP
ObRs3.Close
ObConn3.Close

ObRs2 ("Cuenta") = ObRs2 ("Cuenta") - Total
ObRs2.Update

ObRs2.Close
ObConn2.Close
End If
'************************************************************************************************  Estado En Cuenta

'************************************************************************************************  Estado Pago a Cuenta
If ObRs ("Estado") <> 24 and Request.Form ("Estado") = 24 Then
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & ObRs ("Cliente")
ObRs2.Open SQL,ObConn, 3, 3

SET ObConn3 = Server.CreateObject ("ADODB.Connection")
SET ObRs3 = Server.CreateObject ("ADODB.RecordSet")
ObConn3.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & Request.Form("Id")
ObRs3.Open SQL,ObConn
DO WHILE NOT ObRs3.Eof
Total = Total + ObRs3 ("Cantidad") * ObRs3 ("PrecioUnitario")
ObRs3.MoveNext
LOOP
ObRs3.Close
ObConn3.Close


ObRs2 ("Cuenta") = ObRs2 ("Cuenta") + Total
ObRs2.Update

ObRs2.Close
ObConn2.Close
End If

If ObRs ("Estado") = 24 and Request.Form ("Estado") <> 24 Then
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & ObRs ("Cliente")
ObRs2.Open SQL,ObConn, 3, 3

SET ObConn3 = Server.CreateObject ("ADODB.Connection")
SET ObRs3 = Server.CreateObject ("ADODB.RecordSet")
ObConn3.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & Request.Form("Id")
ObRs3.Open SQL,ObConn
DO WHILE NOT ObRs3.Eof
Total = Total + ObRs3 ("Cantidad") * ObRs3 ("PrecioUnitario")
ObRs3.MoveNext
LOOP
ObRs3.Close
ObConn3.Close

ObRs2 ("Cuenta") = ObRs2 ("Cuenta") - Total
ObRs2.Update

ObRs2.Close
ObConn2.Close
End If

'************************************************************************************************  Estado Pago a Cuenta
'************************************************************************************************  Estado Cobro a Cuenta
If ObRs ("Estado") <> 25 and Request.Form ("Estado") = 25 Then
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & ObRs ("Cliente")
ObRs2.Open SQL,ObConn, 3, 3

SET ObConn3 = Server.CreateObject ("ADODB.Connection")
SET ObRs3 = Server.CreateObject ("ADODB.RecordSet")
ObConn3.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & Request.Form("Id")
ObRs3.Open SQL,ObConn
DO WHILE NOT ObRs3.Eof
Total = Total + ObRs3 ("Cantidad") * ObRs3 ("PrecioUnitario")
ObRs3.MoveNext
LOOP
ObRs3.Close
ObConn3.Close


ObRs2 ("Cuenta") = ObRs2 ("Cuenta") - Total
ObRs2.Update

ObRs2.Close
ObConn2.Close
End If

If ObRs ("Estado") = 25 and Request.Form ("Estado") <> 25 Then
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & ObRs ("Cliente")
ObRs2.Open SQL,ObConn, 3, 3

SET ObConn3 = Server.CreateObject ("ADODB.Connection")
SET ObRs3 = Server.CreateObject ("ADODB.RecordSet")
ObConn3.Open "Sistema"
SQL = "Select * From ConsultaOrdenItem Where Orden = " & Request.Form("Id")
ObRs3.Open SQL,ObConn
DO WHILE NOT ObRs3.Eof
Total = Total + ObRs3 ("Cantidad") * ObRs3 ("PrecioUnitario")
ObRs3.MoveNext
LOOP
ObRs3.Close
ObConn3.Close

ObRs2 ("Cuenta") = ObRs2 ("Cuenta") + Total
ObRs2.Update

ObRs2.Close
ObConn2.Close
End If

'************************************************************************************************  Estado Cobro a Cuenta
'************************************************************************************************  Estado Cobro a Cuenta
If Request.Form ("Estado") = 28 Then
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From Cheque Where Orden = " & Request.Form("Id")
ObRs2.Open SQL,ObConn, 3, 3

if ObRs2.Eof then
ObRs2.AddNew
end if

ObRs2 ("Banco") = Request.Form ("Banco")
ObRs2 ("NCuenta") = Request.Form ("Ncuenta")
ObRs2 ("NCheque") = Request.Form ("NCheque")
ObRs2 ("Fecha") = Request.Form ("Fecha")
if Request.Form ("Cruzado") = "On" Then
ObRs2 ("Cruzado") = true
else
ObRs2 ("Cruzado") = false
end if
ObRs2 ("Orden") = Request.Form("Id")
ObRs2 ("Anulado") = False

ObRs2.Update
ObRs2.Close
ObConn2.Close
End If
'************************************************************************************************  Estado Cobro a Cuenta
'************************************************************************************************  Estado Cobro a Cuenta
If ObRs ("Estado") = 28 and (Request.Form ("Estado") <> 6 or Request.Form ("Estado") <> 9) Then
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
SQL = "Select * From Cheque Where Orden = " & Request.Form("Id")
ObRs2.Open SQL,ObConn, 3, 3
ObRs2 ("Anulado") = True

ObRs2.Update
ObRs2.Close
ObConn2.Close
End If
'************************************************************************************************  Estado Cobro a Cuenta

ObRs ("Estado") = Request.Form ("Estado")
ObRs ("UsuarioEstado") = Request.Form ("UsuarioEstado")
ObRs ("FechaEstado") = Request.Form ("FechaEstado")
ObRs ("TipoTarjeta") = Request.Form ("TipoTarjeta")
ObRs.Update

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
response.redirect "ConsultaDeOrden.asp?Id="& Request.Form("Id")
END IF
%>

</body>
</html>