<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
<%
FUNCTION Sinespacios (palabras)
	Final = ""
	FOR I = 0 TO Len (palabras)
		IF Mid (palabras, I + 1, 1) = " " THEN
			letra = "+"
		ELSE
			letra = Mid (palabras, I + 1, 1)
		END IF
		Final = Final & letra
	NEXT
	Sinespacios = Final
END FUNCTION
%>

<SCRIPT language=javascript type="text/Jscript">
b=""
a=opener.document.location.href
for (i=0; i<a.length; i++) {
	if (a.charAt(i)!="?"){
		b=b+a.charAt(i);
		}
	else{
		i=a.length
		}
	}

window.opener.document.location=b+'?Cliente=<%Response.Write(Sinespacios(Request.Form ("Nombre")))%>'
</SCRIPT>
</head>

<body>
<%
IF Request.Form = "" THEN
%>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & Request.QueryString("IdCliente")
ObRs.Open SQL, ObConn
%>

<b>Modificar datos del Cliente</b><form method="POST" action="ModificarCliente.asp" webbot-action="--WEBBOT-SELF--">
	<p>Id: <select size="1" name="Id">
    <option value="<%Response.Write ObRs("Id")%>" selected><%Response.Write ObRs("Id")%></option>
    </select><br>
	Nombre: <input type="text" name="Nombre" size="30" value="<%Response.Write ObRs("Nombre")%>"><br>	
	Dirección: <input type="text" name="Direccion" size="37" value="<%Response.Write ObRs("Direccion")%>"><br>
	Teléfono: <input type="text" name="Telefono" size="20" value="<%Response.Write ObRs("Telefono")%>"><br>
	Email: <input type="text" name="Email" size="20" value="<%Response.Write ObRs("Email")%>"><br>
	Tipo de Usuario: 
	<select size="1" name="TipoCliente">
<%
SET ObConn2 = Server.CreateObject ("ADODB.Connection")
SET ObRs2 = Server.CreateObject ("ADODB.RecordSet")
ObConn2.Open "Sistema"
ObRs2.Open "TipoCliente",ObConn2
DO WHILE NOT ObRs2.Eof
if ObRs("TipoCliente") = ObRs2 ("Id") then
%>
		<option selected value="<%Response.Write ObRs2 ("Id")%>"><%Response.Write ObRs2 ("TipoCliente")%></option>
<%
else
%>
		<option value="<%Response.Write ObRs2 ("Id")%>"><%Response.Write ObRs2 ("TipoCliente")%></option>
<%
end if
ObRs2.MoveNext
LOOP
ObRs2.Close
ObConn2.Close
%>
		</select><br>
	Obaservaciones:<br>
	&nbsp;<textarea rows="4" name="Observaciones" cols="39"><%Response.Write ObRs("Observaciones")%></textarea><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ObRs.Close
ObConn.Close
%>

<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Clientes Where Id = " & Request.Form ("Id")
ObRs.Open SQL,ObConn, 3, 3

ObRs ("Nombre") = Request.Form ("Nombre")
ObRs ("Direccion") = Request.Form ("Direccion")
ObRs ("Telefono") = Request.Form ("Telefono")
ObRs ("Email") = Request.Form ("Email")
ObRs ("Observaciones") = Request.Form ("Observaciones")
ObRs ("TipoCliente") = Request.Form ("TipoCliente")
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