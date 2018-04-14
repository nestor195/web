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
<b>Ingreso de Cliente</b><form method="POST" action="IngresoCliente%20-%20old.asp" webbot-action="--WEBBOT-SELF--">
	<p>Nombre: <input type="text" name="Nombre" size="30"><br>
	Dirección: <input type="text" name="Direccion" size="37"><br>
	Teléfono: <input type="text" name="Telefono" size="20"><br>
	Email: <input type="text" name="Email" size="20"><br>
	Tipo de Usuario: <select size="1" name="TipoCliente">
    <option selected value="0">Usuario Final</option>
    <option value="1">Gremio</option>
    <option value="2">Operario IMAX</option>
    </select><br>
	Obaservaciones:<br>
	&nbsp;<textarea rows="4" name="Observaciones" cols="39"></textarea><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Clientes",ObConn, 3, 3

ObRs.AddNew
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