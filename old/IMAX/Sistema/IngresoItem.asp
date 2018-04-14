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
<b>Ingreso de Item</b><form method="POST" action="IngresoItem.asp" webbot-action="--WEBBOT-SELF--">
	<p>Código: <input type="text" name="Codigo" size="30"><br>
	Descripción: <input type="text" name="Descripcion" size="37"><br>
	FechaPrecio: <input type="text" name="FechaPrecio" size="20"><br>
	Precio de Costo: <input type="text" name="PrecioCosto" size="20"><br>
	Precio Sugerido: <input type="text" name="PrecioSugerido" size="20"><br>
	Stock: <input type="text" name="Stock" size="20"><br>
	Venta: <select size="1" name="Venta">
    <option selected value="1">Venta Si</option>
    <option value="0">Venta No</option>
    </select><br>

	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Items",ObConn, 3, 3

ObRs.AddNew
ObRs ("Codigo") = Request.Form ("Codigo")
ObRs ("Descripcion") = Request.Form ("Descripcion")
ObRs ("FechaPrecio") = Request.Form ("FechaPrecio")
ObRs ("PrecioCosto") = Request.Form ("PrecioCosto")
ObRs ("PrecioSugerido") = Request.Form ("PrecioSugerido")
ObRs ("Stock") = Request.Form ("Stock")
ObRs ("Venta") = Request.Form ("Venta")
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