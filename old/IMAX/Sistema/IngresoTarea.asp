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
<b>Ingreso de<span lang="es"> Tarea</span></b><form method="POST" action="IngresoTarea.asp" webbot-action="--WEBBOT-SELF--">
	<p><span lang="es">Tarea</span>: <input type="text" name="Tarea" size="94"></p>
	<p><span lang="es">Completado %:
    <input type="text" name="Completado" size="6" value="0"></span><br>
	<input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2">
	</p>
</form>
<%
ELSE
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Tareas",ObConn, 3, 3

ObRs.AddNew
ObRs ("Tarea") = Request.Form ("Tarea")
ObRs ("Completado") = Request.Form ("Completado")
ObRs.Update

ObRs.Close
ObConn.Close
%>
<b>Datos Ingresados</b>
<%
Response.Redirect ("default.asp")
END IF
%>

</body>
</html>