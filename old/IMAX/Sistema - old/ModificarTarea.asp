<html>

<head>
<meta http-equiv="Content-Language" content="es">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<%
IF Request.Form = "" THEN
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Tareas Where IdTarea = " & Request.QueryString("IdTarea")
ObRs.Open SQL, ObConn
%>
<b>Ingreso de<span lang="es"> Tarea</span></b><form method="POST" action="ModificarTarea.asp" webbot-action="--WEBBOT-SELF--">
	<p><span lang="es">IdTarea<select size="1" name="IdTarea">
    <option value="<%Response.Write ObRs ("IdTarea")%>"><%Response.Write ObRs ("IdTarea")%></option>
    </select></span></p>
	<p><span lang="es">Tarea</span>: 
    <input type="text" name="Tarea" size="96" value="<%Response.Write ObRs ("Tarea")%>"></p>
	<p><span lang="es">Completado %:
    <input type="text" name="Completado" size="6" value="<%Response.Write ObRs ("Completado")%>"></span><br>
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
SQL = "Select * From Tareas Where IdTarea = " & Request.Form ("IdTarea")
ObRs.Open SQL,ObConn, 3, 3

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