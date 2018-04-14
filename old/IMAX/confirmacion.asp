<%
IF Session("Session") = "" THEN
Response.Redirect ("inicio.asp")
End If
%>

<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From Ordenes Where Id = " & Session ("Orden")
ObRs.Open SQL, ObConn, 3, 3


SET ObConn3 = Server.CreateObject ("ADODB.Connection")
SET ObRs3 = Server.CreateObject ("ADODB.RecordSet")
ObConn3.Open "Sistema"
SQL = "Select * From Usuarios Where Cliente = " & ObRs ("Cliente")
ObRs3.Open SQL,ObConn
Usuario = ObRs3 ("Id")
ObRs3.Close
ObConn3.Close


ObRs ("Estado") = 3
ObRs ("UsuarioEstado") = Usuario
ObRs ("FechaEstado") = DATE
ObRs.Update

ObRs.Close
ObConn.Close

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
ObRs.Open "Tareas",ObConn, 3, 3

ObRs.AddNew
ObRs ("Tarea") = "<font color='#FF0000'><b>Se confirma la Orden <a href='ConsultaDeOrden.asp?Id=" & Session ("Orden") & "'>" & Session ("Orden") & "</a> via Web "& DATE &" "&FormatDateTime(Now, 4)&"</b></font>"
ObRs ("Completado") = 0
ObRs.Update

ObRs.Close
ObConn.Close



Response.Redirect ("orden.asp?orden=" & Session ("Orden"))
%>
