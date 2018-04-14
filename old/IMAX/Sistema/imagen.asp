<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "Sistema"
SQL = "Select * From " & Request.QueryString("Tabla") & " Where Id = " & Request.QueryString("Id")
ObRs.Open SQL, ObConn
Response.ContentType = ObRs("ContentType")
Response.BinaryWrite ObRs("Imagen")

ObRs.Close
ObConn.Close
%>