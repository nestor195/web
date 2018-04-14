<%
   ' -- file.asp --
   ' Retrieves binary files from the database
   
   Response.Buffer = True
   
   ' ID of the file to retrieve
   Dim ID
      ID = Request("ID")
      
   If Len(ID) < 1 Then
      ID = 7
   End If
   
   ' Connection String
   Dim connStr
      connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & _
         Server.MapPath("FileDB.mdb")
   
   ' Recordset Object
   Dim rs
      Set rs = Server.CreateObject("ADODB.Recordset")
      
      ' opening connection
      rs.Open "select [File Data],[Content Type] from Files where ID = " & _
         ID, connStr, 2, 4

      If Not rs.EOF Then
         Response.ContentType = rs("Content Type")
         Response.BinaryWrite rs("File Data")
      End If
      
      
      rs.Close
      Set rs = Nothing
%>