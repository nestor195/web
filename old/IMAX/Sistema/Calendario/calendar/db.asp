<%
'*********************************************************************************
  ' Script Name   : aspWebCalendar FREE - Database Connection Support File
  ' File Name     : db.asp
  ' Version       : 1.0
  ' Release Date  : 4/10/2006
  '
  ' Copyright (c) 2002 - 2005 by Full Revolution, Inc., All Rights Reserved
'*********************************************************************************

'*********************************************************************************
'******** Open the database  *****************************************************
'*********************************************************************************

  '------- Access Driver Connection String ----------------------------------------------
  'strConn = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("kb/kb.mdb")

  '------- Access OLE DB Driver Connection String ---------------------------------------
  strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("calendar/calendar.mdb")
  strConn = strConn
  Set dbc = Server.CreateObject("ADODB.Connection")
  dbc.open strConn

  '------- Access Connection USING DSN --------------------------------------------------
  'Set dbc = Server.CreateObject("ADODB.Connection")
  'dbc.open "DSN=aspWebCalendarFREE"

%>
