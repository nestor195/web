<%@ Language=VBScript %>
<% Response.Buffer = TRUE %>

<%
  ' Script Name   : aspWebCalendar FREE
  ' File Name     : color_picker.asp
  ' Version       : 1.0
  ' Release Date  : 4/10/2006
  '
  ' Copyright (c) 2002-2006 by Full Revolution, Inc., All Rights Reserved

%>
<%

'******** Checking some things for the calendar pop ups **************************

If Request.Querystring("Form") <> "" Then 
   FormName = Request.Querystring("Form") 
   Session("FormName") = FormName 
Else 
   FormName = Session("FormName") 
End If 
If Request.Querystring("Element") <> "" Then 
   ElementName = Request.Querystring("Element") 
   Session("ElementName") = ElementName 
Else 
   ElementName = Session("ElementName") 
End If

%>

<SCRIPT LANGUAGE="javascript"> 
function calpopulate(dte) { 
window.opener.<%=formname & "." & elementname%>.value = dte; 
  self.close() 
      } 

function changecolor()
{

var ca = new Array('0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F');

 colorvalue = "#";
 r = parseInt(document.ColorForm.RedValue.value);
 h = r%16; 
 k = (r - h)/16;
 colorvalue = colorvalue + ca[k];
 colorvalue = colorvalue + ca[h];
 
 g = parseInt(document.ColorForm.GreenValue.value);
 h = g%16; 
 k = (g - h)/16;
 colorvalue = colorvalue + ca[k];
 colorvalue = colorvalue + ca[h];
 
 b = parseInt(document.ColorForm.BlueValue.value);
 h = b%16; 
 k = (b - h)/16;
 colorvalue = colorvalue + ca[k];
 colorvalue = colorvalue + ca[h];
 
 document.ColorForm.ColorSquare.style.backgroundColor = colorvalue;
 document.ColorForm.HexValue.value = colorvalue;
 
}
</SCRIPT>

<HTML>
<BODY BGCOLOR="#E0DFE3">

<font face="Verdana" size="1"><span lang="en-us">Standard Colors:</span></font><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="200" id="AutoNumber1">
  <tr>
    <td width="100%">
    <table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
      <tr>
        <td width="12%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber3" bgcolor="#000000">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#000000')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="12%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber4" bgcolor="#FFFFFF">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#FFFFFF')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="12%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber5" bgcolor="#008000">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#008000')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="12%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber6" bgcolor="#800000">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#800000')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="13%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber7" bgcolor="#808000">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#808000')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="13%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber8" bgcolor="#000080">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#000080')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="13%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber9" bgcolor="#800080">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#800080')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="13%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber10" bgcolor="#808080">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#808080')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
      </tr>
      <tr>
        <td width="12%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber11" bgcolor="#FFFF00">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#FFFF00')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="12%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber12" bgcolor="#00FF00">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#00FF00')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="12%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber13" bgcolor="#00FFFF">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#00FFFF')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="12%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber14" bgcolor="#FF00FF">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#FF00FF')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="13%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber15" bgcolor="#C0C0C0">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#C0C0C0')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="13%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber16" bgcolor="#FF0000">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#FF0000')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="13%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber17" bgcolor="#0000FF">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#0000FF')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
        <td width="13%">
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber18" bgcolor="#008080">
          <tr>
            <td width="100%" style="cursor: hand" onClick="calpopulate('#008080')"><span lang="en-us"><font face="Verdana" size="1">&nbsp;</font></span></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<form name="ColorForm" method="POST" action="color_picker.asp" webbot-action="--WEBBOT-SELF--">
<TABLE ALIGN=CENTER WIDTH="100%">
  <TR>
    <TD WIDTH="100" ALIGN=LEFT VALIGN=TOP>
      <font face="Verdana" size="1"><B>Custom Color:</B><br>
  R: <input type="text" onchange="changecolor()" name="RedValue" size="3" VALUE="0"><BR>
  G: <input type="text" onchange="changecolor()" name="GreenValue" size="3" VALUE="0"><BR>
  B: <input type="text" onchange="changecolor()" name="BlueValue" size="3" VALUE="0"><BR></FONT>
    </TD>
    <TD WIDTH="100" ALIGN=LEFT valign=TOP>
      <font face="Verdana" size="1">
      HEX:<BR><INPUT TYPE="TEXT" SIZE="7" NAME="HexValue" readonly><BR><BR>
      Preview:<BR><input type="text" name="ColorSquare" size="9" style="border-style: solid; border-width: 1; background-color: #E0DFE3"><br>
    </FONT></TD>
   </TR>
</TABLE>
<CENTER><input type="button" value="Use Custom Color"  onClick="calpopulate(HexValue.value)" name="B1"></CENTER></p>
</form>
</BODY>
</HTML>