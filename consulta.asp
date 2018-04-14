<%
If Session("loginokay") = "" then
Response.redirect "login.asp"
end if
%>

<html>
<head>

<meta content="es-ar" http-equiv="Content-Language">
<title>SOLICITUD DE ACCIONES CORRECTIVAS Y PREVENTIVAS</title>
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type">


<link href="estilo.css" rel="stylesheet" type="text/css">

</head>
<body>
<%
IF Request.QueryString ("nuevoingreso") = "nuevo" then
Id = "Nuevo"
Estado = ""
Fecha = null
Solicita = ""
Area = ""
Accion = ""
NoConformidad = ""
CausaNoConformidad = ""
DescripcionAccion = ""
FechaSolicitante = null
FirmaSolicitante = ""
FechaResponsable = null
FirmaResponsable = ""
FechaGestionCalidad = null
FirmaGestionCalidad = ""
FechaImplementacion = null
AccionARealizar = ""
FechaSectorResponsableV = null
FechaVerificacionEfectividad = null
ResponsableVerificacion = ""
AccionARealizarEfectiva = ""
FechaResponsableEfectivo = null
ResponsableEfectivo = ""
NoConformidadTipo = 0
CausaNoConformidadTipo = 0
Else

SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Planilla Where Id = " & Request.Querystring ("AC")
ObRs.Open Sel,ObConn
Id = ObRs ("Id")
Estado = ObRs ("Estado")
Fecha = ObRs ("Fecha")
Solicita = ObRs ("Solicita")
Area = ObRs ("Area")
Accion = ObRs ("Accion")
NoConformidad = ObRs ("NoConformidad")
CausaNoConformidad = ObRs ("CausaNoConformidad")
Contencion = ObRs ("Contencion")
DescripcionAccion = ObRs ("DescripcionAccion")
FechaSolicitante = ObRs ("FechaSolicitante")
FirmaSolicitante = ObRs ("FirmaSolicitante")
FechaResponsable = ObRs ("FechaResponsable")
FirmaResponsable = ObRs ("FirmaResponsable")
FechaGestionCalidad = ObRs ("FechaGestionCalidad")
FirmaGestionCalidad = ObRs ("FirmaGestionCalidad")
FechaImplementacion = ObRs ("FechaImplementacion")
AccionARealizar = ObRs ("AccionARealizar")
FechaSectorResponsableV = ObRs ("FechaSectorResponsableV")
FechaVerificacionEfectividad = ObRs ("FechaVerificacionEfectividad")
ResponsableVerificacion = ObRs ("ResponsableVerificacion")
AccionARealizarEfectiva = ObRs ("AccionARealizarEfectiva")
FechaResponsableEfectivo = ObRs ("FechaResponsableEfectivo")
ResponsableEfectivo = ObRs ("ResponsableEfectivo")
NoConformidadTipo = ObRs ("NoConformidadTipo")
CausaNoConformidadTipo = ObRs ("CausaNoConformidadTipo")
ObRs.Close
ObConn.Close

End If
%>

<p class="Tilulo">SOLICITUD DE ACCIONES CORRECTIVAS Y PREVENTIVAS<br></p>
<form method="post" action="guardaraacc.asp">
<table cellpadding="2" cellspacing="0" class="tablas" style="width: 100%">
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<table style="width: 100%">
			<tr>
				<td class="celdaazul" style="width: 112px">Accion N� </td>
				<td class="auto-style7">
				<input name="AC" type="text" readonly="readonly" value="<%Response.Write Id%>"></td>
				<td class="celdaazul">Estado</td>
				<td>
				<select name="Estado">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Estados order by Estado asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If Estado = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Estado")%></option>
<%
Selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Estado")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If
ObRs.Close
ObConn.Close
%>
				</select>
				</td>
			</tr>
		</table>
		</td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<table cellpadding="2" cellspacing="0" class="tablas" style="width: 100%">
			<tr>
				<td class="celdaazul" style="width: 82px; height: 25px">Fecha</td>
				<td class="celdaazul" style="width: 301px; height: 25px">
				Solicitada por:</td>
				<td class="celdaazul" style="width: 131px; height: 25px">Area:</td>
				<td class="celdaazul" style="height: 25px">Tipo de acci�n a
				Implementar</td>
			</tr>
			<tr>
				<td class="auto-style2" style="width: 82px">
<%
If Month(Fecha) < 10 Then
guion1 = "-0"
Else
guion1 = "-"
End IF
If Day(Fecha) < 10 Then
guion2 = "-0"
Else
guion2 = "-"
End IF

ValorFecha = Year(Fecha) & guion1 & Month(Fecha) & guion2 & Day(Fecha)
%>
				<input type="date" name="fecha1" value="<%Response.Write ValorFecha%>"></td>
				<td class="auto-style2" style="width: 301px">
				<select name="Solicita">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Responsables order by Responsable asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If Solicita = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
Selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If
ObRs.Close
ObConn.Close
%>
				</select></td>
				<td class="auto-style2" style="width: 131px">
				<select name="Area">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Areas order by Area asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If Area = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Area")%></option>
<%
Selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Area")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If
ObRs.Close
ObConn.Close
%>
				</select></td>
				<td class="auto-style2">
				<select name="Accion" style="height: 22px">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Acciones order by Accion asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If Accion = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Accion")%></option>
<%
Selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Accion")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If
ObRs.Close
ObConn.Close
%>
				</select></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
<%
Select Case NoConformidadTipo
Case 0
NCT0 = "checked='checked'"
NCT1 = ""
NCT2 = ""
NCT3 = ""
Case 1
NCT0 = ""
NCT1 = "checked='checked'"
NCT2 = ""
NCT3 = ""
Case 2
NCT0 = ""
NCT1 = ""
NCT2 = "checked='checked'"
NCT3 = ""
Case 3
NCT0 = ""
NCT1 = ""
NCT2 = ""
NCT3 = "checked='checked'"
End Select
%>
		<td class="celdaazul" colspan="2" style="height: 24">NO CONFIRMIDAD
		EXISTENTE<input <%Response.Write NCT0%> name="NoConformidadTipo" type="radio" value="0">&nbsp;&nbsp;&nbsp; - POTENCIAL<input <%Response.Write NCT1%> name="NoConformidadTipo" type="radio" value="1">&nbsp;&nbsp;&nbsp; - OBSERVACI�N<input <%Response.Write NCT2%> name="NoConformidadTipo" type="radio" value="2">&nbsp;&nbsp;&nbsp; - OPORTUNIDAD DE MEJORA<input <%Response.Write NCT3%> name="NoConformidadTipo" type="radio" value="3"></td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<textarea name="NoConformidad" style="width: 500px; height: 60px"><%Response.Write NoConformidad%></textarea></td>
	</tr>
	<tr>
<%
Select Case CausaNoConformidadTipo
Case 0
CNCT0 = "checked='checked'"
CNCT1 = ""
CNCT2 = ""
Case 1
CNCT0 = ""
CNCT1 = "checked='checked'"
CNCT2 = ""
Case 2
CNCT0 = ""
CNCT1 = ""
CNCT2 = "checked='checked'"
End Select
%>
		<td class="celdaazul" colspan="2">CAUSA DE LA NO CONFORMIDAD EXISTENTE<input <%Response.Write CNCT0%> name="CausaNoConformidadTipo" type="radio" value="0">&nbsp;&nbsp;&nbsp;
		- POTENCIAL<input <%Response.Write CNCT1%> name="CausaNoConformidadTipo" type="radio" value="1">&nbsp;&nbsp;&nbsp; - OBSERVACI�N<input <%Response.Write CNCT2%> name="CausaNoConformidadTipo" type="radio" value="2"></td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<textarea name="CausaNoConformidad" style="width: 500px; height: 60px"><%Response.Write CausaNoConformidad%></textarea></td>
	</tr>
	<tr>
		<td class="celdaazul" colspan="2">DESCRIPCI�N DE LA ACCI�N CORRECTIVA
		O PREVENTIVA A SER IMPLEMENTADA</td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<textarea name="DescripcionAccion" style="width: 500px; height: 60px"><%Response.Write DescripcionAccion%></textarea></td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<table cellpadding="2" cellspacing="0" style="width: 100%" class="tablas">
			<tr>
				<td class="celdaazul" rowspan="4" style="width: 155px">
				Conforme del sector solicitante</td>
				<td class="celdaazul" style="height: 24px; width: 111px">
				Responsable:</td>
				<td class="celdaazul" rowspan="4" style="width: 172px">
				Conforme del sector responsable de implementaci�n</td>
				<td class="celdaazul" style="height: 24px; width: 108px">
				Responsable:</td>
				<td class="celdaazul" rowspan="4" style="width: 177px">
				Conforme Gesti�n de Calidad</td>
				<td class="celdaazul" style="height: 24px">Responsable:</td>
			</tr>
			<tr>
				<td class="auto-style2" style="height: 24px; width: 111px">
				<select name="FirmaSolicitante">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Responsables order by Responsable asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If FirmaSolicitante = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
Selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If
ObRs.Close
ObConn.Close
%>
				</select></td>
				<td class="auto-style2" style="height: 24px; width: 108px">
				<select name="FirmaResponsable">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Responsables order by Responsable asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If FirmaResponsable = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
Selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If
ObRs.Close
ObConn.Close
%>
				</select></td>
				<td class="auto-style2" style="height: 24px">
				<select name="FirmaGestionCalidad">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Responsables order by Responsable asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If FirmaGestionCalidad = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
Selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If

ObRs.Close
ObConn.Close
%>
				</select></td>
			</tr>
			<tr>
				<td class="celdaazul" style="width: 111px">Fecha:</td>
				<td class="celdaazul" style="width: 108px">Fecha:</td>
				<td class="celdaazul">Fecha:</td>
			</tr>
			<tr>
				<td class="auto-style2" style="width: 111px">
<%
If Month(FechaSolicitante) < 10 Then
guion1 = "-0"
Else
guion1 = "-"
End IF
If Day(FechaSolicitante) < 10 Then
guion2 = "-0"
Else
guion2 = "-"
End IF

ValorFecha = Year(FechaSolicitante) & guion1 & Month(FechaSolicitante) & guion2 & Day(FechaSolicitante)
%>
				<input name="FechaSolicitante" type="date" value="<%Response.Write ValorFecha%>"></td>
				<td class="auto-style2" style="width: 108px">
<%
If Month(FechaResponsable) < 10 Then
guion1 = "-0"
Else
guion1 = "-"
End IF
If Day(FechaResponsable) < 10 Then
guion2 = "-0"
Else
guion2 = "-"
End IF

ValorFecha = Year(FechaResponsable) & guion1 & Month(FechaResponsable) & guion2 & Day(FechaResponsable)
%>
				<input name="FechaResponsable" type="date" value="<%Response.Write ValorFecha%>"></td>
				<td class="auto-style2">
<%
If Month(FechaGestionCalidad) < 10 Then
guion1 = "-0"
Else
guion1 = "-"
End IF
If Day(FechaGestionCalidad) < 10 Then
guion2 = "-0"
Else
guion2 = "-"
End IF

ValorFecha = Year(FechaGestionCalidad) & guion1 & Month(FechaGestionCalidad) & guion2 & Day(FechaGestionCalidad)
%>
				<input name="FechaGestionCalidad" type="date" value="<%Response.Write ValorFecha%>"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td class="celdaazul" style="width: 103px">Fecha de implementaci�n:</td>
		<td class="auto-style2" style="width: 198px">
<%
If Month(FechaImplementacion) < 10 Then
guion1 = "-0"
Else
guion1 = "-"
End IF
If Day(FechaImplementacion) < 10 Then
guion2 = "-0"
Else
guion2 = "-"
End IF

ValorFecha = Year(FechaImplementacion) & guion1 & Month(FechaImplementacion) & guion2 & Day(FechaImplementacion)
%>
		<input name="FechaImplementacion" type="date" value="<%Response.Write ValorFecha%>"></td>
	</tr>
	<tr>
		<td class="celdaazul" colspan="2">ACCI�N CORRECTIVA O PREVENTIVA
		IMPLEMENTADA (Descripci�n de evidencia de implementaci�n)</td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<textarea name="AccionARealizar" style="width: 496px; height: 56px"><%Response.Write AccionARealizar%></textarea></td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2" style="height: 24px">

		<table cellpadding="2" cellspacing="0" style="width: 100%" class="tablas">
			<tr>
				<td class="celdaazul" rowspan="2" style="width: 137px">
				Conforme del sector responsable de verificaci�n</td>
				<td class="celdaazul" style="width: 248px">Fecha</td>
				<td class="auto-style2" style="width: 84px">
<%
If Month(FechaSectorResponsableV) < 10 Then
guion1 = "-0"
Else
guion1 = "-"
End IF
If Day(FechaSectorResponsableV) < 10 Then
guion2 = "-0"
Else
guion2 = "-"
End IF

ValorFecha = Year(FechaSectorResponsableV) & guion1 & Month(FechaSectorResponsableV) & guion2 & Day(FechaSectorResponsableV)
%>
				<input name="FechaSectorResponsableV" type="date" value="<%Response.Write ValorFecha%>"></td>
				<td class="celdaazul" rowspan="2" style="width: 117px">
				Responsable:</td>
				<td class="auto-style2" rowspan="2">
				<select name="ResponsableVerificacion" style="height: 22px">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Responsables order by Responsable asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If ResponsableVerificacion = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If
ObRs.Close
ObConn.Close
%>
				</select></td>
			</tr>
			<tr>
				<td class="celdaazul" style="width: 248px">Fecha de
				verificaci�n de efectividad</td>
				<td class="auto-style2" style="width: 84px">
<%
If Month(FechaVerificacionEfectividad) < 10 Then
guion1 = "-0"
Else
guion1 = "-"
End IF
If Day(FechaVerificacionEfectividad) < 10 Then
guion2 = "-0"
Else
guion2 = "-"
End IF
ValorFecha = Year(FechaVerificacionEfectividad) & guion1 & Month(FechaVerificacionEfectividad) & guion2 & Day(FechaVerificacionEfectividad)
%>
				<input name="FechaVerificacionEfectividad" type="date" value="<%Response.Write ValorFecha%>"></td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td class="celdaazul" colspan="2">ACCI�N CORRECTIVA O PREVENTIVA
		EFECTIVA (Descripci�n de evidencia efectiva)</td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<textarea name="AccionARealizarEfectiva" style="width: 496px; height: 55px"><%Response.Write AccionARealizarEfectiva%></textarea></td>
	</tr>
	<tr class="auto-style2">
		<td class="auto-style2" colspan="2">
		<table class="auto-style6" style="width: 100%"class="tablas">
			<tr>
				<td class="celdaazul" style="width: 137px">Conforme del sector
				responsable de verificaci�n</td>
				<td class="celdaazul" style="width: 243px">Fecha</td>
				<td class="auto-style2" style="width: 88px">
<%
If Month(FechaResponsableEfectivo) < 10 Then
guion1 = "-0"
Else
guion1 = "-"
End IF
If Day(FechaResponsableEfectivo) < 10 Then
guion2 = "-0"
Else
guion2 = "-"
End IF
ValorFecha = Year(FechaResponsableEfectivo) & guion1 & Month(FechaResponsableEfectivo) & guion2 & Day(FechaResponsableEfectivo)
%>
				<input name="FechaResponsableEfectivo" type="date" value="<%Response.Write ValorFecha%>"></td>
				<td class="celdaazul" style="width: 110px">Responsable:</td>
				<td class="auto-style2">
				<select name="ResponsableEfectivo">
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
%>
<!--#include file="conector.asp"-->
<%
Sel = "SELECT * FROM Responsables order by Responsable asc"
ObRs.Open Sel,ObConn
Selected = 0
Do While ObRs.EOF = false
If ResponsableEfectivo = ObRs ("Id") then
%>
				<option selected value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
selected = 1
Else
%>
				<option value="<%Response.Write ObRs ("Id")%>"><%Response.Write ObRs ("Responsable")%></option>
<%
End If
ObRs.MoveNext
Loop
If Selected = 0 then
%>
				<option selected value="">Ingrese Opcion</option>
<%
End If
ObRs.Close
ObConn.Close
%>
				</select></td>
			</tr>
		</table>
		</td>
	</tr>
</table>
<input name="Submit1" type="submit" value="Enviar">
</form>

</body>
</html>
