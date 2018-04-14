<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="es-ar" http-equiv="Content-Language" />
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type" />
<title>Ok</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
</head>
<%
SET ObConn = Server.CreateObject ("ADODB.Connection")
SET ObRs = Server.CreateObject ("ADODB.RecordSet")
ObConn.Open "ACYP"

If Request.Form ("AC") = "Nuevo" then
Sel = "Select * From Planilla"
Else
Sel = "Select * From Planilla Where Id = " & Request.Form ("AC")
End If

ObRs.Open Sel, ObConn, 3, 3

If Request.Form ("AC") = "Nuevo" then
ObRs.Addnew
End If

If Request.Form ("Estado") <> "" Then
ObRs ("Estado") = Request.Form ("Estado")
End If
If Request.Form ("Fecha1") <> "" Then
ObRs ("Fecha") = Request.Form ("Fecha1")
End If
If Request.Form ("Solicita") <> "" Then
ObRs ("Solicita") = Request.Form ("Solicita")
End If
If Request.Form ("Area") <> "" Then
ObRs ("Area") = Request.Form ("Area")
End If
If Request.Form ("Accion") <> "" Then
ObRs ("Accion") = Request.Form ("Accion")
End If

ObRs ("NoConformidad") = Request.Form ("NoConformidad")
ObRs ("CausaNoConformidad") = Request.Form ("CausaNoConformidad")
ObRs ("Contencion") = Request.Form ("Contencion")
ObRs ("DescripcionAccion") = Request.Form ("DescripcionAccion")

If Request.Form ("FechaSolicitante") <> "" Then
ObRs ("FechaSolicitante") = Request.Form ("FechaSolicitante")
End If
If Request.Form ("FirmaSolicitante") <> "" Then
ObRs ("FirmaSolicitante") = Request.Form ("FirmaSolicitante")
End If
If Request.Form ("FechaResponsable") <> "" Then
ObRs ("FechaResponsable") = Request.Form ("FechaResponsable")
End If
If Request.Form ("FirmaResponsable") <> "" Then
ObRs ("FirmaResponsable") = Request.Form ("FirmaResponsable")
End If
If Request.Form ("FechaGestionCalidad") <> "" Then
ObRs ("FechaGestionCalidad") = Request.Form ("FechaGestionCalidad")
End If
If Request.Form ("FirmaGestionCalidad") <> "" Then
ObRs ("FirmaGestionCalidad") = Request.Form ("FirmaGestionCalidad")
End If
If Request.Form ("FechaImplementacion") <> "" Then
ObRs ("FechaImplementacion") = Request.Form ("FechaImplementacion")
End If
If Request.Form ("AccionARealizar") <> "" Then
ObRs ("AccionARealizar") = Request.Form ("AccionARealizar")
End If
If Request.Form ("FechaSectorResponsableV") <> "" Then
ObRs ("FechaSectorResponsableV") = Request.Form ("FechaSectorResponsableV")
End If
If Request.Form ("FechaVerificacionEfectividad") <> "" Then
ObRs ("FechaVerificacionEfectividad") = Request.Form ("FechaVerificacionEfectividad")
End If
If Request.Form ("ResponsableVerificacion") <> "" Then
ObRs ("ResponsableVerificacion") = Request.Form ("ResponsableVerificacion")
End If
If Request.Form ("AccionARealizarEfectiva") <> "" Then
ObRs ("AccionARealizarEfectiva") = Request.Form ("AccionARealizarEfectiva")
End If
If Request.Form ("FechaResponsableEfectivo") <> "" Then
ObRs ("FechaResponsableEfectivo") = Request.Form ("FechaResponsableEfectivo")
End If
If Request.Form ("ResponsableEfectivo") <> "" Then
ObRs ("ResponsableEfectivo") = Request.Form ("ResponsableEfectivo")
End If

ObRs ("NoConformidadTipo") = Request.Form ("NoConformidadTipo")
ObRs ("CausaNoConformidadTipo") = Request.Form ("CausaNoConformidadTipo")


ObRs.Update
ObRs.Close
ObConn.Close
%>

<body>

<p>Ok</p>

</body>
<%
Response.Redirect "listado.asp"
%>
</html>
