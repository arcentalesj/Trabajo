<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/usuarios.asp" -->
<%
session("cedula") = request.querystring("cedula1")
session("usuario") = request.querystring("codigo1")
Dim dbasesor
Dim dbasesor_numRows
Set dbasesor = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from asesores where cedula = '"+session("cedula")+"' and codven = '"+rtrim(ltrim(session("usuario")))+"'"
dbasesor.ActiveConnection = MM_usuarios_STRING
dbasesor.CursorType = 0
dbasesor.CursorLocation = 2
dbasesor.LockType = 1
dbasesor.Source = strSQL
dbasesor.Open()
if dbasesor.bof and dbasesor.eof then
	response.redirect("principal.asp?mensajeerror="&"cedula_nohay")		
else
	session("quien")=dbasesor.Fields.Item("alias").value
	session("nombrease")=dbasesor.Fields.Item("nombre").value
	session("tipocliente")=dbasesor.Fields.Item("nomgru").value
	session("cargo")=dbasesor.Fields.Item("cargo").value
	response.redirect("menu.asp?usuario="&session("usuario")&"&nombre="&session("nombrease"))		
end if
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
		<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
	</head>
	<body class="cuerpo">
		<table width="100%" class="gradient" >
			<tr> 
				<td width="55%" rowspan="7" class="Estilo4" ><div align="Center"><%response.write(session("Sistema"))%></div></td>
				<td width="15%" height="50%">&nbsp;</td>
				<td width="20%" rowspan="5" class="gradient"><img name="texto" src="imagenes/logo.gif" width="249" height="153" border="0" alt=""></td>
			</tr>
			<tr>
				<td height="50%"></td>
			</tr>
			<tr>
				<td class="fecha"><script type="text/javascript">fecha();</script></td>
			</tr>
			<tr>
				<td height="100%">&nbsp;</td>
			</tr>
		</table>
		<table cellspacing="0" class="Estilo2">
			<tr></tr>
		</table>
		<p>&nbsp;</p>
	</body>
</html>
