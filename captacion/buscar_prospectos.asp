<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<%
dim cedula1
session("cedula") = request.querystring("cedula1")
session("docum")= request.querystring("docum")
session("bien")= request.querystring("bien")
session("persona")= request.querystring("perso")
'
' primeramente verifica si el cliente no tiene resgistrada la cedula como prohibición
'
strSQL = "select * from cedula_problemas where ced_cedula='"&session("cedula")&"'"
Set rprohibido = Server.CreateObject("ADODB.Recordset")
rprohibido.ActiveConnection = MM_usuarios_STRING
rprohibido.Source = strSQL
rprohibido.CursorType = 0
rprohibido.CursorLocation = 2
rprohibido.LockType = 1
rprohibido.Open()
if rprohibido.bof and rprohibido.eof then
	' Prospecto no tiene problemas
	strSQL = "select * from reasigna_asesor where cedula='"&session("cedula")&"'"
	Set novendedor = Server.CreateObject("ADODB.Recordset")
	novendedor.ActiveConnection = MM_usuarios_STRING
	novendedor.Source = strSQL
	novendedor.CursorType = 0
	novendedor.CursorLocation = 2
	novendedor.LockType = 1
	novendedor.Open()
	if novendedor.bof and novendedor.eof then
		' el Prospecto no esta ingresado
		strSQL = "select * from captados where cedula = '"+session("cedula")+"'"
		Set captados = Server.CreateObject("ADODB.Recordset")
		captados.ActiveConnection = MM_usuarios_STRING
		captados.Source = strSQL
		captados.CursorType = 0
		captados.CursorLocation = 2
		captados.LockType = 1
		captados.Open()
		if captados.bof and captados.eof then
			' Prospecto no existe
			response.redirect("ingreso_prospectos.asp?codven11=1")
		else
			' prospecto ya existe
			if trim(session("usuario"))<> trim(captados("idvendedor")) then
					response.redirect("cedula_prospectos.asp?mensajeerror=" &"cliente_nosuyo")
			end if
			response.redirect("modifica_prospectos.asp?cedula="&session("cedula")&"&cotiza="&(cotiza)&"&tipobien="&(session("bien")))
		end if
	else
		' el Prospecto esta ingresado pero vendedor ha salido de la empresa
		response.redirect("cedula_prospectos.asp?mensajeerror=" &"vende_inact")
	end if
else
	' Prospecto no tiene problemas no puede ingresar
	response.redirect("cedula_prospectos.asp?mensajeerror=" &"Cliente_Prohibido")
end if
%>
<html>
	<head>
		<title>Prospectos de Clientes</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	</head>
	<body text="#000000">
	</body>
</html>
