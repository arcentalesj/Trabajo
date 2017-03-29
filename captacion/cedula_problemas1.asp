<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<% session("estado")= 4 %>
<%
'dim cedula1
titulo="CLIENTES NO DESEADOS"
session("cedula") = rtrim(request.querystring("cedula1"))
' Verifica si el cliente existe y no se encuentre descartado
strSQL = "select * from cedula_problemas where ced_cedula='"&session("cedula")&"'"
Set rtodoburo = CreateObject("ADODB.Recordset")
rtodoburo.Open strSQL, MM_usuarios_STRING
paraactivar=0
'response.write (strsql)
if rtodoburo.bof and rtodoburo.eof then
	strSQL = "select * from captados where cedula='"&session("cedula")&"'"
	Set rtodoburo = CreateObject("ADODB.Recordset")
	rtodoburo.Open strSQL, MM_usuarios_STRING
	if rtodoburo.bof and rtodoburo.eof then
		' El cliente no está registrado en las bases de datos, se genera un registro nuevo
		session("siexiste") = 0
		fechacap = date()
	else
		nombre   = rtodoburo.Fields.Item("nombre").Value
		apellido = rtodoburo.Fields.Item("apellido").Value
		fechacap = rtodoburo.Fields.Item("ingresado").Value
		session("siexiste") = 0
	end if
else
	' existe cliente, verifica el cotizador mas nuevo ingresado
	nombre   = rtodoburo.Fields.Item("ced_nombre").Value
	apellido = rtodoburo.Fields.Item("ced_apellido").Value
	motivo   = rtodoburo.Fields.Item("ced_motivo").Value
	fechacap = rtodoburo.Fields.Item("ced_fecregistro").Value
	session("estatus")  = rtodoburo.Fields.Item("ced_estatus").Value
	session("siexiste") = 1
end if
%>
<html>
	<head>
		<title>Ingreso de Documentos no deseados</title>
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
	<table width="100%" height ="2%"><tr></tr></table>
	<table align="center">
		<tr> 
			<td class="titulo"><%response.write(titulo)%></td>
		</tr>
		<tr><td height="10"></td></tr>
		<tr> 
			<td class="subtitulo"><%response.write(session("titulo"))%></td>
		</tr>
		<tr> 
			<td class="subtitulo"><%response.write(session("titulo1"))%></td>
		</tr>
	</table>
	<table width="100%" height ="2%"><tr></tr></table>
	<form method="post" action="grabadatos.asp" name="form4" AUTOCOMPLETE="OFF">
		<table width="837" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ACE5EE">
			<tr> 
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Cedula</td>
				<td width="165" class="presenta2"><%response.write(session("cedula"))%></td>
				<td width="24">&nbsp;</td>
				<td width="12">&nbsp;</td>
				<td width="131" class="presenta1">Fecha de Registro</td>
				<td width="163"> <font size="2"><%response.write(fechacap)%></font></td>
			</tr>
			<tr> 
				<td width="12" class="presenta1">*</td>
				<td width="131" class="presenta1">Apellidos</td>
				<td width="165" class="presenta2">
					<input name="apellido" type="text" value="<%response.write(apellido)%>"	maxlength="30" tabindex ="1" required >
				</td>
					<td width="24">&nbsp;</td>
				<td width="12" class="presenta1">*</td>
					<td width="131" class="presenta1">Nombres</td>
					<td width="163" class="presenta2">
                    	<input name="nombre" type="text" value="<%response.write(nombre)%>" size="25" maxlength="30" tabindex ="2" required>
					</td>
		  </tr>
			</tr>
		</table>
		<p>&nbsp;</p>
		<table width="728"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ACE5EE">
			<tr> 
			  <td width="17" class="presenta1">*</td>
				<td width="107" class="presenta1">Motivo : </td>
				<td width="540" class="presenta2">
					<input name="motivo" type="text" value="<%response.write(motivo)%>" size="90" maxlength="90" tabindex="3" required>
			  </td>
			</tr>
		</table>
		<table width="728"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#ACE5EE">
        	<%if session("estatus")=1 then
        		estatus="CLIENTE SI PUEDE INGRESAR AL SISTEMA"
	        else
    	    	estatus="CLIENTE NO PUEDE INGRESAR AL SISTEMA"
        	end if%>
			<tr> 
				<td width="17"><font color="#990000" size="2"><div align="right"><b>Estatus:&nbsp;&nbsp;&nbsp;</b></div></font></td>
				<td width="540" align="left" bordercolor="#000000" bgcolor="#FFFFFF"> <font size="3">&nbsp;&nbsp;<%response.write(estatus)%></font></td>
			</tr>
		</table>
		<table width="100%">
			<tr> 
			  <td>&nbsp;</td>
			</tr>
		</table>
        <%if session("siexiste")=1 then%>
			<p>&nbsp;</p>
			<table align="center">
				<tr> 
					<td width="159" align="center"><font size="1">
						<select name="paraactivar" id="paraactivar" tabindex="60">
							<option value=0>Activar Cliente</option>
							<option value=1>SI</option>
							<option value=2>NO</option>
						</select>
						</font>
                    </td>
				</tr>
			</table>
			<p>&nbsp;</p>
        <%end if%>
		<table width="100%" height ="6%"><tr></tr></table>
		<table align="center">
			<tr>
				<td width="33%" align="center">&nbsp;</td>
				<td width="33%" align="center">&nbsp;</td>
				<td colspan="2" align="center"><input type="submit"  name="btnBuscar" value="grabar"></button>
			</tr>
		</table>
	</form>
	<table width="100%">
		<tr>
			<td width="26%" align="center">
				<FORM NAME="volver" ACTION="cedula_problemas.asp">
					<input type="submit" name="" value="Regresar">
				</FORM>
			</td>
			<td width="26%">&nbsp;</td>
		</tr>
	</table>
	</body>
</html>
<%
rtodoburo.Close()
Set rtodoburo = Nothing
%>

