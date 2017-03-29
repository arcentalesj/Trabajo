<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%> 
<!--#include file="Connections/usuarios.asp" -->
<%
titulo = "CAPTACION Y SEGUIMIENTO DE CLIENTES"
%>
<html>
<head>
	<title>Captacion de Clientes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
	<table width="55%" height="130" align="center">
		<tr>
			<td width="20%"><img name="texto" src="imagenes/flecha-derecha.png" width="249" height="153" border="0" alt=""></td>
			<td width="30%" align="center">
				<FORM NAME="volver" ACTION="principal.asp">
					<input type="submit" name="" value="Regresar">
				</FORM>
			</td>
			<td width="20%"><img name="texto" src="imagenes/flecha-izquierda.png" width="249" height="153" border="0" alt=""></td>
		</tr>
	</table>
	<p>&nbsp;</p>
	<table width="42%" height="130" align="center">
		<tr>
			<td class="leyenda">
				Si esta viendo esta pantalla, por favor salga de la aplicación presionando el Bot&oacute;n regresar. </td>
		</tr>
		<tr>
			<td class="leyenda">
				El tiempo que le da el Internet Explorer a caducado.</td>
		</tr>
		<tr>
			<td class="leyenda">
				Gracias</td>
		</tr>
  </table>
</tr>
</body>
</html>
