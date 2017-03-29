<%@LANGUAGE="VBSCRIPT"%>
<% session("estado")= 7 %>
<!--#include file="Connections/usuarios.asp" -->
<%
titulo ="CIERRE DE MES"
%>
<html>
	<head>
		<title>Cierre de Mes PSC</title>
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
	<table border="1" cellpadding="1" cellspacing="0" align="center" width="500"  id="intermitente" style="border:5px solid blue" height="18">
		<tr> 
			<td class="titulo">Usted va a realizar el cierre de Mes</td>
		</tr>
		<tr> 
			<td class="titulo">Recuerde las siguientes indicaciones</td>
		</tr>
	</table>
	<table width="100%" height ="2%"><tr></tr></table>
	<table width=90% align="center">
		<tr>
			<td>
				<table width="44%" height="162%" border="0" align="center" cellpadding=0 cellspacing=0 >
					<tr> 
						<td class="leyenda">Se Eliminarán todos los cliente ingresados que no tengan citas.</font></td>
					</tr>
					<tr> 
						<td class="leyenda">Se Eliminarán todos los cliente con citas de meses anteriores.</font></td>
					</tr>
					<tr> 
						<td class="leyenda">Asigne citas a clientes potenciales.</font></td>
					</tr>
					<tr> 
						<td class="leyenda">Imprima información de clientes captados.</font></td>
					</tr>
			  </table>
				<table width="85%" border = "0"  height="20">
					<td width="60%" class="presentac"  height="20">&nbsp;</td>
					<td width="40%" >&nbsp;</td>
				</table >
			  <tr> 
					<td colspan="7" align="center">
						<FORM NAME="volver" ACTION="cierre_mes_1.asp">
							<input type="submit" name="" value="Cerrar Mes">
						</FORM>
					</td>
				</tr>
			</td>
		</tr>
	</table>
	<table width="100%"  align="center">
		<tr>
			<td width="14%">				</td>
			<td width="86%">
				<FORM NAME="volver" ACTION="principal.asp">
					<input type="submit" name="" value="Regresar">
				</FORM>
			</td>
		</tr>
	</table>
	</center>
	</body>
</html>
