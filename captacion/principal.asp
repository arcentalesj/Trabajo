<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Dim mensajeerror
If (Request.QueryString("mensajeerror") <> "") Then 
	session("dd") = Request.QueryString("mensajeerror")
End If
session("Sistema") = "CAPTACION Y SEGUIMIENTO DE CLIENTES"
session("color") ="#E4E4E4"
titulo1 = "Datos necesarios para el Ingreso al Sistema"
session("presenta")=0
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<title>Ingreso al sistema</title>
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
	<table width="791" align="center">
		<tr> 
			<td class="titulo"><%response.write(titulo1)%></td>
		</tr>
	</table>
	<p>&nbsp;</p>
	<table align="center">
		<tr>
			<td>
				<form name="formulario" action="valdatos()" method="post" >
					<table width="353" align="center" cellspacing="0" class="formulario">
						<tr>
							<td class="Estilo3">CEDULA DE EMPLEADO</font></td>
							<td width="178"><input type="text" pattern="[0-9]+{10}" id="cedula" name="cedula" onBlur="valcedula()" required autofocus></td>
						</tr>
						<tr>
							<td class="Estilo3">CODIGO DE EMPLEADO</td>
							<td><input type="text" pattern="[0-9]+{1,2}" id="codigo"  name="codigo"  onBlur="valcodigo()"required></td>
						</tr>
						<tr>
							<td class="Estilo3" height="50">&nbsp;</td>
							<td></td>
						</tr>
						<tr>
							<td colspan="2" align="center">
								<input type="button" class="boton" name="btnBuscar" value="Ingresar" onclick='javascript:valdatos();'/>
							</td>
						</tr>
					</table>
				</form>
			</td>
		</tr>
	</table>
	<p>&nbsp;</p>
	<table width="30%" height="109" border="0" align="center" cellpadding="0" cellspacing="0" style="BORDER-COLLAPSE: collapse">
		<tr>
			<td height="20" class="titulo">INFORMACION IMPORTANTE</p>
			</td>
		</tr>
		<tr>
			<td height="27" class="leyenda">
				No debe permacer en el sistema CSC por el lapso de 10 minutos sin realizar alguna actividad, caso contrario el sistemas le desactivará
			</td>
		</tr>
	</table>
	<p>&nbsp;</p>
	<table width="30%" height="109" border="0" align="center" style="border-collapse: collapse">
		<tr>
			<td height="20" ><div align="center"><script type="text/javascript">autor(); </script></div></td>
		</tr>
	</table>
	<p>&nbsp;</p>
	<table>
		<tr>
			<td width="26%" align="center"> 
					<%
					select case request.querystring("mensajeerror")
						case "cedula_nohay"
							response.write ("<center><b>La información ingresada no coincide con ningún empleado</b></center>")	
					end select
					%>
			</td>
		</tr>
	</table>
</body>
</html>

