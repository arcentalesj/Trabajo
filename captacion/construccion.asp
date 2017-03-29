<%@ LANGUAGE='VBScript'%>

<html>
	<head>
		<title>Captación de Clientes</title>
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
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td><div align="center"><img src="imagenes/construccion.gif"></div></td>
			</tr>
			<tr>
				<td><div align="center"><font color="#CC0099" size="4"><strong>ESTAMOS 
        TRABAJANDO PARA USTED</strong></font></div></td>
			</tr>
		</table>
        <table width="100%"  align="center">
            <tr>
            <td width="14%">				</td>
                <td width="86%">
                    <FORM NAME="volver" ACTION="menu.asp">
                        <input type="submit" name="" value="Regresar">
                    </FORM>
                </td>
            </tr>
        </table>
		
	</body>
</html>
