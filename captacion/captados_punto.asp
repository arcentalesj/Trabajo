<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<% 
   Response.addHeader "pragma", "no-cache"
   Response.CacheControl = "Private"
   Response.Expires = 0
%>
<%
 cboMeses=month(date())
 nAnio=year(date())
titulo = "CLIENTES CAPTADOS POR PUNTO DE VENTA"
%>
<html>
<head>
	<title>Envio de Informacion A excel</title>
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
    <%response.write(trim(cstr(nMes)))%>
	<form method="post" action="captados_punto1.asp" name="form4" AUTOCOMPLETE="OFF">
		<table border="0" width="15%" id="AutoNumber9" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" align="center">
			<tr>
				<td width="30%" class="presenta1">Mes :</td>
				<td width="100%" class="presenta2">
					<select id="cboMeses"  name="cboMeses"> 
						<option value="0">No Seleccionado</option>
								<OPTION value="1" selected>ENERO</OPTION> 
								<OPTION value="2" selected>FEBRERO</OPTION>
								<OPTION value="3" selected>MARZO</OPTION> 
								<OPTION value="4" selected>ABRIL</OPTION> 
								<OPTION value="5" selected>MAYO</OPTION> 
								<OPTION value="6" selected>JUNIO</OPTION> 
								<OPTION value="7" selected>JULIO</OPTION> 
								<OPTION value="8" selected>AGOSTO</OPTION> 
								<OPTION value="9" selected>SEPTIEMBRE</OPTION> 
								<OPTION value="10" selected>OCTUBRE</OPTION> 
								<OPTION value="11" selected>NOVIEMBRE</OPTION> 
								<OPTION value="12" selected>DICIEMBRE</OPTION> 
					</select>
				</td>
			</tr>
			<tr>
				<td width="30%" class="presenta1">Año:</font></b></td>
				<td width="83%" class="presenta2">
					<select id="cboAnio"  name="cboAnio"> 
						<option value="0" selected>No Seleccionado</option>
							<%i=2015
							do while i<=nAnio%>
								<%if trim(cstr(i))=trim(cstr(nAnio)) then%>
									<OPTION value=<%=i%> selected><%=i%></OPTION> 
								<%else%>
									<OPTION value=<%=i%> ><%=i%></OPTION> 
								<%end if%> 
								<%i=i+1
								loop%>
					</select>
				</td>
			</tr>
		</table>
		<table align="center">
           	<td>
			</td>
		</table>
		<table align="center">
			<tr>
				<td><input type="submit"  name="btnBuscar" value="Procesar"></button>
			</tr>
		</table>
		</form>
		<table width="100%"  align="center">
			<tr>
				<td width="26%" align="center">
					<FORM NAME="volver" ACTION="menu.asp">
						<input type="submit" name="" value="Regresar">
					</FORM>
				</td>
				<td width="26%" align="center">
				</td>
			</tr>
		</table>
</body>
</html>
