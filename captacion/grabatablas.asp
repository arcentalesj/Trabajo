<%@ LANGUAGE='VBScript'%>
<!-- #include file="Connections/usuarios.asp" -->
<%
Set tablas = CreateObject("ADODB.Recordset")
tablas.activeconnection = MM_usuarios_STRING
select case session("tabla")
	case 1
		if (session("motivo")) = 0 then
			strSQL2 ="insert puntoventa (nom_puntoventa, id_estatus)"
			strSQL2 = strSQL2 + " values ('" & UCASE(request.form("observa"))& "','A')"
		else
			strSQL2 ="update puntoventa set nom_puntoventa= '"& (ucase(request.form("observa"))) & "'" 
			strSQL2 = strSQL2 + " where id_puntoventa=" & (Session("codigo"))
		end if
end select
'response.write (strSql2)
tablas.Open strSQL2, MM_usuarios_STRING

%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
		<title>Modificación de Tablas</title>
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
		<table width="100%" height ="8%"><tr></tr></table>
		<table width="58%" border="1" align="center" bgcolor="#FFFFFF" bordercolor="#990000" cellspacing="0" cellpadding="0">
			<tr> 
				<td class="titulo">Ingreso de datos satisfactorios</td>
			</tr>
		</table>
		<table width="100%" height ="8%"><tr></tr></table>
		<table width="27%"  align="center">
			<tr> 
				<td width="26%" align="center">
					<%
					select case session("tabla")
						case 1 
							%> <FORM NAME="volver" ACTION="ingptoventa.asp">
								<input type="submit" name="" value="Ingresa o Modifica otro Punto de Venta">
							</FORM>
							<%
					end select
					%>
				</td> 
			</tr>
		</table>
		<table width="30%" border="0" align="left" cellspacing="0" cellpadding="0">
			<tr> 
				<td align="right"> 
					<FORM NAME="volver" ACTION="menu.asp">
						<input type="submit" name="" value="Regresar">
					</FORM>
				</td>
			</tr>
		</table>
	</body>
</html>
<%
set tablas = nothing
%>
