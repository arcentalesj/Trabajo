<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/usuarios.asp" -->
<%
titulo="INGRESAR O MODIFICAR TABLA DE PUNTOS DE VENTA"
session("tabla") = 1
Dim puntos__MMColParam
puntos__MMColParam = "1"
If (Request.QueryString("id_puntoventa") <> "") Then 
	puntos__MMColParam = Request.QueryString("id_puntoventa")
	if Request.QueryString("id_puntoventa") >= "1" then
		session("motivo") = 1
		session("codigo") = Request.QueryString("id_puntoventa")
	else
		session("motivo") = 0
	end if
End If
%>
<%
Dim puntos
Dim puntos_numRows

Set puntos = Server.CreateObject("ADODB.Recordset")
puntos.ActiveConnection = MM_usuarios_STRING
puntos.Source = "SELECT *  FROM dbo.puntoventa WHERE id_puntoventa = " + Replace(puntos__MMColParam, "'", "''") + ""
puntos.CursorType = 0
puntos.CursorLocation = 2
puntos.LockType = 3
puntos.Open()
%>

<html>
<head>
	<title>Modificación de Tablas</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
	<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">

	<SCRIPT LANGUAGE='VBScript'event='OnClick' for='graba'>
        if trim(form2.observa.value) <> "" then
                    form2.submit
        else
            msgbox "Nombre no puede estar en blanco"
        end if
    </SCRIPT>
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
	<table width="51%" align="center">
		<form method="post" action="grabatablas.asp" name="form2">
			<td colspan="2" class="estilo5">Nombre Punto de Venta : </b></font></td>
			<td width="55%" colspan="2"><input name="observa" type="text" value="<%response.write(puntos("nom_puntoventa"))%>" size="45"></td>
			<tr> 
				<td colspan="2">&nbsp;</td>
				<td colspan="3">&nbsp;</td>
			</tr>
			<tr>
			<table align="center">
				<tr>
					<td colspan="2" align="center"><input type="submit"  name="btnBuscar" value="Graba Datos"></button>
				</tr>
			</table>
		</form>
		<table width="100%" height ="5%"><tr></tr></table>
       <table width="100%"  align="center">
			<tr>
				<td width="14%">				</td>
                <td width="53%">
                    <FORM NAME="volver" ACTION="ingptoventa.asp">
                        <input type="submit" name="asigna" value="Regresar">
                    </FORM>
                </td>
			</tr>
        </table>
	</body>
</html>
<%
puntos.Close()
Set puntos = Nothing
%>

