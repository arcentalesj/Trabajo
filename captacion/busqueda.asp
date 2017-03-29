<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/usuarios.asp" -->
<% 
select case session("accion")
	case "Busqueda de Prospectos"
		titulo="BUSQUEDA PROSPECTOS"
	case "Suspender Prospecto"
		titulo="SUSPENDER PROSPECTOS"
end select
If (Request.QueryString("asigna") <> "") then 
	session("accion")=Request.QueryString("asigna")
	if (Request.QueryString("cdemi")<>"") then
		session("equipo1")=Request.QueryString("cdemi")
	else
		session("equipo1")=Session("equipo1")
	end if
end If
If (Request.QueryString("mensajeerror") <> "") Then 
	session("dd") = Request.QueryString("mensajeerror")
End If
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = 0
session("presenta")=2
%>
<%
Dim vendedor
Dim vendedor_numRows
Set vendedor = Server.CreateObject("ADODB.Recordset")
vendedor.ActiveConnection = MM_usuarios_STRING
vendedor.CursorType = 0
vendedor.CursorLocation = 2
vendedor.LockType = 1
strSql= "SELECT * FROM asesores order by nombre"
vendedor.Source = strSql
vendedor.Open()
vendedor_numRows = 0
%>
<html>
<head>
	<title>Busqueda de Clientes</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language='JavaScript' type="text/javascript">
		function armar()
			{
			if (!isNumeric(formulario.cedula.value)) {
				alert("El documento de identificacion debe ser numérico.")
				return false;
			}
			if (!isNumeric(formulario.mes.value)) {
				alert("El campo del mes debe ser numérico.")
				return false;
			}
			if (!isNumeric(formulario.anio.value)) {
				alert("El campo del año debe ser numérico.")
				return false;
			}
			var campo1=(formulario.nombre.value);
			var campo2=(formulario.apellido.value);
			var campo3=(formulario.cedula.value);
			var campo4=(formulario.mes.value);
			var campo5=(formulario.anio.value);
			var campo6=(formulario.vende1.value);
			if ((campo1 == "") && (campo2 == "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0))
				{ 
				alert('NO ha ingresado datos para procesar la consulta');
				return false; 
			}
			if ((campo1 !== "") && (campo2 == "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0))
				{ 
				formulario.action = "consultatotal.asp?opcion=1&campo="+campo1,"_parent";
				var respuesta=window.open("consultatotal.asp?opcion=1&campo="+campo1,"_parent");
				return true;
			}
			if ((campo1 == "") && (campo2 !== "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0))
				{ 
				formulario.action = "consultatotal.asp?opcion=3&campo="+campo2,"_parent";
				var respuesta=window.open("consultatotal.asp?opcion=3&campo="+campo2,"_parent");
				return true;
			}
			if ((campo1 == "") && (campo2 == "") && (campo3 !== "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0))
				{ 
				formulario.action = "consultatotal.asp?opcion=4&campo="+campo3,"_parent";
				var respuesta=window.open("consultatotal.asp?opcion=4&campo="+campo3,"_parent");
				return true;
			}
			if ((campo1 == "") && (campo2 == "") && (campo3 == "") && (campo4 !== "") &&(campo5 == "") && (campo6 == 0))
				{ 
				formulario.action = "consultatotal.asp?opcion=6&campo="+campo4,"_parent";
				var respuesta=window.open("consultatotal.asp?opcion=6&campo="+campo4,"_parent");
				return true;
			}
			if ((campo1 == "") && (campo2 == "") && (campo3 == "") && (campo4 == "") &&(campo5 !== "") && (campo6 == 0))
				{ 
				formulario.action = "consultatotal.asp?opcion=7&campo="+campo5,"_parent";
				var respuesta=window.open("consultatotal.asp?opcion=7&campo="+campo5,"_parent");
				return true;
			}
			if ((campo1 == "") && (campo2 == "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 !== 0))
				{ 
				formulario.action = "consultatotal.asp?opcion=5&campo="+campo6,"_parent";
				var respuesta=window.open("consultatotal.asp?opcion=5&campo="+campo6,"_parent");
				return true;
			}
			if ((campo1 !== "") && (campo2 !== "") && (campo3 == "") && (campo4 == "") &&(campo5 == "") && (campo6 == 0))
				{ 
				formulario.action = "consultatotal.asp?opcion=2&campo="+campo1+"&campoa="+campo2,"_parent";
				var respuesta=window.open("consultatotal.asp?opcion=2&campo="+campo1+"&campoa="+campo2,"_parent");
				return true;
			}
			function isNumeric(valor) {
				var bNumeric=true;
				for (var ii = 0 ; ii < valor.length ; ii++) {
					if(0 > ('0123456789').indexOf(valor.substring(ii, ii+1))) {
						bNumeric=false;
						break;
					}
				}
				return bNumeric;
			}
		}
	</SCRIPT>
	<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
	<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	

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
	  <form name="formulario" method=post onSubmit='return armar();' >
			<input type=hidden name="pagina" value="resultado"/>
		<table width="33%" height="131" border="0" align="center" cellpadding="2" cellspacing="2">
			<tr> 
				<td width="40%" class="Estilo5">Por Nombres :</td>
			  <td width="60%" colspan="4"><font size="1">
		      <input name="nombre" type="text" value="" size="25" maxlength="30" tabindex="1"></font></td>
			</tr>
			<tr> 
				<td width="40%" class="Estilo5">Por Apellidos :</span></div></td>
				<td width="60%" colspan="4"><font size="1">
			  <input name="apellido" type="text" value="" size="25" maxlength="30" tabindex="2"></font></td>
			</tr>
			<tr> 
				<td width="40%"  class="Estilo5">Por Número Cédula :</span></div></td>
				<td width="60%" colspan="4"><font size="1">
			  <input id="cedula" name="cedula" type="text" value="" size="13" maxlength="13" tabindex="3"></font></td>
			</tr>
			<tr> 
				<td width="40%" class="Estilo5">Por Asesor Comercial :</span></div></td>
				<td width="60%" colspan="4"><font size="2"> 
					<select name="vende1" class="PPRFiel" id="select" tabindex="5" onChange="">
						<option value="0">Escoga Un Asesor Comercial</option>
						<%While (NOT vendedor.EOF)%>
							<option value="<%=(vendedor.Fields.Item("codven").Value)%>" <%If (Not isNull((session("vendedor")))) Then If (trim(vendedor.Fields.Item("codven").Value) = (trim(session("vendedor")))) Then Response.Write("SELECTED") : Response.Write("")%>><%=(vendedor.Fields.Item("nombre").Value)%></option>
							<%vendedor.MoveNext()
						Wend
						If (vendedor.CursorType > 0) Then
							vendedor.MoveFirst
						Else
							vendedor.Requery
						End If%>
					</select></font>
			  </td>
			</tr>
		</table>
		<table width="405" align="center" border="0" cellpadding="2" cellspacing="2">
			<tr> 
				<td width="180"><div align="right"><span class="estilo5">Mes y Año Captación :</span></div></td>
				<td width="70"><div align="right"><span class="estilo5">Mes</span></div></td>
				<td width="30" align="right"><span class="detalle">
					<input name="mes" type="text" id="mes" value="" size="2" maxlength="2" tabindex="7" style="TEXT-ALIGN: Center;"></span></td>
				<td width="39"><div align="right"><span class="estilo5">Año</span></div></td>
				<td width="38">
					<input name="anio" type="text" id="anio" value="" size="4" maxlength="4" tabindex="8" style="TEXT-ALIGN: right;"></td>

			</tr>
		</table>
		<p>&nbsp; </p>
		<table width="100%" align="center">
			<tr> 
				<td width="20%">&nbsp;</td>
				<td width="20%"> <div align="center">
				  <input type="button" class="boton" name="btnBuscar" value="Realizar Busqueda" onclick='javascript:armar();'/>
			  </div></td>
				<td width="20%">&nbsp;</td>
			</tr>
		</table>
	</form>
	<table width="100%"  align="center">
		<tr>
			<td width="26%"></td>
			<td width="26%">
				<FORM NAME="volver" ACTION="menu.asp">
					<input type="submit" name="" value="Regresar">
				</FORM>
			</td>
			<td width="26%"></td>
			<td width="26%"></td>
		</tr>
	</table>
	<table width="100%">
		<tr>
			<td width="100%" align="center"> 
				<p>
					<%
					select case request.querystring("mensajeerror")
						case "cedula_nohay"
							response.write ("<center><b>La información ingresada no coincide con ningún empleado</b></center>")	
						case "cclientet_nohay"
							response.write ("<center><b>No existe informacion para esta Consulta</b></center>")	

					end select
					%>
				</p>
			</td>
		</tr>
	</table>
</body>
</html>
<%
vendedor.Close()
Set vendedor = Nothing
%>
