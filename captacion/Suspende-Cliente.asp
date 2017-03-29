<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/usuarios.asp" -->
<% 
   Response.addHeader "pragma", "no-cache"
   Response.CacheControl = "Private"
   Response.Expires = 0
%>
	
<%
Dim vendedor
Dim vendedor_numRows
Set vendedor = Server.CreateObject("ADODB.Recordset")
vendedor.ActiveConnection = MM_usuarios_STRING
vendedor.CursorType = 0
vendedor.CursorLocation = 2
vendedor.LockType = 1
'if session("codemi") >0 then
'	strsql="SELECT * FROM asesores WHERE cdequipo IN("+session("equipo")&")"
'"SELECT * FROM asesores WHERE cdequipo="+session("equipo") 
'else
	STRSQL= "SELECT * FROM asesores order by nombre"
'	session("vendedor")=0
'end if
'response.write(session("codemi"))
vendedor.Source = strsql
vendedor.Open()
vendedor_numRows = 0
%>
<html>
<head>
<title>Busqueda de Clientes</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
	<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
</head>

<body class="cuerpo">
	<table border="0" class="gradient" cellpadding="0" cellspacing="0" width="100%">
		<tr> 
			<td width="55%" rowspan="7" ><div align="right"></div></td>
			<td width="15%" height="50%">&nbsp;</td>
			<td width="20%" rowspan="5" class="gradient"><img name="texto" src="imagenes/logo.gif" width="249" height="153" border="0" alt=""></td>
		</tr>
		<tr>
			<td height="50%"></td>
		</tr>
		<tr>
			<td width="15%" height="50%">
				<SCRIPT>
                    dows = new Array("Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado");
                    months = new Array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre");
                    now = new Date();
                    dow = now.getDay();
                    d = now.getDate();
                    m = now.getMonth();
                    h = now.getTime();
                    y = now.getFullYear();
                    document.write(dows[dow]+" "+d+" de "+months[m]+" de "+y);
                   </SCRIPT></div>
			</td>
		</tr>
			<td height="100%">&nbsp;</td>
		<tr>
		</tr>
	</table>
	<table width="100%" border="1" cellpadding="0" cellspacing="0" class="Estilo2">
		<tr></tr>
	</table>
	<table width="100%" height="30" border="0" cellpadding="0" cellspacing="0">
		<tr> 
			<td align="center" valign="middle" class="importext">Busqueda de clientes</td>
		</tr>
	</table>
	  <form name="formulario" method=post onSubmit='return armar();' >
			<input type=hidden name="pagina" value="resultado"/>

<!--	<form name="form1" method="post" action="estosi.asp">-->
		<table width="518" height="131" border="0" align="center" cellpadding="2" cellspacing="2">
			<tr> 
				<td width="204"><div align="right"><span class="detalle">Por Nombres :</span></div></td>
				<td colspan="4"><font size="1"><input name="nombre" type="text" value="" size="25" maxlength="30" tabindex="1"></font></td>
			</tr>
			<tr> 
				<td width="204"><div align="right"><span class="detalle">Por Apellidos :</span></div></td>
				<td colspan="4"><font size="1"><input name="apellido" type="text" value="" size="25" maxlength="30" tabindex="2"></font></td>
			</tr>
			<tr> 
				<td width="204"><div align="right"><span class="detalle">Por Número Cédula :</span></div></td>
				<td colspan="4"><font size="1"><input id="cedula" name="cedula" type="text" value="" size="13" maxlength="13" tabindex="3"></font></td>
			</tr>
			<tr> 
				<td width="204"><div align="right"><span class="detalle">Por Asesor Comercial :</span></div></td>
				<td colspan="4"><font size="2"> 
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
		<table width="518" align="center" border="0" cellpadding="2" cellspacing="2">
			<tr> 
				<td width="211"><div align="right"><span class="detalle">Mes y Año Captación :</span></div></td>
				<td width="40"><span class="detalle">
                <input name="mes" type="text" id="mes" value="" size="2" maxlength="2" tabindex="7" style="TEXT-ALIGN: Center;"></span></td>
				<td width="51"><div align="Left"><span class="detalle">Mes</span></div></td>
				<td width="37"><span class="detalle">
                <input name="anio" type="text" id="anio" value="" size="4" maxlength="4" tabindex="8" style="TEXT-ALIGN: right;"></span></td>
				<td width="147"><div align="Left"><span class="detalle">Año</span></div></td>
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
<!--				<td colspan="7" align="center"> <input name="Enviar" type="button" onClick="armar()" value="Realizar Busqueda"> -->
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
</body>
</html>
<%
vendedor.Close()
Set vendedor = Nothing
%>
