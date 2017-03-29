<%@LANGUAGE="VBSCRIPT"%>
<% session("estado")= 5 %>
<!--#include file="Connections/usuarios.asp" -->
<%
If (Request.QueryString("cedula1") <> "") Then 
  session("cedula") = Request.QueryString("cedula1")
End If
 Response.addHeader "pragma", "no-cache"
 Response.CacheControl = "Private"
 Response.Expires = 0
%>
<%
titulo = "REASIGNACION DE ASESOR COMERCIAL"
usuario=session("quien")
Dim captados
Dim captados_numRows
Set captados = Server.CreateObject("ADODB.Recordset")
captados.ActiveConnection = MM_usuarios_STRING
captados.Source = "SELECT *  FROM reasigna_Asesor  WHERE cedula = '" + session("cedula") +"'"
captados.CursorType = 0
captados.CursorLocation = 2
captados.LockType = 1
captados.Open()
captados_numRows = 0
color="#CCCCCC"
%>
<%
Dim obsseguim
Dim obsseguim_numRows
Set obsseguim = Server.CreateObject("ADODB.Recordset")
obsseguim.ActiveConnection = MM_usuarios_STRING
obsseguim.Source = "SELECT *  FROM dbo.obsseguim  WHERE obs_cedula ='"&(session("cedula"))&"'"
obsseguim.CursorType = 0
obsseguim.CursorLocation = 2
obsseguim.LockType = 1
obsseguim.Open()
obsseguim_numRows = 0
%>
<%
Dim vendedor
Dim vendedor_numRows
Set vendedor = Server.CreateObject("ADODB.Recordset")
vendedor.ActiveConnection = MM_usuarios_STRING
vendedor.CursorType = 0
vendedor.CursorLocation = 2
vendedor.LockType = 1
vendedor.Source = "SELECT *  FROM asesores where cargo = '71'"
vendedor.Open()
if vendedor.bof and vendedor.bof then
	response.redirect("menu.asp?mensajeerror=" &"vende_nohay")		
end if
session("cambio") = 1
vendedor_numRows = 0
%>
<%
Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd
Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId
Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
	MM_editAction = MM_editAction & "?" & Request.QueryString
End If
MM_abortEdit = false
MM_editQuery = ""
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = 10
Repeat1__index = 0
obsseguim_numRows = obsseguim_numRows + Repeat1__numRows
%>
<%
Dim obsseguim_total
Dim obsseguim_first
Dim obsseguim_last
obsseguim_total = obsseguim.RecordCount
If (obsseguim_numRows < 0) Then
	obsseguim_numRows = obsseguim_total
Elseif (obsseguim_numRows = 0) Then
	obsseguim_numRows = 1
End If
obsseguim_first = 1
obsseguim_last  = obsseguim_first + obsseguim_numRows - 1
If (obsseguim_total <> -1) Then
	If (obsseguim_first > obsseguim_total) Then
		obsseguim_first = obsseguim_total
	End If
	If (obsseguim_last > obsseguim_total) Then
		obsseguim_last = obsseguim_total
	End If
	If (obsseguim_numRows > obsseguim_total) Then
		obsseguim_numRows = obsseguim_total
	End If
End If
%>
<%
If (obsseguim_total = -1) Then
	obsseguim_total=0
	While (Not obsseguim.EOF)
		obsseguim_total = obsseguim_total + 1
		obsseguim.MoveNext
	Wend
	If (obsseguim.CursorType > 0) Then
		obsseguim.MoveFirst
	Else
		obsseguim.Requery
	End If
	If (obsseguim_numRows < 0 Or obsseguim_numRows > obsseguim_total) Then
		obsseguim_numRows = obsseguim_total
	End If
	obsseguim_first = 1
	obsseguim_last = obsseguim_first + obsseguim_numRows - 1
	If (obsseguim_first > obsseguim_total) Then
		obsseguim_first = obsseguim_total
	End If
	If (obsseguim_last > obsseguim_total) Then
		obsseguim_last = obsseguim_total
	End If
End If
%>
<%
Dim MM_paramName 
%>
<%
Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined
Dim MM_param
Dim MM_index
Set MM_rs    = obsseguim
MM_rsCount   = obsseguim_total
MM_size      = obsseguim_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
	MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
if (Not MM_paramIsDefined And MM_rsCount <> 0) then
	MM_param = Request.QueryString("index")
	If (MM_param = "") Then
		MM_param = Request.QueryString("offset")
	End If
	If (MM_param <> "") Then
		MM_offset = Int(MM_param)
	End If
	If (MM_rsCount <> -1) Then
		If (MM_offset >= MM_rsCount Or MM_offset = -1) Then
			If ((MM_rsCount Mod MM_size) > 0) Then  
				MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
			Else
				MM_offset = MM_rsCount - MM_size
			End If
		End If
	End If
	MM_index = 0
	While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
		MM_rs.MoveNext
		MM_index = MM_index + 1
	Wend
	If (MM_rs.EOF) Then 
		MM_offset = MM_index  
	End If
End If
%>
<%
If (MM_rsCount = -1) Then
	MM_index = MM_offset
	While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
		MM_rs.MoveNext
		MM_index = MM_index + 1
	Wend
	If (MM_rs.EOF) Then
		MM_rsCount = MM_index
		If (MM_size < 0 Or MM_size > MM_rsCount) Then
			MM_size = MM_rsCount
		End If
	End If
	If (MM_rs.EOF And Not MM_paramIsDefined) Then
		If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
			If ((MM_rsCount Mod MM_size) > 0) Then
				MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
			Else
				MM_offset = MM_rsCount - MM_size
			End If
		End If
	End If
	If (MM_rs.CursorType > 0) Then
		MM_rs.MoveFirst
	Else
		MM_rs.Requery
	End If
	MM_index = 0
	While (Not MM_rs.EOF And MM_index < MM_offset)
		MM_rs.MoveNext
		MM_index = MM_index + 1
	Wend
End If
%>
<%
obsseguim_first = MM_offset + 1
obsseguim_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
	If (obsseguim_first > MM_rsCount) Then
		obsseguim_first = MM_rsCount
	End If
	If (obsseguim_last > MM_rsCount) Then
		obsseguim_last = MM_rsCount
	End If
End If
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth
Dim MM_removeList
Dim MM_item
Dim MM_nextItem
MM_removeList = "&index="
If (MM_paramName <> "") Then
	MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If
MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""
For Each MM_item In Request.QueryString
	MM_nextItem = "&" & MM_item & "="
	If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
		MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
	End If
Next
For Each MM_item In Request.Form
	MM_nextItem = "&" & MM_item & "="
	If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
		MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
	End If
Next
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
	MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
	MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
	MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If
Function MM_joinChar(firstItem)
	If (firstItem <> "") Then
		MM_joinChar = "&"
	Else
		MM_joinChar = ""
	End If
End Function
%>
<%
Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev
Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam
MM_keepMove = MM_keepBoth
MM_moveParam = "index"
If (MM_size > 1) Then
	MM_moveParam = "offset"
	If (MM_keepMove <> "") Then
		MM_paramList = Split(MM_keepMove, "&")
		MM_keepMove = ""
		For MM_paramIndex = 0 To UBound(MM_paramList)
			MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
			If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
				MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
			End If
		Next
		If (MM_keepMove <> "") Then
			MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
		End If
	End If
End If
If (MM_keepMove <> "") Then 
	MM_keepMove = MM_keepMove & "&"
End If
MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
	MM_movePrev = MM_urlStr & "0"
Else
	MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<html>
	<head>
		<title>Consulta datos Clientes</title>
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
        <table width="895"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
            <tr bgcolor=#CCCCCC> 
                <td width="12"></td>
                <td width="139" class="presenta1">Cedula</td>
                <td width="288" class="presenta2"><%response.write(session("cedula"))%></td>
                <td width="13">&nbsp;</td>
                <td width="15">&nbsp;</td>
                <td width="139" class="presenta1">Fecha de Registro</td>
                <td width="289" class="presenta2"><%response.write(captados("ingresado"))%></td>
            </tr>
            <tr> 
                <td width="12"></td>
                <td width="139" class="presenta1">Apellidos</b></td>
                <td width="288" class="presenta2"><%response.write(captados("apellido"))%></td>
                <td width="13">&nbsp;</td>
                <td width="15"></td>
                <td width="139" class="presenta1">Nombres</b></td>
                <td width="289" class="presenta2"><%response.write(captados("nombre"))%></td>
            </tr>
            <tr bgcolor=#CCCCCC> 
                <td width="12"></td>
                <td width="139" class="presenta1">Direccion Domicilio</td>
                <td width="288" class="presenta2"><%response.write(TRIM(captados("direccion")))%></td>
                <td width="13">&nbsp;</td>
                <td width="15">&nbsp;</td>
                <td width="139" class="presenta1">Direccion Trabajo</b></td>
                <td width="289" class="presenta2"><%response.write(captados("trabajo"))%></td>
            </tr>
            <tr> 
                <td width="12">&nbsp;</td>
                <td width="139" class="presenta1">E_Mail</td>
                <td width="288" class="presenta2"><%response.write(captados("correo"))%></td>
                <td width="13">&nbsp;</td>
                <td width="15">&nbsp;</td>
                <td width="139" class="presenta1">Celular</b></td>
                <td width="289" class="presenta2"><%response.write(captados("celular"))%></td>
            </tr>
            <tr bgcolor=#CCCCCC> 
                <td width="12"></td>
                <td width="139" class="presenta1">Telefono 1</b></td>
                <td width="288" class="presenta2"><%response.write(captados("codarea1"))%><%response.write("-")%><%response.write(captados("telefono"))%></td>
                <td width="13">&nbsp;</td>
				<td width="15">&nbsp;</td>
              <td width="139" class="presenta1">Telefono2</b></td>
                <td width="289" class="presenta2"><%response.write(captados("codarea2"))%><%response.write("-")%><%response.write(captados("fono2"))%></td>
          </tr>
            <tr> 
               <td width="12">&nbsp;</td>
                <td width="139" class="presenta1">Estado Civil</td>
                <td width="288" class="presenta2"><%response.write(captados("nomestado"))%></td>
                <td width="13">&nbsp;</td>
                <td width="15"></td>
                <td width="139" class="presenta1">Cotizador</b></td>
                <td width="289" class="presenta2"><%'response.write(captados("cotizador"))%></td>
            </tr>
            <tr bgcolor=#CCCCCC> 
                <td width="12">&nbsp;</td>
                <td width="139" class="presenta1">Descripcion del Bien</td>
                <td width="288" class="presenta2"><%response.write(captados("idproducto"))%><%response.write(" - ")%><%response.write(captados("nomproducto"))%></td>

                <td width="13">&nbsp;</td>
                <td width="15"></td>
                <td width="139" class="presenta1">Punto de Venta</td>
                <td width="289" class="presenta2"><%response.write(captados("nomptoventa"))%></td>
            </tr>
        </table>
        <table width="700"  border="0" align="center" cellpadding="1" cellspacing="0" >		
            <tr> 
                <td width="61">&nbsp;</td>
                <td width="122"><div align="left"><font color="#990000" size="2"><b></b></font></div></td>
                <td width="122"><div align="left"><font color="#990000" size="2"><b></b></font></div></td>
                <td width="122"><div align="left"><font color="#990000" size="2"><b></b></font></div></td>
            </tr>
            <tr> 
                <td width="61">&nbsp;</td>
                <td width="122"><div align="left"><font color="#990000" size="2"><b>Primera</b></font></div></td>
                <td width="122"><div align="left"><font color="#990000" size="2"><b>Segunda</b></font></div></td>
                <td width="122"><div align="left"><font color="#990000" size="2"><b>Tercera</b></font></div></td>
            </tr>
            <tr> 
                <td width="61" class="presenta1"> CITAS</font></strong></td>
                <td width="122" class="presenta2"><%response.write(captados("cita1"))%></font></td>
                <td width="122" class="presenta2"><%'response.write(fecita2)%></font></td>
                <td width="127" class="presenta2"><%'response.write(fecita3)%></font></td>
                <td width="40" class="presenta2"><%'response.write(cuentaci)%></font></td>
            </tr>
        </table>
		<table width="884"  border="1" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
			<%While ((Repeat1__numRows <> 0) AND (NOT obsseguim.EOF))%>
				<tr valign="top">
					<td width="100" align="center" bgcolor="<%if ( Repeat1__numRows mod 2)=0 then Response.Write color end if%>"
                    ><font size="2" ><%response.write(obsseguim("obs_fecha"))%></font></td>
					<td width="847" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write color end if%>">
                    <font size="2" ><%response.write(obsseguim("obs_observa"))%></font></td>
				</tr>
				<%
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				obsseguim.MoveNext()
			Wend
			%>
		</table>
		<table width="82%" border="0" align="center">
			<tr> 
				<td width="198"><font color="#990000" size="1">&nbsp;</font></td>
			</tr>
		</table>
		<form method="post" action="grabadatos.asp" name="form4">
			<table width="800" border="0" align="center" cellpadding="0" cellspacing="0"  bgcolor="#FFFFFF">
				<tr> 
				<td width="120" class="presenta1">Asesor Anterior</td>
				<td width="180" class="presenta2"><%response.write(captados("nomvende"))%></td>
					<td width="100" class="presenta1">Asesor Nuevo</td>
					<td width="180" class="presenta2">
						<select name="vende1" class="PPRFiel" id="select" tabindex="10" >
							<%While (NOT vendedor.EOF)%>
								<option value="<%=(vendedor.Fields.Item("codven").Value)%>"
                                 <%If (Not isNull((session("vendedor")))) Then If (trim(vendedor.Fields.Item("codven").Value) = (trim(session("vendedor")))) Then Response.Write("SELECTED") : Response.Write("")%> ><%=(vendedor.Fields.Item("nombre").Value)%></option>
								<%vendedor.MoveNext()
							Wend
							If (vendedor.CursorType > 0) Then
								vendedor.MoveFirst
							Else
								vendedor.Requery
							End If%>
						</select>
					</td>
				<tr>
			</table>
			<table align="center">
				<tr>
					<td width="33%" align="center">&nbsp;</td>
					<td width="33%" align="center">&nbsp;</td>
					<td colspan="2" align="center"><input type="submit"  name="btnBuscar" value="Reasigna Asesor"></button>
				</tr>
			</table>
			
		</form>
		<table width="884"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
			<tr valign="top">
				<td width="100" align="center"><strong><font size="2" font color="#990000">Fecha Cita</font></strong></td>
				<td width="820" align="center"><strong><font size="2" font color="#990000">Observación Cita</font></strong></td>
				<td width="27" align="center"><strong><font size="2" font color="#990000">Eval</font></strong></td>
			</tr>
		</table>
		<table width="884"  border="1" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
			<%While ((Repeat1__numRows <> 0) AND (NOT obsseguim.EOF))%>
				<tr valign="top">
					<td width="100" align="center" bgcolor="
					<%if ( Repeat1__numRows mod 2)=0 then Response.Write color end if%>"><font size="2" ><%response.write(obsseguim("obs_fecha"))%></font></td>
					<td width="820" bgcolor="
					<%if (Repeat1__numRows mod 2)=0 then Response.Write color end if%>"><font size="2" ><%response.write(obsseguim("obs_observa"))%></font></td>
					<td width="27" bgcolor="
					<%if (Repeat1__numRows mod 2)=0 then Response.Write color end if%>"><font size="2" ><%response.write(obsseguim("obs_evalua"))%></font></td>
				</tr>
				<%' 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				obsseguim.MoveNext()
			Wend
			%>
		</table>
		<!-- aqui regresa a la pagina inicial de control de citas -->
        <table width="82%" border="0" align="center">
            <tr> 
                <td width="198"><font color="#990000" size="2"><b><br></b></font></td>
                <td width="164"> <font size="2">&nbsp;</font></td>
                <td width="148">&nbsp;</td>
                <td width="272">&nbsp;</td>
                <td width="272">&nbsp;</td>
                <td width="272"><div align="left"><font color="#990000" size="2"><b>&nbsp;</b></font></div></td>
            </tr>
        </table>
        <table width="100%"  align="center">
            <tr>
                <td width="14%">				</td>
                <td width="53%">
                    <FORM NAME="volver" ACTION="menu.asp">
                        <input type="submit" name="" value="Regresar">
                    </FORM>
                </td>
                <td width="33%">
                    <FORM NAME="volver" ACTION="cambia_asesor.asp">
                        <input type="submit" name="asigna" value="Reasignacion de Vendedor">
                    </FORM>
                </td>
          </tr>
        </table>
	</body>
</html>
<%
'captados.Close()
'Set captados = Nothing
%>
<%
'vendedor.Close()
'Set vendedor = Nothing
%>
<%
'obsseguim.Close()
'Set obsseguim = Nothing
%>
