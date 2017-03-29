<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/usuarios.asp" -->
<%
If (Request.QueryString("cedula") <> "") Then 
  session("cedulac") = Request.QueryString("cedula")
End If
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
Response.Expires = 0
titulo = "CONSULTA INDIVIDUAL DE PROSPECTOS"
StrSql= "SELECT *  FROM captados  WHERE cedula = '" + session("cedulac") +"'"
usuario=session("quien")
Dim captados
Dim captados_numRows
Set captados = Server.CreateObject("ADODB.Recordset")
captados.ActiveConnection = MM_usuarios_STRING
captados.Source = StrSql
captados.CursorType = 0
captados.CursorLocation = 2
captados.LockType = 1
captados.Open()
if captados.eof and captados.bof then
	response.redirect("busqueda.asp?mensajeerror=" &"cclientet_nohay")		
end if
captados_numRows = 0
color="#CCCCCC"
%>
<%
Dim obsseguim
Dim obsseguim_numRows
Set obsseguim = Server.CreateObject("ADODB.Recordset")
obsseguim.ActiveConnection = MM_usuarios_STRING
obsseguim.Source = "SELECT *  FROM dbo.obsseguim  WHERE obs_cedula ='"&(session("cedulac"))&"' order by obs_fecha"
obsseguim.CursorType = 0
obsseguim.CursorLocation = 2
obsseguim.LockType = 1
obsseguim.Open()
obsseguim_numRows = 0
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
		<title>Clientes Contactados por Call Center</title>
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
			<tr><td height="50%"></td></tr>
			<tr>
				<td class="fecha"><script type="text/javascript">fecha();</script></td>
			</tr>
			<tr>
				<td height="100%">&nbsp;</td>
			</tr>
		</table>
		<table cellspacing="0" class="Estilo2"><tr></tr></table>
		<p>&nbsp;</p>
		<table align="center">
			<tr> 
				<td class="titulo"><%response.write(titulo)%><%response.write("  ")%><%response.write(titulo1)%></td>
			</tr>
		</table>
		<table width="100%" height="3%" border="0" ><tr></tr></table>
		<table width="70%" border="1" align="center">
			<tr>
				<td width="19%" class="presenta1">Asesor Comercial</td>
				<td width="5%" class="presenta2"><%response.write(captados("idvendedor"))%></td>
				<td width="37%" class="presenta2"><%response.write(captados("nombrevendedor"))%></td>
				<td width="52%" class="presenta2"><%response.write(captados("nomptoventa"))%></td>
			</tr>
		</table>
		<table width="100%" height="3%" border="0" ><tr></tr></table>
		<table width="70%"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
			<tr bgcolor="<%Response.Write(session("color"))%>"> 
				<td width="121" class="presenta1">Cedula<br></td>
				<td width="331" class="presenta2"><%response.write(session("cedulac"))%></td>
				<td width="10">&nbsp;</td>
				<td width="127" class="presenta1">Fecha Captación</td>
				<td width="309" class="presenta2"><%response.write(captados("contactado"))%></td>
			</tr>
			<tr> 
				<td width="121" class="presenta1">Nombres</td>
				<td width="331" class="presenta2"><%response.write(captados("nombre"))%></td>
				<td width="10">&nbsp;</td>
				<td width="127" class="presenta1">Apellidos</td>
				<td width="309" class="presenta2"><%response.write(captados("apellido"))%></td>
			</tr>
			<tr bgcolor="<%Response.Write(session("color"))%>"> 
				<td width="121" class="presenta1">Dir. Domicilio</td>
				<td width="331" class="presenta2"><%response.write(captados("direccion"))%></td>
				<td width="10">&nbsp;</td>
				<td width="127" class="presenta1">Dir.Trabajo</td>
				<td width="309" class="presenta2"><%response.write(captados("trabajo"))%></td>
			</tr>
			<tr> 
				<td width="121" class="presenta1">E_Mail</td>
				<td width="331" class="presenta2"><%response.write(captados("correo"))%></td>
				<td width="10">&nbsp;</td>
				<td width="127" class="presenta1">Celular</td>
				<td width="309" class="presenta2"><%response.write(captados("celular"))%></td>
			</tr> 
			<tr bgcolor="<%Response.Write(session("color"))%>"> 
				<td width="121" class="presenta1">Telefono 1.</td>
				<td width="331" class="presenta2"><%response.write(captados("codarea1"))%><%response.write(captados("telefono"))%></td>
				<td width="10">&nbsp;</td>
				<td width="127" class="presenta1">Telefono 2.</td>
				<td width="309" class="presenta2"><%response.write(captados("codarea2"))%><%response.write(captados("fono2"))%></td>
			</tr>
			<tr>  
				<td width="121" class="presenta1">Estado Civil</td>
				<td width="331" align="left" class="presenta2"><%response.write(captados("nomestado"))%></td>
				<td width="10">&nbsp;</td>
				<td width="127" class="presenta1">Edad</td>
				<td width="309" class="presenta2"><%response.write(captados("edad"))%><%response.write(" - ")%><%response.write("Años")%></td>
			</tr>
			<tr bgcolor="<%Response.Write(session("color"))%>"> 
				<td width="121" class="presenta1">Producto</td>
				<td width="331"  class="presenta2"><%response.write(captados("producto"))%> </td>
				<td width="10">&nbsp;</td>
				<td width="127" class="presenta1">Cita</td>
				<td width="309"  class="presenta2"><%response.write(captados("cita"))%></td>
			</tr>
		</table>
		<table width="100%" height="3%" border="0" ><tr></tr></table>
		<table width="900"  border="0" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
			<tr valign="top">
				<td width="80" class="rotulos">Fecha Cita</td>
				<td width="750"class="rotulos">Observación Cita</td>
				<td width="50" class="rotulos">Evaluación</td>
			</tr>
		</table>
		<table width="884"  border="1" align="center" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
			<%While ((Repeat1__numRows <> 0) AND (NOT obsseguim.EOF))%>
				<tr valign="top">
					<td width="80" align="center" bgcolor="<%if ( Repeat1__numRows mod 2)=0 then Response.Write color end if%>"><font size="2" >
					<%response.write(obsseguim("obs_fecha"))%></font></td>
					<td width="750" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write color end if%>"><font size="2" >
					<%response.write(obsseguim("obs_observa"))%></font></td>
					<td width="50" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write color end if%>"><font size="2" >
					<%response.write(obsseguim("obs_evalua"))%></font></td>
				</tr>
				<% 
				Repeat1__index=Repeat1__index+1
				Repeat1__numRows=Repeat1__numRows-1
				obsseguim.MoveNext()
			Wend%>
		</table>
		<p>&nbsp;</p>
		<table width="941"  align="center">
			<td width="20%"></td>
			<%select case session("presenta")%>
				<%case 1%>
	        	    <td width="30%">
       	    		    <FORM NAME="volver" ACTION="clientes_ingresados.asp">
  	        	    		    <input type="submit" name="" value="Regresar">
	   	            	</FORM>
					</td>
				<%case 2%>
					<td width="30%">
						<FORM NAME="actualizar" ACTION="busqueda.asp">
							  <input type="submit" name="asigna" value="Busqueda de Prospectos">
						</FORM>
					</td>
			<%end select%>
			<td width="20%"> </td>
		</table>
	</body>
</html>
<%
captados.Close()
Set captados = Nothing
obsseguim.Close()
Set obsseguim = Nothing
%>
