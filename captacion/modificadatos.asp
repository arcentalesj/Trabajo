<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Connections/usuarios.asp" -->

<%
If (Request.QueryString("codigo") <> "") then 
	asesor1=Request.QueryString("codigo")
end If
titulo = "ACTUALIZACION DE DATOS DE CLIENTES"
strSQL = "SELECT * FROM dbo.captados where idvendedor = '"+ session("USUARIO")+ "' order by apellido,nombre"
%>
<%
Dim captados
Dim captados_numRows
Dim captados__MMColParam
captados__MMColParam = "1"
Set captados = Server.CreateObject("ADODB.Recordset")
captados.ActiveConnection = MM_usuarios_STRING
captados.Source = strSQL
captados.CursorType = 0
captados.CursorLocation = 2
captados.LockType = 1
captados.Open()
if captados.bof and captados.bof then
	response.redirect("menu.asp?mensajeerror=" &"cclientet_nohay")
end if
captados_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index
Repeat1__numRows = 20
Repeat1__index = 0
captados_numRows = captados_numRows + Repeat1__numRows
%>
<%
Dim captados_total
Dim captados_first
Dim captados_last
captados_total = captados.RecordCount
If (captados_numRows < 0) Then
	captados_numRows = captados_total
Elseif (captados_numRows = 0) Then
	captados_numRows = 1
End If
captados_first = 1
captados_last  = captados_first + captados_numRows - 1
If (captados_total <> -1) Then
	If (captados_first > captados_total) Then
   		captados_first = captados_total
	End If
	If (captados_last > captados_total) Then
		captados_last = captados_total
	End If
	If (captados_numRows > captados_total) Then
    	captados_numRows = captados_total
	End If
End If
%>
<%
If (captados_total = -1) Then
	captados_total=0
	While (Not captados.EOF)
    	captados_total = captados_total + 1
	    captados.MoveNext
	Wend
	If (captados.CursorType > 0) Then
		captados.MoveFirst
	Else
		captados.Requery
	End If
	If (captados_numRows < 0 Or captados_numRows > captados_total) Then
		captados_numRows = captados_total
	End If
	captados_first = 1
	captados_last = captados_first + captados_numRows - 1
	If (captados_first > captados_total) Then
		captados_first = captados_total
	End If
	If (captados_last > captados_total) Then
		captados_last = captados_total
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
Set MM_rs    = captados
MM_rsCount   = captados_total
MM_size      = captados_numRows
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
		If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
			If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
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
		MM_offset = MM_index  ' set MM_offset to the last possible record
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
captados_first = MM_offset + 1
captados_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
	If (captados_first > MM_rsCount) Then
		captados_first = MM_rsCount
	End If
	If (captados_last > MM_rsCount) Then
		captados_last = MM_rsCount
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
	<title>Clientes de un asesor para modificar datos</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
	<script language="JavaScript" type="text/javascript" src="js/calendar.js"></script>
	<script language="JavaScript" type="text/javascript" src="js/calendar-es.js"></script>
	<script language="JavaScript" type="text/javascript" src="js/calendar-setup.js"></script>
	<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">
	<link rel="stylesheet" type="text/css" href="css/calendario.css" >
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
	<table width="1000" border="1" bgcolor="#CCCCCC" align="center">
		<tr> 
			<td width="90" class="rotulos">Cedula</td>
			<td width="200" class="rotulos">Nombre</td>
			<td width="80" class="rotulos">Fecha Captacion</td>
			<td width="110" class="rotulos">Tipo de producto</td>
			<td width="200" class="rotulos">Punto de venta</td>
		</tr>
	</table>
	<table width="1000" border="1" align="center" cellspacing="0">
		<% 
		While ((Repeat1__numRows <> 0) AND (NOT captados.EOF)) 
			if isdate(captados.Fields.Item("contactado").Value) then
				fechacap = captados.Fields.Item("contactado")
				anio = year(fechacap)
				mes = month(fechacap)
				dia = day(fechacap)
				fechacap = dia &"/" &mes &"/" &anio
			end if
		%>
		<tr>
			<td width="90" class="presentac" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write (session("color")) end if%>">
	            <a href="modifica_prospectos.asp?
				<%= MM_keepBoth&MM_joinChar(MM_keepBoth)&"cedula="&(captados.Fields.Item("cedula").Value)&"&tipobien="&(captados.Fields.Item("idproducto").Value)%>"><%=(captados.Fields.Item("cedula").Value)%></a> 
			</td>
			<td width="200" class="presental" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write (session("color")) end if%>">
				<%=(captados.Fields.Item("apellido").Value)%><%=(response.write("  "))%><%=(captados.Fields.Item("Nombre").Value)%></td>
			<td width="80" class="presentar" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write (session("color")) end if%>">
				<%=(response.write(fechacap))%></td>
			<td width="110" class="presental" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write (session("color")) end if%>">
				<%=(response.write(captados.Fields.Item("producto").Value))%></td>
			<td width="200" class="presental" bgcolor="<%if (Repeat1__numRows mod 2)=0 then Response.Write (session("color")) end if%>">
				<%=(response.write(captados.Fields.Item("nomptoventa").Value))%></td>
		</tr>
		<% 
			Repeat1__index=Repeat1__index+1
			Repeat1__numRows=Repeat1__numRows-1
			captados.MoveNext()
		Wend
		%>
		</tr>
	</table>
	</div>
	<table border="0" width="50%" align="center">
		<tr> 
			<td width="23%" align="center"> <% If MM_offset <> 0 Then %>
				<a href="<%=MM_moveFirst%>"><img src="imagenes/primero.gif" width="38" height="26" border=0></a> 
		  <% End If ' end MM_offset <> 0 %> </td>
			<td width="31%" align="center"> <% If MM_offset <> 0 Then %>
				<a href="<%=MM_movePrev%>"><img src="imagenes/atras.gif" width="38" height="26" border=0></a> 
		  <% End If ' end MM_offset <> 0 %> </td>
			<td width="23%" align="center"> <% If Not MM_atTotal Then %>
				<a href="<%=MM_moveNext%>"><img src="imagenes/adelante.gif" width="38" height="26" border=0></a> 
		  <% End If ' end Not MM_atTotal %> </td>
			<td width="23%" align="center"> <% If Not MM_atTotal Then %>
				<a href="<%=MM_moveLast%>"><img src="imagenes/ultimo.gif" width="38" height="26" border=0></a> 
		  <% End If ' end Not MM_atTotal %> </td>
		</tr>
	</table>
	<td align="center">
    <div align="center"><strong><font size="2">Registros <%=(captados_first)%> al <%=(captados_last)%> de <%=(captados_total)%></font></strong></div></td>
    <table width="100%"  align="center">
        <tr>
            <td width="14%">				</td>
<!--			<%'select case session("tipocliente")%>
				<%'case "sup"%>
		            <td width="53%">
           			    <FORM NAME="volver" ACTION="muestra_asesores.asp">
	    	                <input type="submit" name="asigna" value="Modificacion de datos">
       			        </FORM>
					</td>
					<td width="33%">
        		        <FORM NAME="volver" ACTION="submenuclientes.asp">
                		    <input type="submit" name="" value="Menu Clientes">
		                </FORM>
					</td>
				<%'case "Asesor"%>-->
		            <td width="53%">
        	        <FORM NAME="volver" ACTION="menu.asp">
            	        <input type="submit" name="" value="Regresar">
                	</FORM>
					</td>
					<td width="33%">
				</td>
<!--				<%'case "reg"%>
		            <td width="53%">
           			    <FORM NAME="volver" ACTION="muestra_asesores.asp">
	    	                <input type="submit" name="asigna" value="Modificacion de datos">
       			        </FORM>
					</td>
				<td width="33%">
    	            <FORM NAME="volver" ACTION="submenuclientes.asp">
        	            <input type="submit" name="" value="Menu Clientes">
            	    </FORM>
				</td>
			<%'end select%>-->
      </tr>
    </table>
</body>
</html>
<%
captados.Close()
Set captados = Nothing
%>
