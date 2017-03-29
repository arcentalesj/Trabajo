<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Response.Expires = -1
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
hoy = date()
dia = day(hoy)
anio1 = year(hoy)
mes1 = month(hoy)
hoy = anio&"/"&mes&"/"&dia
hora= time()
cestado=hoy+" "+cstr(hora)
%>
<html>
<head>
	<title>Graficos para gerencia</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="JavaScript" type="text/javascript" src="js/extras.js"></script>	
	<link rel="stylesheet" type="text/css" href="css/estilosCSC.CSS">

</head>

<p align="center"><b><font size="3"></font></b>
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

	<%sub displayverticalgraph(strtitle,strytitle,strxtitle,avalues,alabels)
	//'************************************************************************
	//'           user customizeable values for different formats
	//'			of the graph
	//' 			FREEWARE!!
	//'	Just tell me if you are going to use it - info@cool.co.za
	//'************************************************************************
	const GRAPH_HEIGHT 	= 300 	 		//'set up the graph height
	const GRAPH_WIDTH 	= 500   		//'set up the graph width
	const GRAPH_SPACING = 8	
	const GRAPH_BORDER 	= 0	 			//'if you would like to see the borders to align things differently	
	const GRAPH_BARS 	= 10    		//'loops through different colored bars e.g. 2 bars = gold,blue,gold,blue
	const USELOWVALUE 	= false 		//'this uses the low value of the array as the value of origin (default = 0)
	const SHOWLABELS 	= TRUE  		//'set this to toggle whether or not the labels are shown
	const L_LABEL_SEPARATOR = ""		//' |Label
	const R_LABEL_SEPARATOR = ""		//' Label|
	const LABELSIZE 	= -4			
	const GRAPHBORDERSIZE 	= 2
	const INTIMGBORDER 	= 1  	 		//'border around the bars
	Const ALT_TEXT		= 3				//'Changes the format of the alternate text of the bar image
										//'1 = Labels ,2 = Values , 3 = Labels + Values , 4 = Percent
	//'************************************************************************
	//'array of different bars to loop through
	//'you can change the order of these 
	//'Count = 10 
	//'"dark_green","red","gold","blue","pink","light_blue","light_gold","orange","green","purple"
	//'cut and paste from here and insert into the agraph_bars array below Make sure the 
	//'number specified in the const GRAPH_BARS is the same as or less than that in the array 
	//' 7 graph_bars <= 7 elements in array
	//'************************************************************************
	agraph_bars = array("imagenes/dark_green","imagenes/red","imagenes/gold","imagenes/blue","imagenes/pink","imagenes/light_blue","imagenes/light_gold","imagenes/orange","imagenes/green","imagenes/purple")
	intmax = 0
	//'find the maximum value of the values array
	for i = 0 to ubound(avalues)
		if  cint(intmax) < cint(avalues(i)) then  intmax = cint(avalues(i)) 
	next
	if uselowvalue then 
		intmin = avalues(0)
		for i = 0 to ubound(avalues)
			if  cint(intmin) > cint(avalues(i)) then  intmin = cint(avalues(i)) 
		next
	end if
	//'establish the graph multiplier
	graphmultiplier = round(graph_height-100/intmax)
	
	imgwidth = round(300/(ubound(avalues)))
	if imgwidth > 16 then imgwidth = 16 
	%>
	<table border =<%=GRAPH_BORDER%> width:100% height=<%=graph_height%>>
		<tr>
			<td rowspan=3 valign="middle"><%=strytitle%> </td>
			<td colspan=<%=ubound(avalues)+2%> height=50 align="center">
			<h4><%=strtitle%></h4></td>
	  </tr>
		<% count = 0%>
		<tr>
			<td>
				<table border=<%=graph_border%> cellpadding = 0 cellspacing = <%=graph_spacing%>><tr>
					<tr>
						<TD height="100%">
						</td>
						<td valign="bottom" align="right"><img src="" width="2" height="<%=graphmultiplier+8%>">						
						<%
						// '*******************MAIN PART OF THE CHART************************************
						for i = 0 to ubound(avalues)-1
							strgraph = agraph_bars(count)
							if alt_text = 1 then 
								stralt = alabels(i)
								elseif alt_text = 2 then 
								stralt = avalues(i)
								elseif alt_text = 3 then 
								stralt = alabels(i) &" - "  &avalues(i)
								elseif alt_text = 4 then   
								stralt = round(avalues(i) /intmax  *100,2) &"%"
							end if     
							if uselowvalue then  %>
							<td valign="bottom" align="center">
									<img src="<%=strgraph%>.gif" height="<%=round((avalues(i)-intmin)/intmax*graphmultiplier,0)%>" 
                                    width="<%=imgwidth%>" alt="<%=strAlt%>" border="<%=intimgborder%>"></td>
							<%else%>
								<td valign="bottom" align="center">
									<img src="<%=strgraph%>.gif" height="<%=round(avalues(i)/intmax*graphmultiplier,0)%>" 
                                    width="<%=imgwidth %>" alt="<%=strAlt%>" border="<%=intimgborder%>"></td>
							<%end if 
							if count = graph_bars-1 then 
								count = 0 
							else
								count = count + 1
							end if		
						next  
						//	'write out the border at the bottom of the bars also leave a blank cell for spacing on the right
						response.write "<td width='50'>&nbsp;</td></tr><tr><td width=8>&nbsp;</td><td>&nbsp;</td><td colspan=" &(ubound(avalues)) &" valign='top'>" _
				         &"<img src='imagenes/inferior.gif' width='100%' height='3'</td></tr>"
						if showlabels then %>
							<tr><td width=8 height=1>&nbsp;</td><td>&nbsp;</td>
								<div align="right">
								  <%for i = 0 to ubound(avalues)-1%>
							  </div>
								<td valign="bottom" width=<%=imgwidth%> ><div align="right"><strong><font size=
									<%=labelsize &">" &l_label_separator &alabels(i) &r_label_separator %></font></strong></div></td>
								<div align="right">
								  <%next%>
							      </div>
							</tr>
						<%end if%>
							<tr><td colspan=<%=ubound(avalues)+3%> height=50 align="center"><%=strxtitle%></td>
					</tr>
			</table>			</td>
		</tr>
		<tr>
		<td></td></tr>
	</table>
	<%end sub %>
<%
	Dim aMonthNames 'Pair title
	Dim aMonthValues 'Pair title
	Set Conn1 = Server.CreateObject("ADODB.Connection")
	Conn1.Open("provider=SQLOLEDB.1;Persist Security Info=False;User ID=inform; PASSword=inform;Initial Catalog=csc;Data Source=ANYJUAN-PC\srv_aplicc")
	set rs1 = CreateObject("ADODB.Recordset")
	puntos_Sql = "Select * from puntosdeventa where codpunto='4' and "
	puntos_Sql = puntos_Sql &" anio= "+cstr(anio1)+" and mes= "+cstr(mes1)
	rs1.Open puntos_Sql, Conn1
	reg=0
	do while not rs1.eof
		reg=reg+1
		rs1.movenext
	loop
	set rs1=nothing
	conn1.close
	set conn1=nothing

   	Redim aMonthNames(reg)
	Redim aMonthValues(reg)
	Set Conn2 = Server.CreateObject("ADODB.Connection")
	Conn2.Open("provider=SQLOLEDB.1;Persist Security Info=False;User ID=inform; PASSword=inform;Initial Catalog=csc;Data Source=ANYJUAN-PC\srv_aplicc")
	set rs2 = CreateObject("ADODB.Recordset")
	puntos1_Sql = "Select cuantos,nombre_punto,codvende,codpunto,nombre_vende from puntosdeventa where codpunto='4' and "
	puntos1_Sql = puntos1_Sql +" anio = "+cstr(anio1)+" and mes= "+cstr(mes1)
	rs2.Open puntos1_Sql, Conn2
	p=0
	do while not rs2.eof
		aMonthNames(p) = trim(rs2("nombre_vende").value)+" "+cstr(rs2("cuantos").value)
		aMonthValues(p) = (rs2("cuantos").value)
		p=p+1
		rs2.movenext
	loop
	set rs2=nothing
	conn2.close
	set conn2=nothing
	
	if p>0 then 
		titulo = "Clientes captados en el "+trim(session("puntoventa"))
		displayverticalgraph titulo,"","<b>Asesor Comercial</b>",aMonthValues,aMonthNames 
	else 
		response.redirect("submenu_graficos.asp?mensajeerror=" &"cclientet_nohay")		
	end if
	%>
	<table width="74%"  align="center">
		<tr>
	        <td width="19%"></td>
			<td width="81%">
				<FORM NAME="volver" ACTION="menu.asp">
					<input type="submit" name="asigna" value="Regresar">
				</FORM>
			</td>
			<td width="81%"></td>
		</tr>
</table>
</body>
</html>
