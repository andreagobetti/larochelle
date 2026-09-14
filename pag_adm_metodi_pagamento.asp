<!--#include virtual="/setup.asp" -->
<%

if session("idadmin") = "" then call login()

%>
<%
'------------------------------------

oper=request.form("oper")

select case lcase(request("oper"))
	case "annulla"
		oper="view"
	case "new"
		oper="new"
	case "aggiungi"
		oper="add"
	case "modifica"
		oper="update"
	case else
		oper="view"
		if request.form("modifica")<>"" then oper="edit"
end select
if oper="add" or oper="update" then
	'controlli
end if

if oper="add" or oper="update" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from metodi_pagamento"
	if oper="add" then
		rs.Open sql, conn, 3, 3
		rs.addnew
	else
		sql= sql & " where id=" & request.form("id")
		rs.Open sql, conn, 3, 3
	end if
	rs("descrizione")=ucase(trim(request.form("descrizione")))
	rs("richiedi_banca")=checkbox("richiedi_banca")
	rs("numero_scadenze")=Request.Form("numero_scadenze")
	rs("inizio_scadenze")=Request.Form("inizio_scadenze")
	rs("riba")=checkbox("riba")
	rs("descrizione_checkout")=trim(request.form("descrizione_checkout"))
	rs("pagamentoPA")=trim(request.form("pagamentoPA"))
	
	
	rs("maggiorazione_gestione")=aggiusta_decimale(trim(request.form("maggiorazione_gestione")),"asp")
	rs.update
	rs.close
	set rs=Nothing
	oper="view"
end if
'response.write oper
%>
<!--#include virtual="/sub_head_adm.asp" -->
</head>
<body> 
<div id="wrap">
<div id="header">
<%=titolo_top%>
        <!-- Div Content INIZIO-->      <%barra=0%>
	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Metodi di pagamento</a></div>
        <!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font> 
        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>"  style="margin-top:0px;">
		
        <%
sql="select * from metodi_pagamento order by descrizione "
if oper="view" then
%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr align="right"> 
              <td colspan="3" style="border-bottom:1px solid;"><span style="float:right;">
                <input type="button"  value="Nuovo" style="FONT: 12px;" onclick="window.location.href='<%=questofile%>?oper=new'"></span></td>
            </tr>
            <tr> 
              <%if oper="view" then%>
              <td width="10" align="center" nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">&nbsp;Modifica&nbsp;</td>
              <%end if%>
              <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Tipologia</td>
              <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Scadenze</td>
            </tr>
            <%
	set rs=conn.execute(sql)
	do until rs.EOF
%>
            <tr style="border-bottom:1px solid;"> 
              <td align="center" > <input type="radio" name="modifica" value="<%=rs("id")%>" onclick="this.form.submit()"> 
              </td>

              <td nowrap ><%=rs("descrizione")%></td>
              <td nowrap ><%=rs("numero_scadenze")%></td>
            </tr>
            <%
	rs.movenext
	loop
	set rs = Nothing
	%>
          </table>
          <%
else
txt=request.form("modifica")
if request.form("id")<>"" then txt=request.form("id")
	sql="select * from metodi_pagamento where id =" & txt
	if request.form("modifica")<>"" then
		set rs=conn.execute(sql)
		testo= "Modifica metodo pagamento"
		
	else
		testo= "Nuovo metodo pagamento"
	end if
		  %>
          <input type="hidden" name="id" value="<%=txt%>">
          <%if error<>"" then
			  		descrizione=request.form("descrizione")
			  	elseif oper="new" then
			  		descrizione=""
				end if%>
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr> 
              <td colspan="2" class="ui-widget-header"><%response.write testo%></td>
            </tr>
            <tr> 
              <td align="right" valign="top">Descrizione</td>
              <td>  <%if error<>"" then
			  		val=request.form("descrizione")
				elseif oper="edit" then
					val=rs("descrizione")
			  	elseif oper="new" then
			  		val=""
				end if%> <input type="text" name="descrizione" value="<%=val%>" size="80">
              </td>
            </tr>
            <tr>
              <td valign="top">Richiedi dati bancari</td>
              <td>
	            <%
	            val=""
				if oper="edit" then
					if rs("richiedi_banca") then val=" checked"
				elseif oper="new" then
					val=""
				else
					val=request.form("richiedi_banca")
					if val="1" then val=" checked"
				end if
				if val<>" checked" then val=""
				%>
              <input type="checkbox" name="richiedi_banca" value="1"  <%response.write val%> ></td>
            </tr>
            
            
            <tr> 
              <td align="right" valign="top">Numero scadenze</td>
              <td>  <%if error<>"" then
			  		val=request.form("numero_scadenze")
				elseif oper="edit" then
					val=rs("numero_scadenze")
			  	elseif oper="new" then
			  		val="0"
				end if%> <input type="text" name="numero_scadenze" value="<%=val%>" size="2"> 0= non genera scadenze
              </td>
            </tr>
            <tr> 
              <td align="right" valign="top">Inizio prima scadenza</td>
              <td>  <%if error<>"" then
			  		val=request.form("inizio_scadenze")
				elseif oper="edit" then
					val=rs("inizio_scadenze")
			  	elseif oper="new" then
			  		val=-1
				end if%> 
				<select name="inizio_scadenze">
				<option value="-1" <%if val=-1 then response.write " selected"%>>Data fattura</option>
				<option value="0" <%if val=0 then response.write " selected"%>>Fine mese data fattura</option>
				<option value="1" <%if val=1 then response.write " selected"%>>Fine mese +1 data fattura</option>
				<option value="2" <%if val=2 then response.write " selected"%>>Fine mese +2 data fattura</option>
				<option value="3" <%if val=3 then response.write " selected"%>>Fine mese +3 data fattura</option>
				<option value="4" <%if val=4 then response.write " selected"%>>Fine mese +4 data fattura</option>
				</select>
				
				
              </td>
            </tr>
            <tr>
              <td align="right"  valign="top">Riba</td>
              <td>
	            <%
	            val=""
				if oper="edit" then
					if rs("riba") then val=" checked"
				elseif oper="new" then
					val=""
				else
					val=request.form("riba")
					if val="1" then val=" checked"
				end if
				if val<>" checked" then val=""
				%>
              <input type="checkbox" name="riba" value="1"  <%response.write val%> ></td>
            </tr>
            <tr> 
              <td align="right" valign="top">Codice pagamento fattura elettronica</td>
              <td>  <%if error<>"" then
			  		val=request.form("pagamentoPA")
				elseif oper="edit" then
					val=rs("pagamentoPA")
			  	elseif oper="new" then
			  		val=""
				end if%> <input type="text" name="pagamentoPA" value="<%=val%>" size="4">
              </td>
            </tr>
            <tr> 
              <td align="right" valign="top">Descrizione per checkout</td>
              <td>  <%if error<>"" then
			  		val=request.form("descrizione_checkout")
				elseif oper="edit" then
					val=rs("descrizione_checkout")
			  	elseif oper="new" then
			  		val=""
				end if%> 		<textarea name="descrizione_checkout" class=""  style="width:100%; height:200px;"><%=val%></textarea>
              </td>
            </tr>
            <tr> 
              <td align="right" valign="top">Percentuale maggiorazione</td>
              <td>  <%if error<>"" then
			  		val=request.form("maggiorazione_gestione")
				elseif oper="edit" then
					val=formatnumber(rs("maggiorazione_gestione"),2)
			  	elseif oper="new" then
			  		val=""
				end if%> <input type="text" name="maggiorazione_gestione" value="<%=val%>" size="10">
              </td>
            </tr>
            
          </table>
        <%
	        Set rs = Nothing
if oper="view" then
	txt=puls_new
elseif oper="new" or error<>"" then
	txt="Aggiungi"
elseif oper="edit" then
	txt="Modifica"
end if
%>
    <input type="submit" name="oper" value="<%=txt%>">
    <%if oper<>"view" then%>
    <input type="submit" name="oper" value="Annulla" onclick="annulla=true;">
    <%end if%>
    <%if oper="edit" and num=0 and error="" and false then%>
    <input type="submit" name="elimina" value="Elimina">
    
    
    <%end if
	%>
    
<%end if
	call connclose()
%>
        </form> 

        <!-- Div Content FINE-->
</div>
<div id="footer"></div>
<!--#include virtual="/pag_adm_footer_inc.asp" --></div>
</body>
</html>
<%
rsClose
%>
