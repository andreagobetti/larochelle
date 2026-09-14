<!--#include virtual="/setup.asp" -->
<%

if session("idadmin") = "" then call login()
dim dict_admin, dic_avvisi
call carica_admin()
application("avvisi")="A1:1,4|B1:"
call carica_avvisi()
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
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Configurazione avvisi</a></div>
        <!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font> 
        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>"  style="margin-top:0px;">
		
        <%
if oper="view" then
%>
			<table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
	            <tr style="border-bottom:1px solid;"> 
	              <td    bgcolor="#E5E5E5" colspan="3">Registrazione e login</td>
	            </tr>
	            <tr style="border-bottom:1px solid;"> 
					<td align="center" >
					   <input type="radio" name="modifica" value="A1" onclick="this.form.submit()"> 
					</td>

					<td nowrap >A1: Registrazione utente</td>
					<td><%=get_admins(dict_admin.ITEM("A1"))%></td>
	            </tr>
            
	            <tr style="border-bottom:1px solid;"> 
					<td align="center" >
					   <input type="radio" name="modifica" value="A1" onclick="this.form.submit()"> 
					</td>

					<td nowrap >A2: Richiesta password</td>
					<td><%=get_admins(dict_admin.ITEM("A2"))%></td>
	            </tr>

          </table>
          <%
else
txt=request.form("modifica")
if request.form("id")<>"" then txt=request.form("id")
	if request.form("modifica")<>"" then
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
              <td align="right" valign="top">Avviso</td>
              <td> 
              <%=txt%>
               <input type="text" name="descrizione" value="<%=val%>" size="80">
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
	
	sub carica_admin()
		Set dict_admin=Server.CreateObject("Scripting.Dictionary")
		Set objRS = conn.Execute("select iduser, nominativo from admin where fine is null") ' 1 call here
		If Not objRS.EOF Then
			arrRS = objRS.GetRows() ' 2 calls here
		else
			response.write "admin non trovati"
		end if
		Set objRS = Nothing
		If IsArray(arrRS) Then
			For i = LBound(arrRS, 2) To UBound(arrRS, 2)
				dict_admin.Add arrRS(0, i),arrRS(1, i)
			Next
			Erase arrRS
		End If
		response.write dict_admin.count
	end sub
	sub carica_avvisi()
		Set dict_avvisi=Server.CreateObject("Scripting.Dictionary")
		arr_avvisi=split(application("avvisi"),"|")
		If IsArray(arr_avvisi) Then
			For i = LBound(arr_avvisi) To UBound(arr_avvisi)
				field=split(arr_avvisi(i),":")
				dict_admin.Add field(0),field(1)
			Next
			Erase arr_avvisi
			Erase field
		End If
	end sub
	function get_admins(str)
		tmp=""
		arr=split(str,",")
		if IsArray(arr) then
			for n = 0 to ubound(arr)
				call concatena_stringa(tmp,", ",dict_admin.item(clng(arr(n))))
			next
		else
			tmp=dict_admin.item(str)
		end if
		get_admins=tmp
	end function
	
	
	%>