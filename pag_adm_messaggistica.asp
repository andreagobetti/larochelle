<!--#include virtual="/setup.asp" -->
<%
if session("iduser") = "" then call login()

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
	case "view_single"
		oper="view_single"
	case "reply"
		oper="reply"
	case "add_new"
		oper="add_new"
	case "add_reply"
		oper="add_reply"
	case "close"
		oper="close"
	case else
		oper="list"
		
		
end select
if request.form("annulla")<>"" then oper="view"

if oper="add_new" or oper="add_reply" then
	'controlli
end if

if oper="add_new" or oper="add_reply" then
	iddiscussione=request("iddiscussione")
	Set rs = Server.CreateObject("ADODB.Recordset")
	if oper="add_new" then
		sql="select * from messaggi_discussioni "
		rs.Open sql, conn, 1, 3
		rs.addnew
		rs("oggetto")=firstup(trim(request.form("oggetto")))
		rs("data_apertura")=now()	
		rs("iduser")=session("iduser")
		destinatari = request.form("destinatari")
		rs("destinatari")=destinatari&","&sessioniduser
		rs("scadenza")=request.form("scadenza")
		rs("stato")=request.form("stato")
		rs.update
		iddiscussione=Get_last_id("messaggi_discussioni")
		rs.close
		'Creo destinatari
		conn.execute("insert into discussioni_destinatari (`iduser`, `iddiscussione`) VALUES ("& sessioniduser &","&iddiscussione&")")
		arr_str_elenco = split(destinatari,",")
		for i = 0 to ubound(arr_str_elenco)
			conn.execute("insert into discussioni_destinatari (`iduser`, `iddiscussione`) VALUES ("&arr_str_elenco(i)&","&iddiscussione&")")
		next
	else
		'recupero destinatari
		'destinatari=conn.execute("select GROUP_CONCAT(iduser SEPARATOR ',') from discussioni_destinatari  where iddiscussione="&iddiscussione)(0)
		conn.execute("update messaggi_discussioni set stato="&request.form("stato"))
		destinatari=conn.execute("select destinatari from messaggi_discussioni where iddiscussione="&iddiscussione)(0)
	end if
	
	'aggiorno tabella messaggi
	sql="select * from messaggi	"
	rs.Open sql, conn, 1, 3
	rs.addnew
	rs("data")=now()	
	rs("messaggio")=firstup(trim(request.form("messaggio")))
	rs("iddiscussione")=iddiscussione
	rs("iduser")=session("iduser")
	rs.update
	'rs.movefirst
	idmessaggio=Get_last_id("messaggi")
	rs.close

	'Creo elenco lettura
	arr_str_elenco = split(destinatari,",")
	
	sql="select * from messaggi_utenti"
	rs.Open sql, conn, 1, 3
	
	
	
	for i = 0 to ubound(arr_str_elenco)
		if cint(arr_str_elenco(i))<>cint(sessioniduser) then
			rs.addnew
			rs("idmessaggio")=idmessaggio
			rs("iduser")=arr_str_elenco(i)
			rs("iddiscussione")=iddiscussione
			rs.Update
		end if
	next
	rs.close
	sql="select * from messaggi_discussioni where iddiscussione="&iddiscussione
	rs.Open sql, conn, 1, 3
	rs("ultima_risposta")=now()	
	rs.update
	rs.close
	

	
	select case oper
	case "add_new"
		'call invio_mail ("add_new",id)
	case "add_reply"
		'call invio_mail ("add_reply",id)
	end select
	Application.Lock
	application("ultima_nota")=now()
	Application.Unlock
	oper="list"
end if

sub invio_mail(motivo,id)
select case motivo
case "close"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="SELECT Ticket.*, Utenti.User, Utenti.email FROM Ticket INNER JOIN Utenti ON Ticket.id_user = Utenti.iduser WHERE (((Ticket.iddiscussione)="& id &"));"
		rs.Open sql, conn, 3, 3
		Subject = "Chiusura ticket assistenza su " & nomesito
		HTML="Il ticket assistenza con oggetto <b>" & rs("oggetto") &"</b> inserito dall'utente <b>" & rs("user") &"</b> è stato chiuso."
		HTML=HTML&"<br><br>Puoi consultare i ticket di assistenza a <a href=""http://" & nomesito & "/pag_ticket.asp"">questo indirizzo</a>"
		email_utente=rs("email")
		rs.close
		set rs=nothing
case "add_new"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="SELECT Ticket.oggetto, messaggi.*, Utenti.User, utenti.email, messaggi.id "
		sql=sql & "FROM (messaggi INNER JOIN Ticket ON messaggi.iddiscussione = Ticket.iddiscussione) INNER JOIN Utenti ON messaggi.id_user = Utenti.iduser "
		sql=sql & "WHERE (((messaggi.id)=" &id &")); "
		rs.Open sql, conn, 3, 3
		Subject = "Inserimento nuovo ticket assistenza su " & nomesito
		HTML="E' stato aggiunto un nuovo ticket assistenza con oggetto <b>" & rs("oggetto") &"</b> inserito dall'utente <b>" & rs("user") &"</b><br>" & rs("messaggio") 
		HTML=HTML&"<br><br>Puoi consultare i ticket di assistenza a <a href=""http://" & nomesito & "/pag_ticket.asp"">questo indirizzo</a>"
		email_utente=rs("email")
		rs.close
		set rs=nothing
case "add_reply"
		sql="SELECT Ticket.oggetto, messaggi.*, Utenti.User,utenti.email, messaggi.id "
		sql=sql & "FROM (messaggi INNER JOIN Ticket ON messaggi.iddiscussione = Ticket.iddiscussione) INNER JOIN Utenti ON messaggi.id_user = Utenti.iduser "
		sql=sql & "WHERE (((messaggi.id)=" &id &")); "
		rs.Open sql, conn, 3, 3
		Subject = "Inserimento nuovo messaggio di risposta a ticket assistenza su " & nomesito
		HTML="E' stata aggiunta una nuova risposta al ticket assistenza con oggetto <b>" & rs("oggetto") &"</b> inserito dall'utente <b>" & rs("user") &"</b><br>" & rs("messaggio") 
		HTML=HTML&"<br><br>Puoi consultare i ticket di assistenza a <a href=""http://" & nomesito & "/pag_ticket.asp"">questo indirizzo</a>"
		email_utente=rs("email")
		rs.close
		set rs=nothing


end select

'--EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL----EMAIL--
		Const cdoSendUsingMethod        = "http://schemas.microsoft.com/cdo/configuration/sendusing"
		Const cdoSendUsingPort          = 2
		Const cdoSMTPServer             = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
		Const cdoSMTPServerPort         = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
		Const cdoSMTPConnectionTimeout  = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
		Const cdoSMTPAuthenticate       = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
		Const cdoBasic                  = 1
		Const cdoSendUserName           = "http://schemas.microsoft.com/cdo/configuration/sendusername"
		Const cdoSendPassword           = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
		Dim objConfig
		Dim objMessage
		Dim Campi
		Set objConfig = Server.CreateObject("CDO.Configuration")
		Set Campi = objConfig.Fields
		With Campi
		.Item(cdoSendUsingMethod)       = cdoSendUsingPort
		.Item(cdoSMTPServer)            = SMTPServer
		.Item(cdoSMTPServerPort)        = 25
		.Item(cdoSMTPConnectionTimeout) = 30
		.Update
		End With
		Set objMessage = Server.CreateObject("CDO.Message")
		Set objMessage.Configuration = objConfig
			objMessage.To =  email_sito
			objMessage.To =  "andicot@alice.it;"&email_sito& ";"&email_utente
			objMessage.To =  email_utente&";"&email_sito& ";"
		objMessage.From = nomesito  & "<" & Application(appURL&"email_sito") & ">"
		objMessage.Subject=Subject
		objMessage.HtmlBody = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN""><!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""><html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type""<title></title></head><body>"&HTML&"</body></html>"
		objMessage.Send
		Set objMessage = Nothing

end sub

ticket_chiuso=false
'response.write "oper="&oper&""

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=application(appURL&"brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>
<!--#include virtual="/sub_head.asp" -->
<link rel="stylesheet" type="text/css" href="Jquery/js/colpick/colpick.css"/>
<style>
/*notification-bubble  */
.notification-bubble {
	font-size:inherit;
	float:none;
	display:inline-block;
}
</style>
</head>
<body> 
<div id="wrap">
<div id="header">
<%=application(appURL&"top")%>
        <!-- Div Content INIZIO-->      <%barra=0%>
	  <!--#include virtual="/sub_barra_adminsf2.asp" -->
        <div class="ui-widget-header ui-corner-all titolo_admin"><a href="<%=questofile%>">Note</a></div>
        <!-- Colonna CORPO INIZIO-->      
        <font color="#FF0000"><b><%=error%></b></font>
        <%
		


'select * from (messaggi_discussioni inner join discussioni_destinatari) 
'select * from messaggi_discussioni inner join (select iddiscussione from discussioni_destinatari where iduser=4) as t on messaggi_discussioni.iddiscussione = t.iddiscussione inner join (select sum(idmessaggio) as risposte, iddiscussione from messaggi group by iddiscussione) as messaggi on messaggi_discussioni. iddiscussione = messaggi.iddiscussione

'select * from messaggi_discussioni inner join (select iddiscussione from discussioni_destinatari where iduser=4) as t on messaggi_discussioni.iddiscussione = t.iddiscussione inner join (select count(idmessaggio) as risposte, iddiscussione from messaggi group by iddiscussione) as messaggi on messaggi_discussioni.iddiscussione = messaggi.iddiscussione left join( select idmessaggio, iduser, count(id) from messaggi_utenti where data_lettura is null and iduser=4) as messaggi_utenti on messaggi.idmessaggio = messaggi_utenti.idmessaggio 

'select * from messaggi_discussioni inner join (select iddiscussione from discussioni_destinatari where iduser=4) as t on messaggi_discussioni.iddiscussione = t.iddiscussione inner join (select count(idmessaggio) as risposte, iddiscussione from messaggi group by iddiscussione) as messaggi on messaggi_discussioni.iddiscussione = messaggi.iddiscussione left join( select iddiscussione, iduser, count(id) from messaggi_utenti where data_lettura is null and iduser=4) as messaggi_utenti on messaggi_discussioni.iddiscussione = messaggi_utenti.iddiscussione
response.write sql
'///////////////////////////////Blocco visualizzazione ELENCO discussioni//////////////////////////////
if oper="list" then
	sql="SELECT DISTINCT messaggi_discussioni.iddiscussione, discussioni_destinatari.iduser as destinatario, messaggi_discussioni.data_apertura, messaggi_discussioni.ultima_risposta, messaggi_discussioni.Oggetto, Count(messaggi.idmessaggio) AS risposte FROM ((messaggi_discussioni inner join discussioni_destinatari on messaggi_discussioni.iduser = discussioni_destinatari.iduser) INNER JOIN messaggi ON messaggi_discussioni.iddiscussione = messaggi.iddiscussione) left JOIN messaggi_utenti ON messaggi.idmessaggio = messaggi_utenti.idmessaggio GROUP BY messaggi_discussioni.iddiscussione, messaggi_discussioni.data_apertura, messaggi_discussioni.ultima_risposta, messaggi_discussioni.Oggetto, messaggi_utenti.iduser HAVING (((discussioni_destinatari.iduser)="&sessioniduser&")) order by messaggi_discussioni.ultima_risposta desc"
	
	
	sql="select messaggi_discussioni.*, messaggi.risposte, daleggere.daleggere from messaggi_discussioni inner join (select iddiscussione from discussioni_destinatari where iduser="&sessioniduser&") as t on messaggi_discussioni.iddiscussione = t.iddiscussione inner join (select count(*) as risposte, iddiscussione from messaggi group by iddiscussione) as messaggi on messaggi_discussioni.iddiscussione = messaggi.iddiscussione left join( select iddiscussione, count(*) as daleggere from messaggi_utenti where data_lettura is null and iduser="&sessioniduser&" group by iddiscussione ) as daleggere on messaggi_discussioni.iddiscussione = daleggere.iddiscussione"
	
if utente_andrea then response.write "[soloio]:"&sql
%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class=tabella1>
          <tr align="right">
            <td colspan="6" style="border-bottom:1px solid;"><a href="<%=questofile%>?oper=new"><strong>Nuova</strong></a></td>
          </tr>
          <tr>
            <td  align="center" nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">Data apertura</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Oggetto</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Stato</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Risposte</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Data ultima risposta</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Scadenza</td>
          </tr>
          <%

	'sql= sql & "ORDER BY data_chiusura;"
	set rs=conn.execute(sql)
	do until rs.EOF
	
	if not isnull(rs("daleggere"))  then
		n_msg=" / <span class=""badge"">"&rs("daleggere")&"</span>"
	else
		n_msg=""

	end if
	
%>
          <tr style="border-bottom:1px solid;" class="stato-nota-<%=rs("stato")%>">
            <td align="center" ><%=rs("data_apertura")%></td>
            <td ><a href="<%=questofile%>?oper=view_single&iddiscussione=<%=rs("iddiscussione")%>"><%=rs("oggetto")%></a></td>
            <td align="center" style="border-bottom:1px solid;"><%=stato_nota(rs("stato"))%></td>
            <td ><%=rs("risposte")%><%=n_msg%></td>
            <td ><%=rs("ultima_risposta")%></td>
            <td ><%=rs("Scadenza")%></td>
          </tr>
          <%
	rs.movenext
	loop
	%>
        </table>
        <%
		  end if
'////////////////////////////////Blocco visualizzazione DISCUSSIONE/////////////////////////////////////////
		  if oper="view_single" or oper="reply" then
			sql="select * from messaggi_discussioni where iddiscussione=" & request("iddiscussione") &" "
			if session("idadmin")<>"" then
				'sql="select * from ticket where ((iddiscussione=" & request("iddiscussione") &") ) "
			
			end if

		  
		  set rs=conn.execute(sql)

%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class=tabella1>
          <tr>
            <td class=titoloriquadro>Nota</td>
          </tr>
          <tr align="left">
            <td style="border-bottom:1px solid;" bgcolor="#E5E5E5"><b><%=rs("oggetto")%></b></td>
          </tr>
          <%
			iddiscussione=rs("iddiscussione")
			stato=rs("stato")
			'ticket_chiuso=rs("ticket_chiuso")
			rs.close
		  sql="SELECT messaggi.*, admin.nominativo, messaggi_utenti.data_lettura FROM (messaggi INNER JOIN admin ON messaggi.iduser = admin.iduser) INNER JOIN messaggi_utenti ON messaggi.idmessaggio = messaggi_utenti.idmessaggio WHERE (((messaggi_utenti.iduser)="&session("iduser")&") AND ((messaggi.iddiscussione)=" & iddiscussione&"));"
		  set rs = conn.execute(sql)
		  check_msg=false
		  first=true
		do until rs.EOF
		if first then txt="messaggio" else txt="Risposta"
		if isnull(rs("data_lettura")) then
			idmessaggio=rs("idmessaggio")
			nuovo_msg=" <span class=""badge"">nuovo</span>"
			sql="update messaggi_utenti set data_lettura=now() where idmessaggio="&idmessaggio&" and iduser="&sessioniduser
			response.write sql
			conn.execute(sql)
			
			check_msg=true
		else
			nuovo_msg=""
		end if
		%>
          <tr align="left">
            <td style="border-bottom:1px solid;"><div><%=txt%> di: <b><%=rs("nominativo")%></b> il <b><%=rs("data")%></b><%=nuovo_msg%>  
            
            <%
	            sql_da_leggere="SELECT messaggi_utenti.*, admin.Nominativo FROM admin INNER JOIN messaggi_utenti ON admin.iduser = messaggi_utenti.iduser where  idmessaggio="&rs("idmessaggio")&" and data_lettura is null"
            set rs_letture=conn.execute (sql_da_leggere)
            txt_da_leggere=""
            do while not rs_letture.eof
            
					txt_da_leggere=txt_da_leggere&"<span class=""badge"">"&nomebreve(rs_letture("nominativo"),"")&"</span> "
            rs_letture.movenext
            loop
            rs_letture.close
            if  txt_da_leggere<>"" then 
	            response.write "Non ha ancora letto:"&txt_da_leggere
            end if
            %>
            
            </div>
              <%=rs("messaggio")%></td>
          </tr>
          <%	
	rs.movenext
	first=false
	loop%>
        </table>
        <%
	        if check_msg then
	        	set rs_msg=conn.execute ("SELECT Count(messaggi_utenti.id) AS ConteggioDiid FROM messaggi_utenti GROUP BY messaggi_utenti.data_lettura, messaggi_utenti.iduser HAVING (((messaggi_utenti.data_lettura) Is Null) AND ((messaggi_utenti.iduser)="&session("iduser")&"));")
	        	if not rs_msg.eof then
		        	session("n_msg")=rs_msg("ConteggioDiid")
		        else
		        	session("n_msg")=""
		        end if
	        	session("data_chk")=now()
	        end if

        
		  end if
		  
'Blocco immissione RISPOSTE
		  
if oper="new" or oper="view_single"  then
if request.form("id")<>"" then txt=request.form("id")
	sql=sql & " where id_attrezzatura =" & txt
	select case oper
	case "new"
		titolo= "Nuova nota"
	case "view_single"
		titolo= "Aggiungi messaggio di risposta"
		oper="reply"
	end select	
	%>
        <script Language="JavaScript"> 
pulsanteannulla=false;
function Validator(theForm) 
{
  if (pulsanteannulla == true)  { 
  	return(true);    
  } 
<% if oper="new" then%>
  if (theForm.oggetto.value == "")  { 
    alert("Il campo oggetto è obbligatorio."); 
    theForm.oggetto.focus(); 
    return (false); 
  } 
  <%end if%>

  return (true); 
} 

</script>

        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>"  onSubmit="return Validator(this)" style="margin-top:0px;">
          <input type="hidden" name="iddiscussione" value="<%=iddiscussione%>">
          
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class=tabella1>
            <tr>
              <td  class=titoloriquadro><%response.write titolo%></td>
            </tr>
            <% if oper="new" then%>
            <tr>
              <td>Destinatari
                  <select name="destinatari" id="destinatari" class="select2" style="width:450px;" multiple>
              <%
				Set rsm=conn.execute("select * from admin where iduser>1 and fine is null and iduser<>"&sessionIDUser&" order by nominativo")
				do while not rsm.eof
					response.write "<option value=""" & rsm("iduser") &""""

					response.write ">" & nomebreve(rsm("nominativo"),"")&"</option>"
					rsm.movenext
				loop
				rsm.close
				%>
            </select>
            
</td>
            </tr>
            
            <tr>
              <td><%if error<>"" then
			  		testo=request.form("oggetto")
				elseif oper="edit" then
					testo=rs("oggetto")
			  	elseif oper="new" then
			  		testo=""
				end if%>Oggetto
                <input type="text" name="oggetto" value="<%=testo%>" size="60" style="font-weight:bold; font-size:12pt; width:500px"></td>
            </tr>
            <%end if%>
            <tr>
              <td ><%if error<>"" then
			  		testo=request.form("messaggio")
				elseif oper="edit" then
					testo=rs("messaggio")
			  	elseif oper="new" then
			  		testo=""
				end if%>
                
			<!--#include virtual="/tinymce4_conf2.asp" -->

  
                <textarea name="messaggio"  class="mceEditor" style="width: 500 px" rows="15"  ><%=testo%></textarea>  
            </tr>
            <%
            if oper="new" then
	            if error<>"" then
			  		val=request.form("oggetto")
				elseif oper="edit" then
					val=rs("scadenza")
			  	elseif oper="new" then
			  		val=""
				end if%>
            <tr>
	            <td>
		            Scadenza: <input type="text" name="sadenza" id="scadenza" value="<%=val%>">
	            </td>
            </tr>
            <%
	        end if    
				if oper="new" then
			  		val=1
			  	else
					val=stato
				end if
			%>
            <tr>
	            <td>
		            Stato: <select name="stato">
			            	<option value="1" <%if val=1 then response.write "selected"%>>Da fare</option>
			            	<option value="5" <%if val=5 then response.write "selected"%>>In corso</option>
			            	<option value="9" <%if val=9 then response.write "selected"%>>Completato</option>
		            </select>
		            
	            </td>
            </tr>

            <tr>
              <td  > <input type="hidden" name="oper" value="add_<%=oper%>">
          <input type="submit" name="invio" value="<%=titolo%>">
          <input type="submit" name="annulla" value="Annulla" onClick="pulsanteannulla=true;" ></td>
            </tr>
            
          </table>
         
        </form>
    	<script src="jquery/ui/i18n/jquery.ui.datepicker-it.min.js"></script>

        <script>
	        $(function() {
	        		$( "#scadenza" ).datepicker();
	        });
	        
	        
	    </script>    
        <%
end if%>
        <!-- Colonna CORPO FINE-->
        <!-- Div Content FINE-->
</div>
<div id="footer"></div></div>
</body>
</html>
<script type="text/javascript" src="Jquery/js/select2.js"></script> 
<script>
		$(function() {
		$(".select2").select2();
		});
	</script>

<%
rsClose
function stato_nota(stato)
	select case stato
		case 1
			stato_nota="Da fare"
		case 5
			stato_nota="In corso"
		case 9
			stato_nota="Completato"
	end select
end function
function nomebreve(byval nome,byval cognome)
	dim da_splittare
	if cognome="" then
        anominativo=split(nome," ")
        cognome=anominativo(1)
        nome=anominativo(0)
	end if
    nomebreve=nome&" "&left(cognome,1)&"."

end function
%>
