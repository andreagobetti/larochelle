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
		oper="view"
		
		
end select
if request.form("annulla")<>"" then oper="view"
if oper="close" then
	sql="UPDATE ticket SET ticket.data_chiusura = now(), ticket.ticket_chiuso=true WHERE (ticket.id_ticket="& request("id_ticket")&")"
	conn.execute(sql)
	call invio_mail ("close",request("id_ticket"))
	oper="view"
end if
if oper="add_new" or oper="add_reply" then
	'controlli
end if

if oper="add_new" or oper="add_reply" then
id_ticket=request("id_ticket")
	Set rs = Server.CreateObject("ADODB.Recordset")
	if oper="add_new" then
	'aggiorno tabella ticket
		sql="select * from ticket 	 "
		sql=sql & " where id_ticket=0"
		rs.Open sql, conn, 3, 3
		rs.addnew
		
		rs("oggetto_ticket")=firstup(trim(request.form("oggetto_ticket")))
		rs("data_apertura")=now()	
		rs("id_user")=session("iduser")
		rs.update
		rs.movefirst
		id_ticket=rs("id_ticket")
		'response.write "idticket:"&id_ticket
		rs.close
	end if
'aggiorno tabella ticket_messaggi

	sql="select * from ticket_messaggi	where id_messaggio_ticket=0 "
	rs.Open sql, conn, 3, 3
	rs.addnew
	rs("data_messaggio_ticket")=now()	
	rs("messaggio_ticket")=firstup(trim(request.form("messaggio_ticket")))
	rs("id_ticket")=id_ticket
	rs("id_user")=session("iduser")
	rs.update
	rs.movefirst
	id_messaggio_ticket=rs("id_messaggio_ticket")
	rs.close
	
	select case oper
	case "add_new"
		call invio_mail ("add_new",id_messaggio_ticket)
	case "add_reply"
		call invio_mail ("add_reply",id_messaggio_ticket)
	end select
	oper="view"
end if

sub invio_mail(motivo,id)
select case motivo
case "close"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select Ticket.*, utenti.User, utenti.email FROM Ticket INNER JOIN utenti ON Ticket.id_user = utenti.iduser WHERE (((Ticket.id_ticket)="& id &"));"
		rs.Open sql, conn, 3, 3
		Subject = "Chiusura ticket assistenza su " & nomesito
		HTML="Il ticket assistenza con oggetto <b>" & rs("oggetto_ticket") &"</b> inserito dall'utente <b>" & rs("user") &"</b> è stato chiuso."
		HTML=HTML&"<br><br>Puoi consultare i ticket di assistenza a <a href=""http://" & nomesito & "/pag_ticket.asp"">questo indirizzo</a>"
		email_utente=rs("email")
		rs.close
		set rs=nothing
case "add_new"
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql="select Ticket.Oggetto_ticket, Ticket_messaggi.*, utenti.User, utenti.email, Ticket_messaggi.id_messaggio_ticket "
		sql=sql & "FROM (Ticket_messaggi INNER JOIN Ticket ON Ticket_messaggi.id_ticket = Ticket.id_ticket) INNER JOIN utenti ON Ticket_messaggi.id_user = utenti.iduser "
		sql=sql & "WHERE (((Ticket_messaggi.id_messaggio_ticket)=" &id &")); "
		rs.Open sql, conn, 3, 3
		Subject = "Inserimento nuovo ticket assistenza su " & nomesito
		HTML="E' stato aggiunto un nuovo ticket assistenza con oggetto <b>" & rs("oggetto_ticket") &"</b> inserito dall'utente <b>" & rs("user") &"</b><br>" & rs("messaggio_ticket") 
		HTML=HTML&"<br><br>Puoi consultare i ticket di assistenza a <a href=""http://" & nomesito & "/pag_ticket.asp"">questo indirizzo</a>"
		email_utente=rs("email")
		rs.close
		set rs=nothing
case "add_reply"
		sql="select Ticket.Oggetto_ticket, Ticket_messaggi.*, utenti.User,utenti.email, Ticket_messaggi.id_messaggio_ticket "
		sql=sql & "FROM (Ticket_messaggi INNER JOIN Ticket ON Ticket_messaggi.id_ticket = Ticket.id_ticket) INNER JOIN utenti ON Ticket_messaggi.id_user = utenti.iduser "
		sql=sql & "WHERE (((Ticket_messaggi.id_messaggio_ticket)=" &id &")); "
		rs.Open sql, conn, 3, 3
		Subject = "Inserimento nuovo messaggio di risposta a ticket assistenza su " & nomesito
		HTML="E' stata aggiunta una nuova risposta al ticket assistenza con oggetto <b>" & rs("oggetto_ticket") &"</b> inserito dall'utente <b>" & rs("user") &"</b><br>" & rs("messaggio_ticket") 
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
		objMessage.From = nomesito  & "<" & Application("email_sito") & ">"
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
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

<!--#include virtual="/sub_head.asp" -->
</head>
<body>
<center>
  <table height="100%" border="0" cellpadding="<%=cellp%>" cellspacing="<%=cells%>" class="tabellacentrale"> 
    <tr align="center" valign="middle" >
      <td height="30" colspan="3"><!--#include virtual="/sub_top.asp" --></td>
    <tr>
      <td width="<%=col_1%>" valign="top"><!-- Colonna SINISTRA INIZIO-->
        <!--#include virtual="/sub_menu.asp" -->
        <!-- Colonna SINISTRA FINE--></td>
      <td width="<%=col_23%>" align="center" valign="top"><!-- Colonna CORPO INIZIO-->
        <font color="#FF0000"><b><%=error%></b></font>
        <%
		
		sql="select Ticket.id_ticket, Ticket.id_user, Ticket.Oggetto_ticket, Ticket.Ticket_chiuso, Ticket.data_apertura, Ticket.data_chiusura, Count(Ticket_messaggi.id_messaggio_ticket) AS risposte, Max(Ticket_messaggi.data_messaggio_ticket) AS ultima_risposta, Ticket.id_user "
sql=sql & " FROM Ticket INNER JOIN Ticket_messaggi ON Ticket.id_ticket = Ticket_messaggi.id_ticket "
sql=sql & " GROUP BY Ticket.id_ticket, Ticket.id_user, Ticket.Oggetto_ticket, Ticket.Ticket_chiuso, Ticket.data_apertura, Ticket.data_chiusura, Ticket.id_user "
sql=sql & " HAVING (((Ticket.id_user)="& session("iduser") &")) "
'sql="select * from ticket where id_user="& session("iduser") &" "
if session("idadmin")<>"" then
	'sql="select * from ticket "
end if
'Blocco visualizzazione ELENCO TICKET

if oper="view" or oper="view_single" then
%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr>
            <td colspan="5" class="ui-widget-header">Elenco ticket assistenza</td>
          </tr>
          <tr align="right">
            <td colspan="5" style="border-bottom:1px solid;"><a href="<%=questofile%>?oper=new"><strong>Nuovo ticket assistenza</strong></a></td>
          </tr>
          <tr>
            <td  align="center" nowrap bgcolor="#E5E5E5" style="border-bottom:1px solid;">Data apertura</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Oggetto</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Risposte</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Data ultima risposta</td>
            <td bgcolor="#E5E5E5" style="border-bottom:1px solid;">Data chiusura</td>
          </tr>
          <%

	sql= sql & "ORDER BY data_chiusura;"
	set rs=conn.execute(sql)
	do until rs.EOF
%>
          <tr>
            <td align="center" style="border-bottom:1px solid;"><%=rs("data_apertura")%></td>
            <td style="border-bottom:1px solid;"><a href="<%=questofile%>?oper=view_single&id_ticket=<%=rs("id_ticket")%>"><%=rs("oggetto_ticket")%></a></td>
            <td style="border-bottom:1px solid;"><%=rs("risposte")%></td>
            <td style="border-bottom:1px solid;"><%=rs("ultima_risposta")%></td>
            <td style="border-bottom:1px solid;"><%=rs("data_chiusura")%></td>
          </tr>
          <%
	rs.movenext
	loop
	%>
        </table>
        <%
		  end if
'Blocco visualizzazione TICKET SINGOLO
		  if oper="view_single" or oper="reply" then
			sql="select * from ticket where ((id_ticket=" & request("id_ticket") &") and (id_user="& session("iduser") &")) "
			if session("idadmin")<>"" then
				sql="select * from ticket where ((id_ticket=" & request("id_ticket") &") ) "
			
			end if

		  
		  set rs=conn.execute(sql)

%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
          <tr>
            <td class="ui-widget-header">Ticket assistenza</td>
          </tr>
          <tr align="left">
            <td style="border-bottom:1px solid;" bgcolor="#E5E5E5"><b><%=rs("oggetto_ticket")%></b></td>
          </tr>
          <%
			id_ticket=rs("id_ticket")
			ticket_chiuso=rs("ticket_chiuso")
			rs.close
		  sql="select Ticket_messaggi.*, utenti.User FROM Ticket_messaggi INNER JOIN utenti ON Ticket_messaggi.id_user = utenti.iduser  where Ticket_messaggi.id_ticket=" & id_ticket
		  set rs=conn.execute(sql)
		do until rs.EOF%>
          <tr align="left">
            <td style="border-bottom:1px solid;">Messaggio di: <b><%=rs("user")%></b> il <b><%=rs("data_messaggio_ticket")%></b><br>
              <br>
              <%=rs("messaggio_ticket")%></td>
          </tr>
          <%	
	rs.movenext
	loop
	if ticket_chiuso=false  then
	%>
          <tr align="left">
            <td align="center" bgcolor="#E5E5E5" style="border-bottom:1px solid;"><form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>" style="margin-top:0px;">
                <input type="hidden" name="id_ticket" value="<%=id_ticket%>">
                <input type="hidden" name="oper" value="close">
                <input type="submit" name="invio" value="Chiudi il ticket">
              </form></td>
          </tr>
          <%end if%>
        </table>
        <%
		  end if
		  
'Blocco immissione RISPOSTE
		  
if oper="new" or oper="view_single" and ticket_chiuso=false then
if request.form("id")<>"" then txt=request.form("id")
	sql=sql & " where id_attrezzatura =" & txt
	select case oper
	case "new"
		titolo= "Nuovo ticket assistenza"
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
  if (theForm.oggetto_ticket.value == "")  { 
    alert("Il campo oggetto è obbligatorio."); 
    theForm.oggetto_ticket.focus(); 
    return (false); 
  } 
  <%end if%>
    if (tinyMCE.get('messaggio_ticket').getContent() == "")  { 
    alert("Il campo messaggio è obbligatorio."); 
	 tinyMCE.execCommand('mceFocus', false, "messaggio_ticket");
    return (false); 
  } 

  return (true); 
} 

</script>

        <form name="form1" method="post" action="<%=Request.ServerVariables("Script_Name")%>"  onSubmit="return Validator(this)" style="margin-top:0px;">
          <input type="hidden" name="id_ticket" value="<%=id_ticket%>">
          <table width="100%" border="0" cellpadding="2" cellspacing="0" class="tabella1">
            <tr>
              <td colspan="3" class="ui-widget-header"><%response.write titolo%></td>
            </tr>
            <% if oper="new" then%>
            <tr>
              <td align="right" valign="top">Oggetto ticket</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Campo obbligatorio" width="16" height="16"></td>
              <td><%if error<>"" then
			  		testo=request.form("oggetto_ticket")
				elseif oper="edit" then
					testo=rs("oggetto_ticket")
			  	elseif oper="new" then
			  		testo=""
				end if%>
                <input type="text" name="oggetto_ticket" value="<%=testo%>" size="60" style="font-weight:bold; font-size:12pt; width:500px"></td>
            </tr>
            <%end if%>
            <tr>
              <td align="right" valign="top">Messaggio</td>
              <td valign="top"><img src="images/question_s.gif" style="cursor:hand" alt="Campo obbligatorio" width="16" height="16"></td>
              <td><%if error<>"" then
			  		testo=request.form("messaggio_ticket")
				elseif oper="edit" then
					testo=rs("messaggio_ticket")
			  	elseif oper="new" then
			  		testo=""
				end if%>
                
              <script language="javascript" type="text/javascript" src="/tiny_mce/tiny_mce.js"></script>
<script language="javascript" type="text/javascript">
	tinyMCE.init({
    file_browser_callback: 'openKCFinder',mode : "textareas", theme : "advanced",theme_advanced_toolbar_location : "top", editor_selector : "mceEditor", theme_advanced_buttons1 : "bold,italic,underline,justifyleft,justifycenter,justifyright,justifyfull,forecolor,backcolor, bullist,numlist,separator,outdent,indent,separator,undo,redo,separator,link,unlink",theme_advanced_buttons2 : "",theme_advanced_buttons3:"", language : "it" });
</script>
  
                <textarea name="messaggio_ticket"  class="mceEditor" style="width: 500 px" rows="15"  ><%=testo%></textarea>  
            </tr>
          </table>
          <input type="hidden" name="oper" value="add_<%=oper%>">
          <input type="submit" name="invio" value="<%=titolo%>">
          <input type="submit" name="annulla" value="Annulla" onClick="pulsanteannulla=true;" >
        </form>
        <%
end if%>
        <!-- Colonna CORPO FINE--></td>
    </tr>
    <tr>
      <td colspan="2" height="1">&nbsp;</td>
    </tr>
  </table>
</center>
</body>
</html>
