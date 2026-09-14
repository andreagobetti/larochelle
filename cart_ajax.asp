<!--#include virtual="/setup.asp" -->
<!--#include virtual="/cart_small_inc.asp" -->
<%
idpro=request("idpro")
idvara=request("idvara")
idvarb=request("idvarb")
quantita=request("quantita")
codice=request("codice")
if session("iduser")<>"" then
	iduser=session("iduser")
elseif isnumeric(request("iduser")) then
	iduser=request("iduser")
elseif SessionIdGuest<>"" then
	idloged=SessionIdGuest
elseif isnumeric(request("idloged")) then
	idloged=request("idloged")
elseif request("iduser")="null" or request("idloged")="null" or idpro="" then
		add2log "Manca idloged, iduser, idpro"&vbcrlf&queryeform(),3
		response.end
end if

if iduser<>"" then
	where_carrello="iduser="&iduser
	campo="iduser"
	valore=session("iduser")
	testoDa="[utente="&iduser&"]"&session("nominativo")&"[/utente]"
elseif idloged<>"" then
	where_carrello="idloged="&idloged
	campo="idloged"
	valore=idloged
	testoDa="utente non loggato"
else
		add2log "Errore non intercettato"&vbcrlf&queryeform(),3
		response.end
end if

sqlString = "select carrello.* FROM carrello WHERE  idpro=" & idpro & " and "&where_carrello
if idvara<>"undefined" then
	sqlString=sqlString & " AND idvara=" & idvara 
else
	idvara=0
end if
if idvarb<>"undefined" then
	sqlString=sqlString & " AND idvarb=" & idvarb 
else
	idvarb=0
end if
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sqlString, conn, 1, 3
if not rs.eof then
	rs("quantita")=rs("quantita")+quantita
else
	rs.addnew
	rs("idpro")=idpro
	rs("data")=now()
	rs("quantita")=quantita
	rs("idvara")=idvara
	rs("idvarb")=idvarb
	rs(campo)=valore
end if
rs.update
idcar=rs("idcar")
RS.Close
set rs=nothing
session("cache_carrellino")=""
'session("cart_ajax")=sqlString
'Conn.Execute sqlString	
call carrellino()

add2log "[articolo="& idpro&"]"&codice&"[/articolo] aggiunto al carrello da "&testoDa&" idcar:"&idcar,1
conn.close
set conn=nothing
%>