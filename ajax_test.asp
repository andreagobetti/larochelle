<!--#include virtual="/setup.asp" -->
<%
session("ajax_test")=time()
iduser=request.querystring("iduser")
if iduser<>"undefined" then
	where_iduser=" and iduser<>"&iduser
	session("ajax_test")=session("ajax_test")&"IDUSER:"&iduser
end if
if request("test")=1 then 'controlla fatture fornitore
	nfat=request.querystring("nfat")
	idfor=request.QueryString("idfor")
	session("ajax_test")=session("ajax_test")&"test:1 nfat:"&nfat&" idfor:"&idfor&"<br>"
	if nfat<>"" and idfor<>"" then
		sql="select idfor, nfat from fatture_for where idfor = "&request("idfor")&" and nfat='"&request("nfat")&"'"
		set rs=conn.execute(sql)
		if rs.eof then 
			session("ajax_test")=session("ajax_test")&"true"
			response.write "true"
		else
			session("ajax_test")=session("ajax_test")&"false"
			response.write "false"
		end if
	end if
	rs.close
elseif request("test")=2 then 'controlla email utente
	email=request.querystring("email")
	session("ajax_test")=session("ajax_test")&time()&"test:2 email:"&email&"<br>"
	if email<>"" then
		sql="select email,iduser,nome,cognome,azienda from utenti where email = '"&email&"'"&where_iduser
		set rs=conn.execute(sql)
		if rs.eof then 
			session("ajax_test")=session("ajax_test")&"true"
			response.write "true"
		else
			session("ajax_test")=session("ajax_test")&"false"
			response.write "ID utente: "&rs("iduser") &" "&nominativo(rs("cognome"),rs("nome"),rs("azienda"))
		end if
	end if
	rs.close
elseif request("test")=3 then 'controlla email utente
	piva=request.querystring("piva")
	session("ajax_test")=session("ajax_test")&"test:3 piva:"&piva&"<br>"
	if piva<>"" then
		sql="select piva,iduser from utenti where piva = '"&piva&"'"&where_iduser
		set rs=conn.execute(sql)
		if rs.eof then 
			session("ajax_test")=session("ajax_test")&"true"
			response.write "true"
		else
			session("ajax_test")=session("ajax_test")&"false"
			response.write "false"
		end if
	end if
	rs.close
elseif request("test")=4 then 'controlla email utente
	cf=request.querystring("cf")
	session("ajax_test")=session("ajax_test")&"test:4 cf:"&cf&"<br>"
	if cf<>"" then
		sql="select piva,iduser from utenti where cf = '"&cf&"'"&where_iduser
		set rs=conn.execute(sql)
		if rs.eof then 
			session("ajax_test")=session("ajax_test")&"true"
			response.write "true"
		else
			session("ajax_test")=session("ajax_test")&"false"
			response.write "false"
		end if
	end if
	rs.close

end if
%>