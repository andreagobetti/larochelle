<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->
<%   
Response.ContentType = "text/html"
Dim Js

if  request("term")<>"" and request("plugin")="" then
	term=replace(request("term"),"'","''")
	call determina_sqlinjection(term)

	Set Js = jsArray()
	strArr=""
	sql="select elenco_comuni.Comune, elenco_province.Sigla_automobilistica, elenco_province.Codice_regione, elenco_comuni.cap FROM elenco_comuni INNER JOIN elenco_province ON elenco_comuni.Codice_Provincia = elenco_province.Codice_provincia where  comune like '%"&term&"%' order by comune;"
	set rs=conn.execute(sql)
	do while not rs.eof
		Set Js(Null) = jsObject()
		js(null)("value")=rs("comune")
		Js(null)("targa")=rs("sigla_automobilistica")
		Js(null)("cap")=rs("cap")
		Js(null)("regione")=rs("codice_regione")
		rs.movenext
	loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	js.Flush
	set js = Nothing
end if
if  request("term")<>"" and request("plugin")="select2" then
	term=replace(request("term"),"'","''")
	call determina_sqlinjection(term)
	Set Js = jsArray()
	strArr=""
	
	sql="SELECT elenco_comuni.Comune, elenco_province.Sigla_automobilistica, elenco_province.Codice_regione, elenco_comuni.cap, MATCH(comune) AGAINST('"&term&"*' IN BOOLEAN MODE) AS attinenza FROM elenco_comuni inner join elenco_province on elenco_comuni.codice_provincia=elenco_province.codice_provincia WHERE MATCH(comune) AGAINST('"&term&"*' IN BOOLEAN MODE) ORDER BY comune"
	
	
	'sql="select elenco_comuni.Comune, elenco_province.Sigla_automobilistica, elenco_province.Codice_regione, elenco_comuni.cap FROM elenco_comuni INNER JOIN elenco_province ON elenco_comuni.Codice_Provincia = elenco_province.Codice_provincia where  comune like '%"&term&"%' order by comune;"
	set rs=conn.execute(sql)
	do while not rs.eof
		Set Js(Null) = jsObject()
		js(null)("text")=rs("comune")
		Js(null)("id")=rs("sigla_automobilistica")
		Js(null)("cap")=rs("cap")
		Js(null)("regione")=rs("codice_regione")
		rs.movenext
	loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	js.Flush
	set js = Nothing
end if

%>