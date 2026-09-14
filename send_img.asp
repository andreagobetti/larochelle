<!--#include virtual="/config/mdb.asp" -->
<%
id=request.querystring("id")
Response.Expires = 0
Width=400 'cint(dim_img(0))
Height=400 'cint(dim_img(1))
filigrana=true

set Image = Server.CreateObject("Persits.Jpeg")

Image.PreserveAspectRatio  = true
Image.Quality =90
if request.querystring("path")<>"" then
	immaginePath=server.MapPath(request.querystring("path"))
elseif request.querystring("imgprod")<>"" then
	immaginePath=server.MapPath(upl_img_cat)&"/"&request.querystring("imgprod")
elseif request.querystring("imgprod_m")<>"" then
	immaginePath=server.MapPath(upl_img_cat)&"/"&request.querystring("imgprod_m")
	Width=228 
	Height=319 
	filigrana=false
elseif request.querystring("imgs1")<>"" then	'Immagini carrello piccolo
	immaginePath=server.MapPath(upl_img_cat)&"/"&request.querystring("imgs1")
	Width=85 
	Height=119
	filigrana=false
elseif request.querystring("imgs2")<>"" then	'Immagini product.asp prodotto
	immaginePath=server.MapPath(upl_img_cat)&"/"&request.querystring("imgs1")
	Width=228 
	Height=319 
	filigrana=false
elseif request.querystring("imgs3")<>"" then	'Immagini pag_adm_artic.asp elenco
	immaginePath=server.MapPath(upl_img_cat)&"/"&request.querystring("imgs3")
	Width=150 
	Height=150 
	filigrana=false
else
	immaginePath=server.MapPath(upl_img_cat)&"/"&id&"_b.jpg"
end if
	'session("test")="send_img.asp:"&immaginePath
	'response.write server.MapPath(upl_img_cat)&"/"&id&"_b.jpg"
	on error resume Next
	Image.Open immaginePath
	'dim_img=replace(Application("dim_img_b"),":","x")
	if err<>0 then
		
		immaginePath=server.MapPath("/public")&"/no_img.jpg"
		session("immaginePath")=immaginePath
		Image.Open immaginePath
	end if
	'dim_img=split(dim_img,"x")
	s=request.querystring("s")
	if s<>"" then
		Width=cint(s)
		Height=cint(s)
		filigrana=false
	elseif request.QueryString("w")<>"" then
		'ridimensiona per larghezza imposta
		width=cint(request.QueryString("w"))
		Height=image.OriginalHeight
	elseif request.QueryString("op")="b" then
		'nessun ridimensionamento
		Width=image.OriginalWidth
		Height=image.OriginalHeight
		filigrana=true
	elseif request.QueryString("max")<>"" then
		' ridimensionamento per larghezza e altezza max
		dim_img=split(request.QueryString("max"),"x")
		Width=cint(dim_img(0))
		Height=cint(dim_img(1))
		filigrana=false
		'session("send_img")=session("send_img")&"Ridimensiono per max "&cint(dim_img(0))&"x"&cint(dim_img(1))
	

	else
	
	end if
	'calcolo fattori di riduzione
	'session("send_img")=session("send_img")&"Dimensione originale: w"&image.OriginalWidth&" h"&image.OriginalHeight
	fattore_riduzione_w=image.OriginalWidth/Width
	fattore_riduzione_h=image.OriginalHeight/Height
	'session("send_img")=session("send_img")&" Fattori di riduzione: w"&fattore_riduzione_w&" h"&fattore_riduzione_h
	if fattore_riduzione_w<=1 and fattore_riduzione_h<=1 then
		'nessuna riduzione
		'session("send_img")=session("send_img")&" Nessuna riduzione"
	elseif fattore_riduzione_w>=fattore_riduzione_h then
		Image.Width= Width
		'Image.Height= image.OriginalHeight/fattore_riduzione_w
		'session("send_img")=session("send_img")&" Regolato larghezza w="&Image.Width&" h="&Image.Height
	else
		Image.Height= Height
		'Image.Width=image.OriginalWidth/fattore_riduzione_h
		'session("send_img")=session("send_img")&" Regolato altezza w="&Image.Width&" h="&Image.Height
	end if
	if width<200 then
		Image.Interpolation=0

	else
		Image.Interpolation=1
	end if
	'session("send_img")=session("send_img")&" Dimensione finale "&Image.Width&"x"&Image.Height
'session("send_img")=""
Image.SendBinary
set Image = nothing

%> 
