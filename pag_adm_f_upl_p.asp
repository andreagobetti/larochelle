<!--#include virtual="/setup.asp" -->
<%
id=Request("id")
Response.Expires=0
Response.Buffer = TRUE
Response.Clear
byteCount = Request.TotalBytes
RequestBin = Request.BinaryRead(byteCount)
Dim UploadRequest
Set UploadRequest = CreateObject("Scripting.Dictionary")
BuildUploadRequest  RequestBin
contentType = UploadRequest.Item("blob").Item("ContentType")
filepathname = UploadRequest.Item("blob").Item("FileName")
filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
if filename="" then response.redirect "pag_adm_artic.asp"
tmp=split(filename,".")
estensione=lcase(tmp(ubound(tmp)))

  value = UploadRequest.Item("blob").Item("Value")

  'Create filesytemObject Component
  Set ScriptObject = Server.CreateObject("Scripting.filesystemObject")

  'Create and Write to a File
  'response.write Request.ServerVariables("APPL_PHYSICAL_PATH")& mid(upl_img_cat,2)& id& "_b." & estensione
 Set MyFile = ScriptObject.CreateTextFile(Request.ServerVariables("APPL_PHYSICAL_PATH")& mid(upl_img_cat,2)& id& "_b." & estensione)

  For i = 1 to LenB(value)
	 MyFile.Write chr(AscB(MidB(value,i,1)))
  Next
  MyFile.Close
  
sql="select * FROM prodotti where idpro= " & id
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn, 3, 3

'response.write upl_img_cat&nomefile&"."&estensione
rs("fileimg")=id & "_b." & estensione
if not image_util then
	if instr("GIF JPG JPEG BMP OTB PNG WBMP",estensione)>0 then
		Set objImageSize = New ImageSize
		With objImageSize
		  .ImageFile = Server.MapPath(upl_img_cat&nomefile&"."&estensione)
		  If .IsImage Then
			x=.ImageHeight
			y=.ImageWidth
			rs("dimimg")=  x&"x"&y
		  Else
		  
		  End If 
		End With
		Set objImageSize = Nothing
	end if
end if
rs.update

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
    <tr align="center" valign="middle"> 
      <td height="30"> <%=titolo_top%> </td>
    </tr>
    <tr> 
      <td align="center" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0" class="tabella1">
          <tr>
            <td class="ui-widget-header">&nbsp;File ricevuto con successo</td>
          </tr>
          <tr>
            <td align="center">
			<br>
	<!-- INIZIO BLOCCO CENTRALE -->

<p align="center"><font face="Verdana" size="2"> File "<b><%=filename%></b>" ricevuto 
  con successo<br><br>
<%
if image_util then
	Response.Write("Ridimensionamento effettuato con Persits.Jpeg<br>")

	set Image = Server.CreateObject("Persits.Jpeg")
	Image.PreserveAspectRatio  = true
	Image.Quality =90
	Image.Open server.MapPath(upl_img_cat)&"/"&id&"_b."&estensione
	Response.Write("Dimensioni originali immagine:"&Image.OriginalWidth&"x"&Image.OriginalHeight&"<br>")
	rs("dimimg")=Image.Width&"x"&Image.Height
	rs.update
	
	'ridimensiona grande
	Response.Write("<br><b>Immagine grande</b><br>")
	dim_img=replace(Application("dim_img_b"),":","x")
	dim_img=split(dim_img,"x")
	Width=cint(dim_img(0))
	Height=cint(dim_img(1))
	Response.Write("Dimensioni massime immagine:"&dim_img(0)&"x"&dim_img(1)&"<br>")

	if image.OriginalWidth>Width or image.OriginalHeight>Height then
		Response.Write("Immagine ridimensionata<br>")
		altezza_prevista=image.OriginalHeight/(image.OriginalWidth/Width)
		Response.Write("altezza prevista "&altezza_prevista&"<br>")

		if altezza_prevista>Height then
			Response.Write("Ridimensiono per altezza<br>")

			Image.Height= Height
			
		else
			Response.Write("Ridimensiono per larghezza<br>")

			Image.Width= Width
		end if

	else
		Response.Write("Immagine non ridimensionata<br>")

	end if
	Image.Save (server.MapPath(upl_img_cat)&"/"&id&"_b."&estensione )
	Response.Write("Dimensioni finali immagine:"&Image.Width&"x"&Image.Height&"<br>")
	
	'ridimensiona piccola
	Response.Write("<br><b>Immagine piccola</b><br>")
	dim_img=replace(Application("dim_img_s"),":","x")
	dim_img=split(dim_img,"x")
	Width=cint(dim_img(0))
	Height=cint(dim_img(1))
	Response.Write("Dimensioni massime immagine:"&dim_img(0)&"x"&dim_img(1)&"<br>")

	if image.OriginalWidth>Width or image.OriginalHeight>Height then
		Response.Write("Immagine ridimensionata<br>")
		altezza_prevista=image.OriginalHeight/(image.OriginalWidth/Width)
		Response.Write("altezza prevista "&altezza_prevista&"<br>")

		if altezza_prevista>Height then
			Response.Write("Ridimensiono per altezza<br>")

			Image.Height= Height
			
		else
			Response.Write("Ridimensiono per larghezza<br>")

			Image.Width= Width
		end if

	else
		Response.Write("Immagine non ridimensionata<br>")

	end if
	
	Image.Save ( server.MapPath(upl_img_cat)&"/"&id&"_s."&estensione)
	Response.Write("Dimensioni finali immagine:"&Image.Width&"x"&Image.Height&"<br>")
	'if Image.LastError <> "" then
	'	Response.Write("Image Error: " + Image.LastError)
	'else
	'	Response.Write("Nessun errore")
	'end if
	Image.Close()
end if
%></font><br>
    <a href="product.asp?idpro=<%=id%>">Torna all'articolo</a></font></p>

<p align="center">Immagine grande<br><img src="<%= upl_img_cat & id&"_b"&".jpg"%>"><br>
<%if image_util then%>
Immagine piccola<br><img src="<%= upl_img_cat & id&"_s"&".jpg"%>"><%end if%></p>

	<!-- FINE BLOCCO CENTRALE -->
	    </td>
          </tr>
        </table> 
      </td>
    </tr>
    <tr> 
      <td height="1">&nbsp;</td>
    </tr>
  </table>
</center>
</body>
</html>
<%
rsClose
%>
<%

Sub BuildUploadRequest(RequestBin)	
	PosBeg = 1
	PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
	boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
	boundaryPos = InstrB(1,RequestBin,boundary)
	Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
		Dim UploadControl
		Set UploadControl = CreateObject("Scripting.Dictionary")
		Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
		Pos = InstrB(Pos,RequestBin,getByteString("name="))
		PosBeg = Pos+6
		PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
		Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
		PosBound = InstrB(PosEnd,RequestBin,boundary)
		If  PosFile<>0 AND (PosFile<PosBound) Then
			PosBeg = PosFile + 10
			PosEnd =  InstrB(PosBeg,RequestBin,getByteString(chr(34)))
			FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			UploadControl.Add "FileName", FileName
			Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
			PosBeg = Pos+14
			PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
			ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
			UploadControl.Add "ContentType",ContentType
			PosBeg = PosEnd+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
			Else
			Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
			PosBeg = Pos+4
			PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
			Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
		End If
		UploadControl.Add "Value" , Value	
		UploadRequest.Add name, UploadControl	
		BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
	Loop
  End Sub
  Function getByteString(StringStr)
	For i = 1 to Len(StringStr)
		char = Mid(StringStr,i,1)
		getByteString = getByteString & chrB(AscB(char))
	Next
  End Function
  Function getString(StringBin)
	getString =""
	For intCount = 1 to LenB(StringBin)
		getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
	Next
  End Function
%>