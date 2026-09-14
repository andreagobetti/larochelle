<!--#include virtual="/setup.asp" -->

<%
Dim objFso, objFolder, objfiles, strFolder, strFile 
strPath = upl_img_cat
strPathOld=strPath&"old/"
Set objFso = Server.CreateObject("Scripting.filesystemObject") 
Set objFso2 = Server.CreateObject("Scripting.filesystemObject") 
Set objFolder = objFso.GetFolder(Server.MapPath(strPath)) 
Set objfiles = objFolder.files 

if not objFso.FolderExists(Server.MapPath(strPathOld)) then
	objFso.CreateFolder(Server.MapPath(strPathOld))
	response.write "cartella creata<br>"
else
	response.write "cartella esistente<br>"
end if

session("test")=""
n=0
For Each strFile in objfiles
	sql="select files.* from files where filename='"&strFile.name&"'"
	set esiste=conn.Execute(sql)
	if esiste.eof then
		session("test")=session("test")&strFile.name&" orfano<br>"
		response.write "sposto "&Server.MapPath(strPath&strFile.name)&" "&Server.MapPath(strPathOld&strFile.name)&"<br>"
		objFso2.moveFile Server.MapPath(strPath&strFile.name), Server.MapPath(strPathOld&strFile.name)
		'objFso2.DeleteFile(strPath&strFile.name)
		
	else
		session("test")=session("test")&strFile.name&" TROVATO<br>"
	end if
	if n=2000 then
		response.write session("test")
		Set objFso = Nothing 
		Set objFolder = Nothing 
		Set objfiles = Nothing 
		set esiste =Nothing
		response.end
	end if
Next 
set esiste =Nothing

Set objFso = Nothing 
Set objFolder = Nothing 
Set objfiles = Nothing 
%>
