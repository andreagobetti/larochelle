<%
	' dbpath=server.MapPath("\mdb-database\server_start.mdb")
' 	mdb="Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & dbpath
' 	Set conn = Server.CreateObject("ADODB.Connection")
' 	conn.Open mdb
' 	Set rs_log = Server.CreateObject("ADODB.Recordset")
' 	rs_log.Open "log", conn, 1, 3
' 	rs_log.addnew
' 	rs_log("evento")="Application_OnStart"
' 	rs_log("causale")=1
' 	rs_log.update
' 	rs_log.close
' 	set rs_log=nothing
' 	conn.Close
' 	set conn=nothing
	

	filePath = Server.Mappath("/public/Application_OnStart.txt")
	Set objfilesystem = Server.CreateObject("Scripting.filesystemObject")
	if not objfilesystem.FileExists(filePath) then
		'se non esiste lo creo
		objfilesystem.CreateTextFile(filePath)
	end if
	Set objFile = objfilesystem.OpenTextFile(filePath, 8)
	objFile.WriteLine(now()&" Application_OnStart")
	objFile.Close
	Set objFile=Nothing
	Set objfilesystem=Nothing
%>

