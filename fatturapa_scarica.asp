
<%
	
if request("xml")<>"" then
	
		' Recupero il file da scaricare
	Dim download, file
	file = "/public/fatture_pa/"&request("xml")
	
	' Creo l'oggetto ADODB.Stream
	Set download = Server.CreateObject("ADODB.Stream")
	
	' Apro la connessione e carico il file
	download.Type = 1
	download.Open
	download.LoadFromFile Server.MapPath(file)
	
	' Aggiungo le intestazioni del tipo di file
	Response.AddHeader "Content-Disposition", "attachment; filename=" & request("xml")
	Response.ContentType = "application/octet-stream"
	Response.BinaryWrite download.Read
	
	' Un po di pulizia...
	download.Close
	Set download = Nothing
	response.end	
else
	
	Response.write "niente"
end if

%>