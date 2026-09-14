Buongiorno,
ho bisogno di leggere un file csv con campi numerici nel formato corretto<br>
Sembra che l'unico tipo di csv supportato sia quello separato da virgola ma i decimali non vengono interpretati correttamente. Vedi esempio sotto<br>


<%
		Dim FileObject
		Set FileObject=Server.CreateObject("Scripting.FileSystemObject")

		'FILE A
		Set InStream=FileObject.OpenTextFile(Server.MapPath("/public/a.csv"),1,False,False)
		linee=Instream.ReadAll
		InStream.Close
		Set InStream=Nothing
		response.write "<b>Contenuto del file a.csv separato da virgola</b><br><hr>"
		'MOSTRO A VIDEO IL CONTENUTO DEL FILE
		Response.write "<pre>" & linee & "</pre>"
		response.write "<hr>"
		Set csv = Server.CreateObject("ADODB.Connection")
		csv.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
       "Data Source=" & Server.MapPath("/public/") & ";" & _ 
       "Extended Properties=""text;HDR=Yes;FMT=Delimited"""

		Set RS_csv = csv.Execute("SELECT * FROM a.csv")
		response.write "Lettura del file con ADO(<b>I decimali non vengono letti correttamente</b>)<table border=1 >"
		do while not rs_csv.eof
		response.write "<tr><td>"&rs_csv("id")&"</td><td>"&rs_csv("nome")&"</td><td>"&rs_csv("prezzo")&"</td></tr>"
		
		
		rs_csv.movenext
		loop
		response.write "</table><br><br>"
		csv.close
		%>
		Questo invece è un file csv separato da punto e virgola ma i campi non vengono riconosciuti.<br>
		
		<%
		
		'FILE B
		Set InStream=FileObject.OpenTextFile(Server.MapPath("/public/b.csv"),1,False,False)
		linee=Instream.ReadAll
		InStream.Close
		Set InStream=Nothing
		response.write "<b>Contenuto del file b.csv separato da punto e virgola</b><br><hr>"
		'MOSTRO A VIDEO IL CONTENUTO DEL FILE
		Response.write "<pre>" & linee & "</pre>"
		response.write "<hr>"
		Set csv = Server.CreateObject("ADODB.Connection")
		csv.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
       "Data Source=" & Server.MapPath("/public/") & ";" & _ 
       "Extended Properties=""text;HDR=Yes;FMT=Delimited"""

		Set RS_csv = csv.Execute("SELECT * FROM b.csv")
		response.write "Lettura del file con ADO (<b>Va in errore perchè i campi non vengono riconosciuti</b>)<table border=1>"
		do while not rs_csv.eof
		response.write "<tr><td>"&rs_csv("id")&"</td><td>"&rs_csv("nome")&"</td><td>"&rs_csv("prezzo")&"</td></tr>"
		
		
		rs_csv.movenext
		loop
		response.write "</table>"
		csv.close

		
		set csv = nothing
		Set FileObject=Nothing
%>