<!--#include virtual="/setup.asp" -->
<%
		'Sostituisco nel file di testo ' - ' con ' '
		'if false then
		'	Const ForReading = 1
		'	Const ForWriting = 2
		'	Set objFSO = CreateObject("Scripting.FileSystemObject")
		'	
		'	Set objFile = objFSO.OpenTextFile(Server.MapPath(savefolder&newname), ForReading)
		'	strText = objFile.ReadAll
		'	objFile.Close
		'
		'	strNewText = Replace(strText, ", - ", ", ")
		'
		'	Set objFile = objFSO.OpenTextFile(Server.MapPath(savefolder&newname), ForWriting)
		'
		'	objFile.write( strNewText)
		'
		'	objFile.Close
		'	
		'	set objFile = Nothing
		'end if
		Session.LCID=2057

		dim rs_csv, rs_prodotti
		
		savefolder="/public/"
		newname="borse.csv"
		'conn.execute ("delete from prodotti")
		
		Set csvconn = Server.CreateObject("ADODB.Connection")
		csvconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _ 
       "Data Source=" & Server.MapPath(savefolder) & ";" & _ 
       "Extended Properties=""text;HDR=YES;FMT=Delimited"""
		session(questofile)=""
		on error resume next
		Set rs_csv = Server.CreateObject("ADODB.Recordset")
		set rs_prodotti = Server.CreateObject("ADODB.Recordset")
		sql="SELECT * FROM "&newname
		
		rs_csv.Open sql, csvconn,3 ,3
		
		
		if err.number<>0 then
			call add2log ("SELECT * FROM "&newname,0)
			response.End
		end if	
		on error goto 0
		n=rs_csv.fields.count-1
		Set Flds = rs_csv.Fields
		dim nome_campo()
		redim nome_campo(n)
		dim valore_campo()
		redim valore_campo(n)
		response.write formatnumber(100.2,2)&"<br>"
		'memorizzo i nomi dei campi
		for i=0 to n
			
			nome_campo(i)=rs_csv.Fields(i).Name
			response.write "Campo: "&nome_campo(i)&":"
			'for h=1 to len(nome_campo(i))
			'	response.write asc(mid(nome_campo(i),h,1))&","
			'next
			response.write "<br>"
			
		next
		response.write "UM:"&um&"<br>"
		Session.LCID=2057
		rs_csv.addnew
		rs_csv("prezzo")="5a5"
		rs_csv.update
		do while not rs_csv.EOF
			codice=rs_csv("codice articolo fornitore")
			'Cerco nei prodotti
			rs_prodotti.open "Select prodotti.* from prodotti where codice='"&codice&"'", conn, 3, 3
			if not rs_prodotti.eof then
				'Articolo esiste
				idpro=rs_prodotti("idpro")
			else
				'Articolo non esiste, lo aggiungo
				rs_prodotti.addnew
				rs_prodotti("codice")=codice
				rs_prodotti("data")=now()
				rs_prodotti("datamod")=now()
				rs_prodotti("descrizione")=rs_csv("Nome commerciale")
				prezzo="a"&cstr(rs_csv("prezzo"))
				response.write prezzo&"<br>"
				'rs_prodotti("prezzo")=rs_csv("prezzo")
				rs_prodotti("um")=rs_csv("Unit"&chr(224)&" di misura")
				rs_prodotti.update
				
			end if
			rs_prodotti.close
				
			
			rs_csv.MoveNext
		loop
		

	
	%>