<!--#include virtual="/setup.asp" -->
<form method="post">
	<input type="submit" name="articoli" value="Importa articoli">
	<input type="submit" name="svuota" value="Svuota">
	
	
	
</form>
<%
	if request.form("svuota")<>"" then
		'conn.execute ("delete from utenti where fornitore=1")
		'conn.execute ("delete from settori")
		'conn.execute ("delete from prodotti")
		'conn.execute ("delete from varianti_a")
		'conn.execute ("delete from varianti_b")
		'conn.execute ("delete from magazzino")
	
	end if
	if request.form("articoli")<>"" then
		
		response.write "Importo articoli<br>"
		
	
	
		'carico i dati csv
		Set csv = Server.CreateObject("ADODB.Connection")
		'csv.Open 
       
       driver="Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & Server.MapPath("/public/") & ";Extensions=asc,csv,tab,txt;"
	   driver="Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & Server.MapPath("/public/") & ";" & "Extended Properties=""text;HDR=Yes;FMT=Delimited"""
	   'driver="Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & Server.MapPath("/public/") & ";" & "Extended Properties=""text;HDR=Yes;FMT=TabDelimited"""
	   
	   response.write "Driver:"&driver&"<br>"
	   
	   csv.Open driver
	   
	   
	   
		Set RS_csv = csv.Execute("SELECT * FROM articoli.csv")
		txt=""
		if RS_csv.eof then
			response.write "Il file CSV non contiene record"
			
			call connclose()
			response.End
			
		end if
		
		Response.write "<br>"
		
		
		
		
		Set rs = Server.CreateObject("ADODB.Recordset")
		Set rs_tmp = Server.CreateObject("ADODB.Recordset")
		rs.Open "select * from prodotti", conn, 3, 3
		Do Until RS_csv.EOF
		
			codice=ucase(RS_csv("codice"))
			n=clng(conn.execute("select count(*) from prodotti where codice='"&codice&"'")(0))
			if n=0 then
				'cerco il settore
				idsettore=0
				nome_settore=trim(RS_csv("settore"))
				if nome_settore<>"" then
					sql="select settori.* from settori where nome_settore='"&nome_settore&"'"
					rs_tmp.Open sql, conn, 3, 3
					if rs_tmp.eof then
						Response.write "Creo il settore "&nome_settore&"<br>"
						rs_tmp.addnew
						rs_tmp("nome_settore")=nome_settore
						rs_tmp("idpadre")=0
						rs_tmp("nascondi")=0
						rs_tmp.update
						idsettore=Get_last_id("settori")
					else
						idsettore=rs_tmp("idsettore")
					end if
					rs_tmp.Close
				end if
				'cerco il fornitore
				fornitore=trim(RS_csv("fornitore"))
				if not isnull(fornitore) then fornitore=replace(fornitore,"'","")
				idfor=0
				if fornitore<>"" then
					sql="select utenti.* from utenti where azienda='"&fornitore&"'"
					rs_tmp.Open sql, conn, 3, 3
					if rs_tmp.eof then
						Response.write "Creo il fornitore "&fornitore&"<br>"
						rs_tmp.addnew
						rs_tmp("azienda")=left(fornitore,rs_tmp("azienda").DefinedSize)
						rs_tmp("data")=now()
						rs_tmp("fornitore")=1
						rs_tmp.update
						idfor=Get_last_id("utenti")
					else
						idfor=rs_tmp("iduser")
					end if
					rs_tmp.Close
				end if
				rs.addnew
				rs("data")=now()
				rs("datamod")=now()
				rs("codice")=codice
				'
				response.write RS_csv("descrizione")&"<br>"
				rs("articolo")=left(RS_csv("descrizione"),rs("articolo").DefinedSize)
				rs("um")=rs_csv("unita_misura")
				if isnull(rs_csv("prezzo")) then
					rs("prezzo")=-1
				else
					response.write rs_csv("prezzo")&"<bR>"
					
					rs("prezzo")=cdbl(rs_csv("prezzo"))/100
				end if
				rs("descrizione")=rs_csv(2)
				rs("attivo")=1
				if clng(idfor)>0 then rs("idfor")=idfor
				
				rs.update
				idpro=Get_last_id("prodotti")
				if clng(idsettore)>0 then
					'Creo legame settore
					conn.execute ("insert into settori_prodotti set idpro="&idpro&", idsettore="&idsettore)
				end if
				'Implementare Taglie
				
				
				
				
			
			else
				response.write "<br>Codice "&codice&" gi&agrave; esistente"
			end if
		    RS_csv.MoveNext
		Loop
		RS_csv.Close
		csv.Close
		set csv = Nothing
		call add2log("Analizzato csv ebay "&txt,1)

		
		
		
	end if


%>
