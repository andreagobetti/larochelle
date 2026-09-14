<!--#include virtual="/setup.asp" -->
<%   
	response.write "<b>"&request.form("trasferisci")&"</b><br>"

	dim campo(50)
	Set conn_origine = Server.CreateObject("ADODB.Connection")
	Set conn_destinazione = Server.CreateObject("ADODB.Connection")
	conn_origine.Open mdb

	dbpath=server.MapPath("\mdb-database\start_server.mdb")
	mdb="Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & dbpath

	conn_destinazione.Open mdb

	Set rs_origine = conn_origine.execute ("riavvio_server") 'Server.CreateObject("ADODB.Recordset")
	Set rs_destinazione = Server.CreateObject("ADODB.Recordset")
	sql="log"
	rs_destinazione.Open sql, conn_destinazione, 0, 3
	n_aggiunte=0
	do while not rs_origine.eof

			rs_destinazione.addnew
			rs_destinazione("data")=rs_origine("data")
			rs_destinazione("evento")="Application_OnStart"
			rs_destinazione("causale")=1
			rs_destinazione.update
			n_aggiunte=n_aggiunte+1
			
			
			
	loop
	response.write "Aggiunte "&n_aggiunte&" traduzioni"



%>
