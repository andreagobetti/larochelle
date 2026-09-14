<!--#include virtual="/setup.asp" -->
<%   
if request.form("svuota")<>"" then
	Set conn_destinazione = Server.CreateObject("ADODB.Connection")

	dbpath=server.MapPath("\mdb-database\traduzioni.mdb")
	mdb="Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & dbpath
	conn_destinazione.Open mdb
	conn_destinazione.execute "DELETE traduzioni.id FROM traduzioni;",n
	response.write n&" traduzioni eliminate<br>"
 
end if
if request.form("trasferisci")<>"" then
	response.write "<b>"&request.form("trasferisci")&"</b><br>"

	dim campo(50)
	Set conn_origine = Server.CreateObject("ADODB.Connection")
	Set conn_destinazione = Server.CreateObject("ADODB.Connection")
	if request.form("trasferisci")="Da sito a backup" then
		conn_origine.Open mdb

		dbpath=server.MapPath("\mdb-database\traduzioni.mdb")
		mdb="Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & dbpath

		conn_destinazione.Open mdb
	else
		conn_destinazione.Open mdb
		dbpath=server.MapPath("\mdb-database\traduzioni.mdb")
		mdb="Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & dbpath
		conn_origine.Open mdb
	end if

	Set rs_origine = conn_origine.execute ("select traduzioni.* from traduzioni") 'Server.CreateObject("ADODB.Recordset")
	Set rs_destinazione = Server.CreateObject("ADODB.Recordset")

	n_aggiunte=0
	do while not rs_origine.eof
		sql="select * from traduzioni where  pagina = '"&rs_origine("pagina")&"' and chiave = '"&rs_origine("chiave")&"'"
		rs_destinazione.Open sql, conn_destinazione, 0, 3

		if rs_destinazione.eof then
			n=rs_origine.fields.count-1
			for i=0 to n
				campo(i)=rs_origine(i)
			next			
			rs_destinazione.addnew
			for i=1 to n
					rs_destinazione(i)=campo(i)
			next
			rs_destinazione.update
			response.write "Trasferita pagina:"&rs_destinazione("pagina")&" chiave: "&rs_destinazione("chiave")&"<br>"
			n_aggiunte=n_aggiunte+1
			
			
			
		end if
		rs_destinazione.close
		rs_origine.movenext
	loop
	response.write "Aggiunte "&n_aggiunte&" traduzioni"

end if



%>
      <form name="form1" method="post" action="<%=questofile%>">
        <input type="submit" name="trasferisci" value="Da sito a backup">
        <input type="submit" name="trasferisci" value="Da backup a sito">
		<input type="submit" name="svuota" value="Svuota backup">

      </form>