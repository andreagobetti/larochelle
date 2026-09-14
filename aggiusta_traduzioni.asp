<!--#include virtual="/setup.asp" -->
<%   

Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_co = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")

		sql="select DISTINCTROW traduzioni.pagina, traduzioni.chiave FROM traduzioni GROUP BY traduzioni.pagina, traduzioni.chiave;"
		rs.Open sql, conn, 3, 3
	
		do while not rs.eof
		sql="select * from traduzioni where  pagina = '"&rs("pagina")&"' and lingua='en' and StrComp(trim(chiave), '"&trim(rs("chiave"))&"', 0) = 0"
		response.write sql&"<br>"
		set rs_co=conn.execute(sql)
		if not rs_co.eof then
			inglese=rs_co("valore")
			sql="select * from traduzioni where  pagina = '"&rs("pagina")&"' and lingua='it' and StrComp(trim(chiave), '"&trim(rs("chiave"))&"', 0) = 0"
			rs1.Open sql, conn, 3, 3
			if not rs1.eof then
				rs1("valore_it")=rs1("valore")
				rs1("valore_en")=inglese
				rs1.update
			else
			response.write "Manca chiave IT"&sql
			end if
			rs1.close

		else
			response.write "Nessuna voce trovata" & "<br>"
		end if
			




			rs.movenext
		loop







%>
