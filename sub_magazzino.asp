<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

idpro=request("idpro")
aggiorna_inventario=false

' Aggiorna quantità
IF Request( "aggiorna" ) <> "" THEN
	idpro=request("idpro")
	sql = "select * FROM magazzino WHERE idpro=" & idpro
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	WHILE NOT RS.EOF
	if Request( "idmag" & RS( "idmag" ) )<>"" then 'controllo esistenza record nella pagina
		newQ = TRIM( Request( "quantita_" & RS( "idmag" ) ) )
		IF isNumeric( newQ ) THEN
			RS("quantità") = newQ
			if request.form("inventario")="si" then
				rs("inventario")=newQ
			end if
		END IF
		newQ = TRIM( Request( "riordino_" & RS( "idmag" ) ) )
		IF isNumeric( newQ ) THEN
			RS("quantità_riordino") = newQ
		END IF
		newQ = TRIM( Request( "ordinato_" & RS( "idmag" ) ) )
		IF isNumeric( newQ ) THEN
			RS("quantità_ordinata") = newQ
		END IF
		rs("data_aggiornamento")=now()
		if isnull(rs("idvara")) then rs("idvara")=0
		if isnull(rs("idvarb")) then rs("idvarb")=0
		rs.update
	end if
	RS.MoveNext
	WEND
	RS.Close
	'Aggiorna dati prodotto
	sql = "select * FROM prodotti WHERE idpro=" & request("idpro")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql, conn, 3, 3
	if request.form("vendita")="si" then
		rs("vendita")=true
	else
		rs("vendita")=false
	end if
	rs("disponibilita")=request.form("disponibilita")
	rs.update
	add2log "Aggiornato magazzino articolo [articolo="& idpro&"]"&rs("codice")&"[/articolo] eseguito da [admin="&session("idadmin")&"]"&session("nominativo")&"[/admin]", 2
	rs.close
END IF
'crea e visualizza magazzino
IF idpro <> "" THEN
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

</head>
<body>
<center>
 <form name="form1" method="post" action="<%=questofile%>"> 
<%
	Set rs_mag = Server.CreateObject("ADODB.Recordset")
	sql = "select * FROM prodotti WHERE idpro=" & idpro
	set rs_pro=conn.execute(sql)
	if not rs_pro.eof then
	Set rs_var_a = Server.CreateObject("ADODB.Recordset")
	Set rs_var_b = Server.CreateObject("ADODB.Recordset")

	rs_var_a.Open "select * from varianti_a where idpro=" & idpro, conn, 3, 3
	rs_var_b.Open "select SQL_CALC_FOUND_ROWS * from varianti_b where idpro=" & idpro, conn, 3, 3
	lngTotalRecords=clng(conn.Execute("Select Found_Rows();")(0).Value)
	variante_a=false
	variante_b=false
	if not rs_var_a.eof then variante_a=true
	if not rs_var_b.eof then variante_b=true

		%><table width="100%" class='tabella1' style='border-collapse: collapse;'><%

		n_colonne=lngTotalRecords
		if n_colonne=0 then n_colonne=1
		colspan=n_colonne*6
	
		%><tr><td colspan=" <%=colspan%>" class='titolobox'>magazzino</td></tr>
      <tr><%
		do until rs_var_b.eof
			%><td colspan="6" align="center" style='border: 1px solid black; border-collapse: collapse;' ><%=rs_var_b("variante_b")%></td><%
					rs_var_b.movenext
					loop

		%></tr><tr>
                      
                      <%for i = 1 to n_colonne
					  %>
<td>&nbsp;</td>
    <td class='corpobox' style='border-bottom: 1px solid black; ' align='center'>A magazzino</td>
    <td class='corpobox' style='border-bottom: 1px solid black;' align='center'>Scorta</td>
    <td class='corpobox' style='border-bottom: 1px solid black;' align='center'>In ordine</td>
    <td class='corpobox' style='border-bottom: 1px solid black;' align='center'>Aggiornato</td>
    <td class='corpobox' style='border-bottom: 1px solid black;border-right:1px solid black;' align='center'>Venduti</td>
    <%
					next
		%>
    </tr><%
		'crea le righe con variante_a
		do 
				%><tr><%
				'response.Write "Ciclo varianti1<br>"
		if variante_b then rs_var_b.movefirst		
		do 
					'response.Write "Ciclo varianti2<br>"
					'response.write "variante 2"& replace(varianti2(ii),"'","''")
					
					sql="select * from magazzino where idpro="& idpro  
					sql_totali="select ordini_dett.idpro, ordini_dett.idvara, ordini_dett.idvarb, Sum(ordini_dett.quantita) AS SommaDiquantita, Count(ordini.idord) AS ConteggioDiidord FROM ordini INNER JOIN ordini_dett ON ordini.idord = ordini_dett.idord WHERE (((Year(data))=Year(Now()))) GROUP BY ordini_dett.idpro, ordini_dett.idvara, ordini_dett.idvarb HAVING (((ordini_dett.idpro)="& idpro &")"
					if variante_a then
						sql=sql&" and idvara=" & rs_var_a("idvara") 
						sql_totali=sql_totali &"  AND ((ordini_dett.idvara)="& rs_var_a("idvara")&") "
					end if
					if variante_b then
						sql=sql&" and idvarb=" & rs_var_b("idvarb")
						sql_totali=sql_totali &"   AND ((ordini_dett.idvarb)="& rs_var_b("idvarb")&")"

					end if
					sql_totali=sql_totali &" )"
'response.write sql_totali
					rs_mag.Open sql, conn, 3, 3
					set rs_totali=conn.execute(sql_totali)
					if ii=0 then
					%><td style='border: 1px solid black; border-collapse: collapse;' class='corpobox'><%if not rs_var_a.eof then response.write rs_var_a("variante_a")%></td><%
					end if
					%>
                    <td  class='corpobox' style='border-bottom: 1px solid black;  border-left: 1px solid black;' align='center'><%
					if rs_mag.eof then
						'response.write "Aggiungo record<br>"
						rs_mag.addnew
						rs_mag("idpro")=idpro
						if variante_a then
							rs_mag("idvara")=rs_var_a("idvara")
						else
							rs_mag("idvara")=0						
						end if
						if variante_b then
							rs_mag("idvarb")=rs_var_b("idvarb")
						else
							rs_mag("idvarb")=0
						end if
						rs_mag("quantità_riordino")=0
						rs_mag("quantità")=0
						rs_mag("quantità_riordino")=0
						rs_mag.update
					else
					%>
						
			      <%
					end if
					if rs_mag("quantità")<rs_mag("quantità_riordino") then
					stile=" style='background:#F00'"
					else
					stile=""
					end if
					%><input type="hidden" name="idmag<%=rs_mag("idmag")%>" value="SI" />
						<%=rs_mag("quantità")%>
<%
					%></td>
					<td  class='corpobox' style='border-bottom: 1px solid black;' align='center'><%=rs_mag("quantità_riordino")%>
                        </td>
                        <td  class='corpobox' style='border-bottom: 1px solid black;' align='center'><%=rs_mag("quantità_ordinata")%></td>
<td  class='corpobox' style='border-bottom: 1px solid black;' align='center'><%if rs_mag("data_aggiornamento")<>"" then%> <%=formatDateTime(rs_mag("data_aggiornamento"), vbShortDate)%><%end if%></td><td  class='corpobox' style='border-bottom: 1px solid black;' align='center'><%if not rs_totali.eof then%><%=rs_totali("sommadiquantita")%>/<%=rs_totali("conteggiodiidord")%><%end if%>
</td>
					
					<%
					rs_mag.close
				if not rs_var_b.eof then rs_var_b.movenext	
				loop until rs_var_b.eof%>
				</tr>
<%'loop per variante_a	
if not rs_var_a.eof then rs_var_a.movenext
loop until rs_var_a.eof%>
		</table>
    <%
	end if
	%>
    
 
  </form></center></body>
    <%
end if	
%>