<%
'Sintassi

	'set modificheRS= (new ClasseModificheRS)(oper)
	'modifichers.leggi(rs)	
	'val=modificheRS.confronta(rs)	





Class ClasseModificheRs
	Private m_log_classe
	Private m_log
	Private m_modifiche_rs
	Private n_fields
	private m_oper
	Private nome_campo()
	Private valore_campo()

	
	
    Public Default Function Init(oper)
	    m_log=false
	    m_modifiche_rs=""
	    m_oper=oper
	    m_log_classe="Init oper:"&m_oper&"<br>"
	    Set Init = Me
    End Function
    
    Private Sub Class_Terminate
	    if m_log then call add2log(m_log_classe,0)
	End Sub
	
	
	Public sub leggi(rs)
 		n_fields=rs.fields.count-1
		redim nome_campo(n_fields)
		redim valore_campo(n_fields)
		if m_log then
			m_log_classe=m_log_classe&"Leggo RS, trovati "&n_fields&" campi<br>"
			on error resume next
		end if
		'memorizzo i nomi dei campi e i valori attuali
		for i=0 to n_fields
			if m_log then m_log_classe=m_log_classe&nome_campo(i)&"="&valore_campo(i)&"<br>"
			nome_campo(i)=rs.Fields(i).Name
			valore_campo(i)=rs(i).value
		next
	end sub
	
	Public function confronta(rs)
		'on error resume next
		if m_log then
			m_log_classe=m_log_classe&"Confronto RS<br>"
			on error resume next
		end if
		for i=0 to n_fields
		
			m_log_classe=m_log_classe&nome_campo(i)&"="&valore_campo(i)&"<br>"
			'if err.number<>0 then
			'	response.write "<b>Errore:</b>"&m_log_classe
			'	response.end
			'end if
			if oper="update" then
				if (valore_campo(i)<>rs(i)) or (valore_campo(i)<>null and rs(i)<>"") then
				'if (valore_campo(i)<>rs(i).value)  then
					m_modifiche_rs=m_modifiche_rs&"modificato campo "&nome_campo(i)&" da: "&valore_campo(i)&" a: "&rs(i)&"<br>"
				end if
			else
				
				if (valore_campo(i)<>rs(i)) or (valore_campo(i)<>null and rs(i)<>"") then
					m_modifiche_rs=m_modifiche_rs&"impostato campo "&nome_campo(i)&" a: "&rs(i)&"<br>"
				end if
			end if
		next
		if err.number<>0 then
			call add2log(m_log_classe,-1)
			response.end
		end if

		confronta=m_modifiche_rs
	end function

	Public function elenca(rs)
		if m_log then
			m_log_classe=m_log_classe&"Elenco RS<br>"
			on error resume next
		end if
 		n_fields=rs.fields.count-1
		redim nome_campo(n_fields)
		redim valore_campo(n_fields)
		if m_log then
			m_log_classe=m_log_classe&"Leggo RS, trovati "&n_fields&" campi<br>"
			on error resume next
		end if
		
		'memorizzo i nomi dei campi e i valori attuali
		for i=0 to n_fields
			m_modifiche_rs=m_modifiche_rs&rs.Fields(i).Name&": "&rs(i).value&"<br>"
		next
		elenca=m_modifiche_rs
	end function
	
End Class

%>