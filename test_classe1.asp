<!--#include virtual="/setup.asp" -->

<%
Class Test
    Private m_s
    Private m_i
    Private m_idord
    Private m_tabella
	Private m_n_campi
	dim m_recordset
	dim m_campi_recordset()
	Private m_esiste
	dim rs_ordine
	
	dim ordine_Dictionary

	
    Public Default Function Init(parameters)
         Select Case UBound(parameters)
             Case 0
                Set Init = InitOneParam(parameters(0))
             Case 1
                Set Init = InitTwoParam(parameters(0), parameters(1))
             Case else
                Set Init = Me
         End Select
    End Function

    Private Function InitOneParam(parameter1)
        If TypeName(parameter1) = "String" Then
            m_s = parameter1
        Else
			m_tabella="ordini"
            m_idord = parameter1
        End If
        Set InitOneParam = Me
		call ApriRecord()
    End Function

    Private Function InitTwoParam(parameter1, parameter2)
        m_tabella = parameter1
        m_idord = parameter2
        Set InitTwoParam = Me
		call ApriRecord()
    End Function
	Public Property Get idord()
		idord=m_idord
	end property
	
	Private sub ApriRecord
		dim rs_ordine
		dim n
		dim flds
		sql="select ordini.*, utenti.trattamento_iva,utenti.email, utenti.telefono, utenti.fax, utenti.cellulare, utenti.tipologia,utenti.note_trasporto_2 FROM ordini INNER JOIN utenti ON ordini.iduser = utenti.iduser"' where idord="&m_idord
		response.write sql
		'Set rs_ordine = Server.CreateObject("ADODB.Recordset")
		set rs_ordine=conn.execute (sql)
		'rs_ordine.Open sql, conn, 1, 3

		if not rs_ordine.eof then
			m_esiste=true
			'memorizzo i nomi dei campi e i valori attuali

			m_n_campi=rs_ordine.fields.count-1
			'response.write "m_n_campi:"&m_n_campi&"<br>"
			Set Flds = rs_ordine.Fields
			redim m_campi_recordset(m_n_campi)
			redim m_recordset(m_n_campi)
			for i=0 to m_n_campi
				'response.write "i:"&i &" rs_ordine.Fields(i).Name="&rs_ordine.Fields(i).Name&"<br>"
				m_campi_recordset(i)=rs_ordine.Fields(i).Name
			
				m_recordset(i)=rs_ordine(i)
				if vartype(m_recordset(i))=14 then m_recordset(i)=cdbl(m_recordset(i))
				
			next
		else
			m_esiste=false
		end if
		
		
	   
    End sub
	Public sub elenco_campi
			for i=0 to m_n_campi
				response.write m_campi_recordset(i)&"="&m_recordset(i)&" tipo="&TypeName(m_recordset(i))&"<br>"
			next
	end sub
	
End Class

set a= (new test)(array("ordini",1))
response.write a.idord
a.elenco_campi

%>