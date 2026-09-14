<!--#include virtual="/setup.asp" -->

<%
'Option Explict

Class ordine
'######################################################################
'Named Dictionary Recordset Object
'######################################################################

Public SetKey
Private toc
Private AddNewRecord
'######################################################################
Private Sub Class_Initialize
	response.write "Class_Initialize<br>"


End Sub
'######################################################################
Private Sub Class_Terminate
	response.write "Class_Terminate<br>"
	Set toc = Nothing
	Set AddNewRecord = Nothing
End Sub
'######################################################################



Public Function imposta(byVal tabella, byVal idord)
	Set toc = Server.CreateObject("Scripting.Dictionary")
	toc.add "ciao","CIAO"
	toc.add "cia1o","CIAO"
	toc.add "cia2o","CIAO"
	
	
		sql="select ordini.*, utenti.trattamento_iva,utenti.email, utenti.telefono, utenti.fax, utenti.cellulare, utenti.tipologia,utenti.note_trasporto_2 FROM ordini INNER JOIN utenti ON ordini.iduser = utenti.iduser"' where idord="&m_idord
		'response.write sql
		'Set rs_ordine = Server.CreateObject("ADODB.Recordset")
		set rs_ordine=conn.execute (sql)
		'rs_ordine.Open sql, conn, 1, 3

		if not rs_ordine.eof then
			m_esiste=true
			'memorizzo i nomi dei campi e i valori attuali

			m_n_campi=rs_ordine.fields.count-1
			'response.write "m_n_campi:"&m_n_campi&"<br>"
			Set Flds = rs_ordine.Fields
			
			for i=0 to m_n_campi
				value=rs_ordine(i)
				if VarType(value)=14 then value=cdbl(Value)
				toc.add rs_ordine.Fields(i).Name,value
			next
		else
			m_esiste=false
		end if
	
	
	
	'On Error Goto 0
End Function

Public sub elenco_campi
	response.write "elenco_campi<br>"
	
	For Each Key in toc
	  Response.Write Key & " is " & toc.item(Key) & "<br>"
	Next
	
	'For Each key in ordine_Dictionary.keys
	'    Response.Write key&"="&ordine_Dictionary.item(key)&"<br>"
	'Next
	


end sub





Public Sub Update
	toc.Add SetKey, AddNewRecord
	Set AddNewRecord = Server.CreateObject("Scripting.Dictionary")
End Sub
'######################################################################
Public Function GetCollection
	Set GetCollection = toc
End Function
'######################################################################
Public Sub SetField(byVal Key, byVal Value)
If AddNewRecord.Exists(Key) = False Then
 AddNewRecord.Add Key, Value
Else
 AddNewRecord(Key) = Value
End If
End Sub
'######################################################################
Public Function Records
 Records = toc.Keys
End Function
'######################################################################
Public Function Fields(byVal PrimaryKey)
Fields = toc(PrimaryKey).Keys
End Function
'######################################################################
Public Function Item(byVal Key, byVal Value)
On Error Resume Next
If toc.Item(Key).Exists(Value) Then
 Item = toc.Item(Key).Item(Value)
End If
On Error Goto 0
End Function
'######################################################################
Public Function Exists(byVal Key, byVal Value)
'On Error Resume Next
If IsNull(Value) Or Value = "" Then
 Exists = toc.Item(Key).Exists
Else
 Exists = toc.Item(Key).Exists(Value)
End If
'On Error Goto 0
End Function

Public Function Count
 Count = toc.Count
End Function
'######################################################################
End class

%>
<html>
<body>
<%
'######################################################################
'Example Usage
Dim Dict    'The Dict Object
Dim Record    'The Record Object
Dim Field    'The Field Object
Set Dict = New ordine 'Create an Instance of the Class
dict.imposta "ordini" ,  10
dict.elenco_campi 
%>