<%
mdb="Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " & server.MapPath("\mdb-database\prova.mdb")
Set conn_tr = Server.CreateObject("ADODB.Connection")
conn_tr.Open mdb
set rs_tr=conn_tr.execute("select * from traduzioni where pagina='pag_cat.asp' and lingua='it'")

Dim dizionario
Set dizionario=Server.CreateObject("Scripting.Dictionary")
do while not rs_tr.eof
	response.write "carico " & rs_tr("chiave")&"<br>"
	chiave=rs_tr("chiave")
	valore=rs_tr("valore")
	dizionario.Add chiave,valore
	rs_tr.movenext
loop
response.write "Caricate "&dizionario.count & " voci<br>"
response.write dizionario.item("teasto1")
conn_tr.close
set conn_tr=nothing
 %>
