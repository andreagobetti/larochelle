<%
'-------------------------Funzioni per incassi INIZIO

class cl_Causale_Ddt
	'Proprietà
	'Public n_Elementi
	dim lista,lista_nonord
	
	'Costruttore
	Private Sub Class_Initialize()
		set lista_nonord = Server.CreateObject("System.Collections.Sortedlist")
		with lista_nonord
			.add "VENDITA",1
			.add "CONTO VENDITA",2
			.add "CONTO LAVORAZIONE",3
			.add "CONTO VISIONE",4
			.add "RESO",5
			.add "RESO NON CONFORME",6
			.add "RESO DA C/LAVORO",7
			.add "RESO DA C/VISIONE",8
			.add "RESO DA RIPARARE",9
			.add "RESO DA SOSTITUIRE",10
			.add "RESO RIPARATO",11
			.add "RESO RIPARATO IN GARANZIA",12
			.add "SOSTITUZIONE",13
			.add "OMAGGIO",14
		end with
		set lista=Server.CreateObject("Scripting.Dictionary")
		For i = 0 To lista_nonord.Count - 1
            lista.add lista_nonord.GetByIndex(i),lista_nonord.GetKey(i)
        Next
		set lista_nonord=nothing
	end sub
	
	'Distruttore
	Private Sub Class_Terminate()
		set lista=nothing
	End Sub
	
	Public function elemento(key)
		elemento=lista.item(key)
	end function
	
	public Property Get n_Elementi ()
	   n_Elementi=lista.count
	End Property
	
	public Property Get stampa_option ()
		For Each Entry In lista
		  Response.write Entry &":"&lista(Entry) & "<br />"
		Next
	End Property
end class


%>
