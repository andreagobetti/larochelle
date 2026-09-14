<%
function conto(stato)
	if not isnumeric(stato) then
		conto=stato
	else
		select case stato
		case -1
		conto=36
		case 0
		conto=36
		case 1
		conto="Abbuoni e arrot passivi"
		case 2
		conto="Assicurazioni auto"
		case 3
		conto="Assicurazioni varie"
		case 4
		conto="Beni ammortizzabili"
		case 5
		conto="Bolli auto"
		case 6
		conto="Cancelleria e stampati"
		case 7
		conto="Canone affitto"
		case 8
		conto="Canoni leasing"
		case 9
		conto="Carburanti e lubrif."
		case 10
		conto="Compensi amministratori"
		case 11
		conto="Consulenze commercialisti"
		case 12
		conto="Consulenze del lavoro"
		case 13
		conto="Contributi enasarco"
		case 14
		conto="Imposte e tasse deducibili"
		case 15
		conto="Imposte e tasse indeducibili"
		case 16
		conto="Interessi passivi"
		case 17
		conto="Luce-acqua-gas"
		case 18
		conto="Manut. e rip. auto"
		case 19
		conto="Manut. e rip. varie"
		case 20
		conto="Materiale di consumo"
		case 21
		conto="Materie prime e sussid."
		case 22
		conto="Multe e sanzioni indeducibili"
		case 23
		conto="Pedaggi autostrada"
		case 24
		conto="Prodotti per rivendita"
		case 25
		conto="Provvig. agenti dio commercio"
		case 26
		conto="Servizi internet"
		case 27
		conto="Software"
		case 28
		conto="Spese bancarie"
		case 29
		conto="Spese condominiali"
		case 30
		conto="Spese di rappresentanza"
		case 31
		conto="Spese di trasferta"
		case 32
		conto="Spese legali e notarili"
		case 33
		conto="Spese locazione locali affitto"
		case 34
		conto="Spese postali"
		case 35
		conto="Spese telefoniche"
		case 36
		conto="Stipendi"
		
		end select
	end if
end function
%>