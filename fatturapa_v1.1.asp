<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/config/fattura_pa_conf.asp" -->
<%

if request("downloada")<>"" then
	

	' Recupero il file da scaricare
	Dim download, file
	file = request("download")
	
	' Creo l'oggetto ADODB.Stream
	Set download = Server.CreateObject("ADODB.Stream")
	
	' Apro la connessione e carico il file
	download.Type = 1
	download.Open
	download.LoadFromFile Server.MapPath(file)
	
	' Aggiungo le intestazioni del tipo di file
	Response.AddHeader "Content-Disposition", "attachment; filename=" & file
	Response.ContentType = "application/octet-stream"
	Response.BinaryWrite download.Read
	
	' Un po di pulizia...
	download.Close
	Set download = Nothing
	response.end	
	
end if
idfat=request("idfat")

if request("download")="" then
	Set rs_fatturepa = Server.CreateObject("ADODB.Recordset")
	rs_fatturepa.Open "select fatturepa.* from fatturepa where idfatt="&idfat, conn, 1, 3
	'if rs_fatturepa.eof then
	rs_fatturepa.addnew
	rs_fatturepa("idfatt")=idfat
	rs_fatturepa.update
	ProgressivoInvio_v=get_last_id("fatturepa")
	rs_fatturepa.close
	rs_fatturepa.Open "select fatturepa.* from fatturepa where id="&ProgressivoInvio_v, conn, 1, 3
	
	nomefile=PA_DatiAnagrafici_IdFiscaleIVA_IdPaese&PA_DatiAnagrafici_IdFiscaleIVA_IdCodice&"_"&ProgressivoInvio_v&".xml"

	rs_fatturepa("nomefile")=nomefile
	rs_fatturepa.update
	rs_fatturepa.close
	set rs_fatturepa=nothing
	
	
	set rs_fatture=conn.execute("select ordini.*, ordini.data as ordini_data, fatture.*, utenti_clienti.FatturaPACodiceDestinatario FROM ((ordini INNER JOIN fatture ON ordini.idord = fatture.idord) INNER JOIN utenti ON ordini.iduser = utenti.iduser) INNER JOIN utenti_clienti ON utenti.iduser = utenti_clienti.iduser where idfat="&idfat)
	
	'Very Important : Set the ContentType property of
	'the Response object to text/xml.
	'Response.ContentType = "text/xml"
	
	totale_fattura=cdbl(rs_fatture("totale_fattura"))
	
	Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	xmlDoc.async = False
	xmlDoc.load( Server.MapPath("/fatturapa_v1.1.xml") )
	
	Set FatturaElettronicaHeader = xmlDoc.createElement("FatturaElettronicaHeader")
	xmlDoc.documentElement.AppendChild FatturaElettronicaHeader
	
		Set DatiTrasmissione = xmlDoc.createElement("DatiTrasmissione")
		FatturaElettronicaHeader.AppendChild DatiTrasmissione
		
			Set IdTrasmittente = xmlDoc.createElement("IdTrasmittente")
			DatiTrasmissione.AppendChild IdTrasmittente
			
				Set IdPaese = xmlDoc.createElement("IdPaese")
				IdPaese.text=PA_IdTrasmittente_IdPaese
				IdTrasmittente.AppendChild IdPaese
				
				Set IdCodice = xmlDoc.createElement("IdCodice")
				IdCodice.text=PA_IdTrasmittente_IdCodice
				IdTrasmittente.AppendChild IdCodice
	
			Set ProgressivoInvio = xmlDoc.createElement("ProgressivoInvio")
			ProgressivoInvio.text=ProgressivoInvio_v
			DatiTrasmissione.AppendChild ProgressivoInvio
	
			Set FormatoTrasmissione = xmlDoc.createElement("FormatoTrasmissione")
			FormatoTrasmissione.text="SDI11"
			DatiTrasmissione.AppendChild FormatoTrasmissione
	
			Set CodiceDestinatario = xmlDoc.createElement("CodiceDestinatario")
			if isnull(rs_fatture("CodiceDestinatario")) or rs_fatture("CodiceDestinatario")="" then
				tmp=split(rs_fatture("FatturaPACodiceDestinatario"),vbcrlf)
				tmp2=split(tmp(0),":")
				CodiceDestinatario.text=tmp2(0)
			else
				CodiceDestinatario.text=rs_fatture("CodiceDestinatario")
			end if
			DatiTrasmissione.AppendChild CodiceDestinatario
			
		Set ContattiTrasmittente = xmlDoc.createElement("ContattiTrasmittente")
			DatiTrasmissione.AppendChild ContattiTrasmittente
	
				Set Telefono = xmlDoc.createElement("Telefono")
				Telefono.text=PA_ContattiTrasmittente_Telefono
				ContattiTrasmittente.AppendChild Telefono
				
				Set Email = xmlDoc.createElement("Email")
				Email.text=PA_ContattiTrasmittente_Email
				ContattiTrasmittente.AppendChild Email
				
		Set CedentePrestatore = xmlDoc.createElement("CedentePrestatore")
		FatturaElettronicaHeader.AppendChild CedentePrestatore
				
			Set DatiAnagrafici = xmlDoc.createElement("DatiAnagrafici")
			CedentePrestatore.AppendChild DatiAnagrafici
				
				Set IdFiscaleIVA = xmlDoc.createElement("IdFiscaleIVA")
				DatiAnagrafici.AppendChild IdFiscaleIVA
				
					Set IdPaese = xmlDoc.createElement("IdPaese")
					IdPaese.text=PA_DatiAnagrafici_IdFiscaleIVA_IdPaese
					IdFiscaleIVA.AppendChild IdPaese
				
					Set IdCodice = xmlDoc.createElement("IdCodice")
					IdCodice.text=PA_DatiAnagrafici_IdFiscaleIVA_IdCodice
					IdFiscaleIVA.AppendChild IdCodice
					
				Set Anagrafica = xmlDoc.createElement("Anagrafica")
				DatiAnagrafici.AppendChild Anagrafica
				
					Set Denominazione_ = xmlDoc.createElement("Denominazione")
					Denominazione_.text=PA_DatiAnagrafici_Anagrafica_Denominazione
					Anagrafica.AppendChild Denominazione_
	
	'				Set CodEORI = xmlDoc.createElement("CodEORI")
	'				CodEORI.text="1234"
	'				Anagrafica.AppendChild CodEORI
	
				Set RegimeFiscale = xmlDoc.createElement("RegimeFiscale")
				RegimeFiscale.text=PA_DatiAnagrafici_RegimeFiscale
				DatiAnagrafici.AppendChild RegimeFiscale
				
			Set Sede = xmlDoc.createElement("Sede")
			CedentePrestatore.AppendChild Sede
	
				Set Indirizzo = xmlDoc.createElement("Indirizzo")
				Indirizzo.text=PA_DatiAnagrafici_Sede_Indirizzo
				Sede.AppendChild Indirizzo
				Set numeroCivico = xmlDoc.createElement("NumeroCivico")
				numeroCivico.text=PA_DatiAnagrafici_Sede_NumeroCivico
				Sede.AppendChild numeroCivico
				Set CAP = xmlDoc.createElement("CAP")
				CAP.text=PA_DatiAnagrafici_Sede_CAP
				Sede.AppendChild CAP
				Set Comune = xmlDoc.createElement("Comune")
				Comune.text=PA_DatiAnagrafici_Sede_Comune
				Sede.AppendChild Comune
				Set Provincia = xmlDoc.createElement("Provincia")
				Provincia.text=PA_DatiAnagrafici_Sede_Provincia
				Sede.AppendChild Provincia
				Set Nazione = xmlDoc.createElement("Nazione")
				Nazione.text=PA_DatiAnagrafici_Sedea_Nazione
				Sede.AppendChild Nazione
	'-----------------------------------------------------------------------------CessionarioCommittente	
		Set CessionarioCommittente = xmlDoc.createElement("CessionarioCommittente")
		FatturaElettronicaHeader.AppendChild CessionarioCommittente
			Set DatiAnagrafici = xmlDoc.createElement("DatiAnagrafici")
			CessionarioCommittente.AppendChild DatiAnagrafici
				'Set IdFiscaleIva = xmlDoc.createElement("IdFiscaleIva")
				'DatiAnagrafici.AppendChild IdFiscaleIva
				
					'Set IdPaese = xmlDoc.createElement("IdPaese")
					'IdPaese.text=PA_IdTrasmittente_IdPaese
					'IdFiscaleIva.AppendChild IdPaese
					
					'Set IdCodice = xmlDoc.createElement("IdCodice")
					'IdCodice.text=rs_fatture("piva")
					'IdFiscaleIva.AppendChild IdCodice
	
	
				Set CodiceFiscale = xmlDoc.createElement("CodiceFiscale")
				CodiceFiscale.text=rs_fatture("cf")
				DatiAnagrafici.AppendChild CodiceFiscale
				
					
				Set Anagrafica = xmlDoc.createElement("Anagrafica")
				DatiAnagrafici.AppendChild Anagrafica
				
					Set Denominazione_ = xmlDoc.createElement("Denominazione")
					Denominazione_.text=rs_fatture("azienda")
					Anagrafica.AppendChild Denominazione_
				
	'			Set RegimeFiscale = xmlDoc.createElement("RegimeFiscale")
	'			RegimeFiscale.text="RF01"
	'			DatiAnagrafici.AppendChild RegimeFiscale
				
			Set Sede = xmlDoc.createElement("Sede")
			CessionarioCommittente.AppendChild Sede
	
				Set Indirizzo = xmlDoc.createElement("Indirizzo")
				Indirizzo.text=rs_fatture("indirizzo")
				Sede.AppendChild Indirizzo
				
				Set CAP = xmlDoc.createElement("CAP")
				CAP.text=rs_fatture("cap")
				Sede.AppendChild CAP
				Set Comune = xmlDoc.createElement("Comune")
				Comune.text=rs_fatture("citta")
				Sede.AppendChild Comune
				Set Provincia = xmlDoc.createElement("Provincia")
				Provincia.text=rs_fatture("provincia")
				Sede.AppendChild Provincia
				Set Nazione = xmlDoc.createElement("Nazione")
				Nazione.text="IT"
				Sede.AppendChild Nazione
		'Set SoggettoEmittente = xmlDoc.createElement("SoggettoEmittente")
		'SoggettoEmittente.text="CP"
		'FatturaElettronicaHeader.AppendChild SoggettoEmittente
	'-------------------------------------------------------------------------------FatturaElettronicaBody
	Set FatturaElettronicaBody = xmlDoc.createElement("FatturaElettronicaBody")
	xmlDoc.documentElement.AppendChild FatturaElettronicaBody
	
		Set DatiGenerali = xmlDoc.createElement("DatiGenerali")
		FatturaElettronicaBody.AppendChild DatiGenerali
		
			Set DatiGeneraliDocumento = xmlDoc.createElement("DatiGeneraliDocumento")
			DatiGenerali.AppendChild DatiGeneraliDocumento
	
				Set TipoDocumento = xmlDoc.createElement("TipoDocumento")
				if totale_fattura>0 then
					TipoDocumento.text="TD01" 'FATTURA
				else
					TipoDocumento.text="TD04" 'NOTA DI CREDITO
				end if
				DatiGeneraliDocumento.AppendChild TipoDocumento
				Set Divisa = xmlDoc.createElement("Divisa")
				Divisa.text="EUR"
				DatiGeneraliDocumento.AppendChild Divisa
				Set Data = xmlDoc.createElement("Data")
				'Data.text=rs_fatture("fatture.data")
				Data.text=aggiustadata(rs_fatture("data"))
				DatiGeneraliDocumento.AppendChild Data
				Set Numero = xmlDoc.createElement("Numero")
				Numero.text=rs_fatture("nfat")
				DatiGeneraliDocumento.AppendChild Numero
				Set ImportoTotaleDocumento = xmlDoc.createElement("ImportoTotaleDocumento")
				ImportoTotaleDocumento.text=aggiustanumero(totale_fattura*(1+rs_fatture("aliquota_iva")/100))
				DatiGeneraliDocumento.AppendChild ImportoTotaleDocumento
				
				'if rs_fatture("impegno_spesa")<>"" then
				'	Set Causale = xmlDoc.createElement("Causale")
				'	Causale.text=aggiustanumero(totale_fattura*(1+rs_fatture("aliquota_iva")/100))
				'	DatiGeneraliDocumento.AppendChild Causale
				'end if
				'if rs_fatture("determina")<>"" then
				'	Set Causale = xmlDoc.createElement("Causale")
				'	Causale.text=aggiustanumero(totale_fattura*(1+rs_fatture("aliquota_iva")/100))
				'	DatiGeneraliDocumento.AppendChild Causale
				'end if
				
				
				'Set Art73 = xmlDoc.createElement("Art73")
				'Art73.text="SI"
				'DatiGeneraliDocumento.AppendChild Art73
				
			'2.1.2
			Set DatiOrdineAcquisto = xmlDoc.createElement("DatiOrdineAcquisto")
			DatiGenerali.AppendChild DatiOrdineAcquisto
				Set RiferimentoNumeroLinea = xmlDoc.createElement("RiferimentoNumeroLinea")
				RiferimentoNumeroLinea.text="1"
				DatiOrdineAcquisto.AppendChild RiferimentoNumeroLinea
				Set IdDocumento = xmlDoc.createElement("IdDocumento")
				Set Data = xmlDoc.createElement("Data")
				if rs_fatture("mepa_testo")<>"" and rs_fatture("mepa_data") then
					IdDocumento.text=rs_fatture("mepa_testo")
					Data.text=aggiustadata(cdate(rs_fatture("mepa_data")))
				else
					IdDocumento.text=rs_fatture("idord")
					Data.text=aggiustadata(rs_fatture("ordini_data"))
				end if
				
				
				DatiOrdineAcquisto.AppendChild IdDocumento
				DatiOrdineAcquisto.AppendChild Data
				Set NumItem = xmlDoc.createElement("NumItem")
				NumItem.text="1"
				DatiOrdineAcquisto.AppendChild NumItem
				if rs_fatture("impegno_spesa")<>"" then
					'2.1.2.5 - Impegno di spesa
					impegno_spesa=rs_fatture("impegno_spesa")
					Set CodiceCommessaConvenzione = xmlDoc.createElement("CodiceCommessaConvenzione")
					CodiceCommessaConvenzione.text=impegno_spesa
					DatiOrdineAcquisto.AppendChild CodiceCommessaConvenzione
				end if
				'if rs_fatture("cup")<>"" then
				'	'2.1.2.6
				'	Set CodiceCUP = xmlDoc.createElement("CodiceCUP")
				'	CodiceCUP.text="CodiceCUP"
				'	DatiOrdineAcquisto.AppendChild CodiceCUP
				'end if
				if rs_fatture("cig")<>"" then
					'2.1.2.7
					Set CodiceCIG = xmlDoc.createElement("CodiceCIG")
					CodiceCIG.text=rs_fatture("cig")
					DatiOrdineAcquisto.AppendChild CodiceCIG
				end if
				
			'Set DatiContratto = xmlDoc.createElement("DatiContratto")
			'DatiGenerali.AppendChild DatiContratto
	
				'Set RiferimentoNumeroLinea = xmlDoc.createElement("RiferimentoNumeroLinea")
				'RiferimentoNumeroLinea.text="1"
				'DatiContratto.AppendChild RiferimentoNumeroLinea
				'Set IdDocumento = xmlDoc.createElement("IdDocumento")
				'IdDocumento.text=rs_fatture("ordini.idord")
				'DatiContratto.AppendChild IdDocumento
				'Set Data = xmlDoc.createElement("Data")
				'Data.text=aggiustadata(rs_fatture("ordini.data"))
				'DatiContratto.AppendChild Data
				'Set NumItem = xmlDoc.createElement("NumItem")
				'NumItem.text="1"
				'DatiContratto.AppendChild NumItem
				'Set CodiceCUP = xmlDoc.createElement("CodiceCUP")
				'CodiceCUP.text="CodiceCUP"
				'DatiContratto.AppendChild CodiceCUP
				'Set CodiceCIG = xmlDoc.createElement("CodiceCIG")
				'CodiceCIG.text="CodiceCIG"
				'DatiContratto.AppendChild CodiceCIG
				
'			Set DatiConvenzione = xmlDoc.createElement("DatiConvenzione")
'			DatiGenerali.AppendChild DatiConvenzione
'	
'				Set RiferimentoNumeroLinea = xmlDoc.createElement("RiferimentoNumeroLinea")
'				RiferimentoNumeroLinea.text="1"
'				DatiConvenzione.AppendChild RiferimentoNumeroLinea
'				Set IdDocumento = xmlDoc.createElement("IdDocumento")
'				IdDocumento.text=rs_fatture("idord")
'				DatiConvenzione.AppendChild IdDocumento
'				Set Data = xmlDoc.createElement("Data")
'				Data.text=aggiustadata(rs_fatture("ordini_data"))
'				DatiConvenzione.AppendChild Data
'				Set NumItem = xmlDoc.createElement("NumItem")
'				NumItem.text="1"
'				'DatiConvenzione.AppendChild NumItem
'				'Set CodiceCUP = xmlDoc.createElement("CodiceCUP")
'				'CodiceCUP.text="CodiceCUP"
'				'DatiConvenzione.AppendChild CodiceCUP
'				Set CodiceCIG = xmlDoc.createElement("CodiceCIG")
'				CodiceCIG.text="CodiceCIG"
'				DatiConvenzione.AppendChild CodiceCIG
'				
'			Set DatiRicezione = xmlDoc.createElement("DatiRicezione")
'			DatiGenerali.AppendChild DatiRicezione
'	
'				Set RiferimentoNumeroLinea = xmlDoc.createElement("RiferimentoNumeroLinea")
'				RiferimentoNumeroLinea.text="1"
'				DatiRicezione.AppendChild RiferimentoNumeroLinea
'				Set IdDocumento = xmlDoc.createElement("IdDocumento")
'				IdDocumento.text=rs_fatture("idord")
'				DatiRicezione.AppendChild IdDocumento
'				Set Data = xmlDoc.createElement("Data")
'				Data.text=aggiustadata(rs_fatture("ordini_data"))
'				DatiRicezione.AppendChild Data
'				Set NumItem = xmlDoc.createElement("NumItem")
'				NumItem.text="1"
'				DatiRicezione.AppendChild NumItem
				'Set CodiceCUP = xmlDoc.createElement("CodiceCUP")
				'CodiceCUP.text="CodiceCUP"
				'DatiRicezione.AppendChild CodiceCUP
				'Set CodiceCIG = xmlDoc.createElement("CodiceCIG")
				'CodiceCIG.text="CodiceCIG"
				'DatiRicezione.AppendChild CodiceCIG
				
			'Set DatiTrasporto = xmlDoc.createElement("DatiTrasporto")
			'DatiGenerali.AppendChild DatiTrasporto
	
			'	Set DatiAnagraficiVettore = xmlDoc.createElement("DatiAnagraficiVettore")
			'	DatiTrasporto.AppendChild DatiAnagraficiVettore
					
			'		Set IdFiscaleIVA = xmlDoc.createElement("IdFiscaleIVA")
			'		DatiAnagraficiVettore.AppendChild IdFiscaleIVA
					
			'			Set IdPaese = xmlDoc.createElement("IdPaese")
			'			IdPaese.text="IT"
			'			IdFiscaleIVA.AppendChild IdPaese
					
			'			Set IdCodice = xmlDoc.createElement("IdCodice")
			'			IdCodice.text="0123456"
			'			IdFiscaleIVA.AppendChild IdCodice
					
			'	Set Anagrafica = xmlDoc.createElement("Anagrafica")
			'	DatiAnagraficiVettore.AppendChild Anagrafica
				
			'		Set Denominazione_ = xmlDoc.createElement("Denominazione")
			'		Denominazione_.text="Abitidalavoro.it"
			'		Anagrafica.AppendChild Denominazione_
					
				'Set DataOraConsegna = xmlDoc.createElement("DataOraConsegna")
				
				'DataOraConsegna.text="2012-10-22T16:46:12.000+02:00"
				'DataOraConsegna.text=aggiustadataora(rs_fatture("ordini.data"))
	
				'DatiTrasporto.AppendChild DataOraConsegna
				
		Set DatiBeniServizi = xmlDoc.createElement("DatiBeniServizi")
		FatturaElettronicaBody.AppendChild DatiBeniServizi
		
			
			'set rs_ordini_dett=conn.execute("select ordini_dett_esteso.* FROM ordini_dett_esteso where prezzo>0 and idord="&rs_fatture("ordini.idord"))
			set rs_ordini_dett=conn.execute(ordini_dett_esteso&" where quantita>0 and idord="&rs_fatture("idord"))
			n_linea=1
			do while not rs_ordini_dett.eof
				Set DettaglioLinee = xmlDoc.createElement("DettaglioLinee")
				DatiBeniServizi.AppendChild DettaglioLinee
			
				Set NumeroLinea = xmlDoc.createElement("NumeroLinea")
				NumeroLinea.text=n_linea
				DettaglioLinee.AppendChild NumeroLinea
				'Set CodiceArticolo = xmlDoc.createElement("CodiceArticolo")
				'CodiceArticolo.text=rs_ordini_dett("codice_ordine")
				'DettaglioLinee.AppendChild CodiceArticolo
				prezzo= converti_typevar14(rs_ordini_dett("prezzo"))
				if isnull(prezzo) or prezzo=0 then
					riga_descrizione=true
					prezzoeff=0
					v_aliquota_iva=0
					v_PrezzoTotale=0
					v_quatita=0
				else
					riga_descrizione=false
					prezzoeff=prezzo-(cdbl(rs_ordini_dett("sconto_prodotto"))/100)*prezzo
					v_aliquota_iva=rs_fatture("Aliquota_IVA")
					v_PrezzoTotale=prezzoeff*rs_ordini_dett("quantita")
					v_quatita=rs_ordini_dett("quantita")
				end if
				
				
				Set Descrizione = xmlDoc.createElement("Descrizione")
				Descrizione.text=replace(rs_ordini_dett("codice_ordine")&" "&rs_ordini_dett("articolo_ordine"),"€","euro")
				DettaglioLinee.AppendChild Descrizione
				
				Set Quantita = xmlDoc.createElement("Quantita")
				Quantita.text=aggiustanumero(v_quatita)
				DettaglioLinee.AppendChild Quantita
				
				if rs_ordini_dett("um")<>"" then
					Set UnitaMisura = xmlDoc.createElement("UnitaMisura")
					UnitaMisura.text=rs_ordini_dett("um")
					DettaglioLinee.AppendChild UnitaMisura
				end if
				
				Set PrezzoUnitario = xmlDoc.createElement("PrezzoUnitario")
				PrezzoUnitario.text=aggiustanumero(prezzoeff)
				DettaglioLinee.AppendChild PrezzoUnitario
				
				Set PrezzoTotale = xmlDoc.createElement("PrezzoTotale")
				PrezzoTotale.text=aggiustanumero(v_PrezzoTotale)
				DettaglioLinee.AppendChild PrezzoTotale
				
				Set AliquotaIVA = xmlDoc.createElement("AliquotaIVA")
				AliquotaIVA.text=aggiustanumero(v_aliquota_iva)
				DettaglioLinee.AppendChild AliquotaIVA
				
				if riga_descrizione then
					Set Natura = xmlDoc.createElement("Natura")
					Natura.text="N2"
					DettaglioLinee.AppendChild Natura

					
					
				end if
				
				'2.2.1.15	-	Capitoli di spesa
				'Set RiferimentoAmministrazione = xmlDoc.createElement("RiferimentoAmministrazione")
				'RiferimentoAmministrazione.Text=
				'DettaglioLinee.AppendChild AltriDatiGestionali
				
				
				'2.2.1.16
'				Set AltriDatiGestionali = xmlDoc.createElement("AltriDatiGestionali")
'				DettaglioLinee.AppendChild AltriDatiGestionali
'					'2.2.1.16.1
'					Set TipoDato = xmlDoc.createElement("TipoDato")
'					TipoDato.text="DET"
'					AltriDatiGestionali.AppendChild TipoDato
				
				
				
				
				
				totord=totord+round(prezzoeff*rs_ordini_dett("quantita"),2)
				
				rs_ordini_dett.movenext
				n_linea=n_linea+1

			loop	
			if cdbl(rs_fatture("trasporto"))>0 then
				Set DettaglioLinee = xmlDoc.createElement("DettaglioLinee")
				DatiBeniServizi.AppendChild DettaglioLinee
				Set NumeroLinea = xmlDoc.createElement("NumeroLinea")
				NumeroLinea.text=n_linea
				DettaglioLinee.AppendChild NumeroLinea
				'Set CodiceArticolo = xmlDoc.createElement("CodiceArticolo")
				'CodiceArticolo.text=rs_ordini_dett("codice_ordine")
				'DettaglioLinee.AppendChild CodiceArticolo
				prezzo= converti_typevar14(rs_fatture("trasporto"))
				prezzoeff=prezzo
				v_aliquota_iva=rs_fatture("Aliquota_IVA")
				v_PrezzoTotale=prezzoeff
				v_quatita=1
				
				
				Set Descrizione = xmlDoc.createElement("Descrizione")
				Descrizione.text="Trasporto"
				DettaglioLinee.AppendChild Descrizione
				
				Set Quantita = xmlDoc.createElement("Quantita")
				Quantita.text=aggiustanumero(v_quatita)
				DettaglioLinee.AppendChild Quantita
				
				
				Set PrezzoUnitario = xmlDoc.createElement("PrezzoUnitario")
				PrezzoUnitario.text=aggiustanumero(prezzoeff)
				DettaglioLinee.AppendChild PrezzoUnitario
				
				Set PrezzoTotale = xmlDoc.createElement("PrezzoTotale")
				PrezzoTotale.text=aggiustanumero(v_PrezzoTotale)
				DettaglioLinee.AppendChild PrezzoTotale
				
				Set AliquotaIVA = xmlDoc.createElement("AliquotaIVA")
				AliquotaIVA.text=aggiustanumero(v_aliquota_iva)
				DettaglioLinee.AppendChild AliquotaIVA
				
				
				totord=totord+prezzoeff

				
				
				
				
				
				
				
			end if
			
			
			'totord=totord+cdbl(rs_fatture("trasporto"))
			sconto=cdbl(rs_fatture("sconto_ordine"))
			sconto_totale=sconto_totale+sconto
			if isnull(sconto) then sconto=0	
			totord=round(totord,2)
			if not isnull(rs_fatture("spese_bancarie")) then totord=totord+cdbl(rs_fatture("spese_bancarie"))
			ImponibileImporto_v=totord
			
			if esenzione_iva(rs_fatture("trattamento_iva")) then
				valore_iva=0
				AliquotaIVA_v=0
				if rs_fatture("trattamento_iva")="10" then
					valore_iva=totord*rs_fatture("aliquota_iva")/100
					AliquotaIVA_v=rs_fatture("aliquota_iva")
					Imposta_v=ImponibileImporto_v*(rs_fatture("aliquota_iva")/100)
				end if
				
			else
				valore_iva=totord*rs_fatture("aliquota_iva")/100
				AliquotaIVA_v=rs_fatture("aliquota_iva")
				Imposta_v=ImponibileImporto_v*(rs_fatture("aliquota_iva")/100)
			end if
			'totord=totord-sconto_totale





			Set DatiRiepilogo = xmlDoc.createElement("DatiRiepilogo")
			DatiBeniServizi.AppendChild DatiRiepilogo
	
				Set AliquotaIVA = xmlDoc.createElement("AliquotaIVA")
				AliquotaIVA.text=aggiustanumero(AliquotaIVA_v)
				DatiRiepilogo.AppendChild AliquotaIVA
				Set ImponibileImporto = xmlDoc.createElement("ImponibileImporto")
				ImponibileImporto.text=aggiustanumero(ImponibileImporto_v)
				DatiRiepilogo.AppendChild ImponibileImporto
				Set Imposta = xmlDoc.createElement("Imposta")
				Imposta.text=aggiustanumero(Imposta_v)
				DatiRiepilogo.AppendChild Imposta
				Set EsigibilitaIVA = xmlDoc.createElement("EsigibilitaIVA")
				if rs_fatture("trattamento_iva")="2" then
					EsigibilitaIVA.text="D" 'DIFFERITA
				elseif rs_fatture("trattamento_iva")="10" then
					EsigibilitaIVA.text="S" 'SPLIT PAYMENT
				else
					EsigibilitaIVA.text="I"
				end if
				DatiRiepilogo.AppendChild EsigibilitaIVA	
				Set RiferimentoNormativo = xmlDoc.createElement("RiferimentoNormativo")
				RiferimentoNormativo.text= trattamento_iva(rs_fatture("trattamento_iva"))
				DatiRiepilogo.AppendChild RiferimentoNormativo
			
				
		Set DatiPagamento = xmlDoc.createElement("DatiPagamento")
		FatturaElettronicaBody.AppendChild DatiPagamento
		
			Set CondizioniPagamento = xmlDoc.createElement("CondizioniPagamento")
			n_scadenze=pagamento_fattura(rs_fatture("pagamento"),2)
			if n_scadenze>1 then
				CondizioniPagamento.text="TP01" 'PAGAMENTO A RATE
			else
				CondizioniPagamento.text="TP02" 'PAGAMENTO COMPLETO
			end if
			DatiPagamento.AppendChild CondizioniPagamento
			
			
			
			Set DettaglioPagamento = xmlDoc.createElement("DettaglioPagamento")
			DatiPagamento.AppendChild DettaglioPagamento
			
				Set ModalitaPagamento = xmlDoc.createElement("ModalitaPagamento")
				if instr(pagamento_fattura(rs_fatture("pagamento"),0),"BONIFICO")>0 then
					ModalitaPagamento.text="MP05" 'BONIFICO
				elseif instr(pagamento_fattura(rs_fatture("pagamento"),0),"RI.BA.")>0 then
					ModalitaPagamento.text="MP12" 'RIBA
				end if
				
				DettaglioPagamento.AppendChild ModalitaPagamento
				
				Set DataScadenzaPagamento = xmlDoc.createElement("DataScadenzaPagamento")
				DataScadenzaPagamento.text=aggiustadata(finemese(rs_fatture("data"),pagamento_fattura(rs_fatture("pagamento"),3)+1))
				DettaglioPagamento.AppendChild DataScadenzaPagamento
				Set ImportoPagamento = xmlDoc.createElement("ImportoPagamento")
				ImportoPagamento.text=aggiustanumero(totord)
				DettaglioPagamento.AppendChild ImportoPagamento
				
				Set IBAN = xmlDoc.createElement("IBAN")
				IBAN.text=PA_DatiPagamento_DettaglioPagamento_IBAN
				DettaglioPagamento.AppendChild IBAN
				
				
			
			
				
	set accountEl = Nothing
	set typeEl = Nothing
	
	xmlDoc.Save Server.MapPath("/public/fatture_pa/"&nomefile)
else
	nomefile=request("download")
	'rs_fatturepa.close
	'set rs_fatturepa=nothing
end if
'response.write xmlDoc.xml

Response.Buffer = False
Dim objStream
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Type = 1 'adTypeBinary
objStream.Open
objStream.LoadFromFile(Server.MapPath("/public"&"/fatture_pa/"&nomefile))
Response.ContentType = "application/x-unknown"
Response.Addheader "Content-Disposition", "attachment; filename=" & nomefile
Response.BinaryWrite objStream.Read
objStream.Close
Set objStream = Nothing

function aggiustadata(data)

	anno=year(data)
	mese=month(data)
	giorno=day(data)
	if len(mese)=1 then mese="0"&mese
	if len(giorno)=1 then giorno="0"&giorno
	aggiustadata=anno&"-"&mese&"-"&giorno

end function
function aggiustadataora(data)

	anno=year(rs_fatture("fatture.data"))
	mese=month(rs_fatture("fatture.data"))
	giorno=day(rs_fatture("fatture.data"))
	if len(mese)=1 then mese="0"&mese
	if len(giorno)=1 then giorno="0"&giorno
	aggiustadataora=anno&"-"&mese&"-"&giorno&"-12:00"

end function

function aggiustanumero(numero)
	if isnull(numero) then
		aggiustanumero=replace(formatnumber(0,2),",",".")
	else
		aggiustanumero=replace(formatnumber(numero,2,,,0),",",".")
	end if
end function

%>