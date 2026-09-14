<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!--#include virtual="/config/fattura_pa_conf.asp" -->
<%

idfat=request("idfat")	
Set rs_fatturepa = Server.CreateObject("ADODB.Recordset")
rs_fatturepa.Open "select fatturepa.* from fatturepa where idfatt="&idfat, conn, 1, 3
if true then
'if rs_fatturepa.eof then
	rs_fatturepa.addnew
	rs_fatturepa("idfatt")=idfat
	ProgressivoInvio_v=cint(rs_fatturepa("id"))
	nomefile=PA_DatiAnagrafici_IdFiscaleIVA_IdPaese&PA_DatiAnagrafici_IdFiscaleIVA_IdCodice&"_"&ProgressivoInvio_v&".xml"
	rs_fatturepa("nomefile")=nomefile
	rs_fatturepa.update
	rs_fatturepa.close
	set rs_fatturepa=nothing
	
	
	set rs_fatture=conn.execute("select fatture.*, ordini.*, utenti_clienti.FatturaPACodiceDestinatario FROM ((ordini INNER JOIN fatture ON ordini.idord = fatture.idord) INNER JOIN utenti ON ordini.iduser = utenti.iduser) INNER JOIN utenti_clienti ON utenti.iduser = utenti_clienti.iduser where idfat="&idfat)
	
	'Very Important : Set the ContentType property of
	'the Response object to text/xml.
	'Response.ContentType = "text/xml"
	
	Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	xmlDoc.async = False
	xmlDoc.load( Server.MapPath("/fatturapa_v1.0.xml") )
	
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
			FormatoTrasmissione.text=PA_FormatoTrasmissione
			DatiTrasmissione.AppendChild FormatoTrasmissione
	
			Set CodiceDestinatario = xmlDoc.createElement("CodiceDestinatario")
			CodiceDestinatario.text=rs_fatture("FatturaPACodiceDestinatario")
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
				TipoDocumento.text="TD01"
				DatiGeneraliDocumento.AppendChild TipoDocumento
				Set Divisa = xmlDoc.createElement("Divisa")
				Divisa.text="EUR"
				DatiGeneraliDocumento.AppendChild Divisa
				Set Data = xmlDoc.createElement("Data")
				'Data.text=rs_fatture("fatture.data")
				Data.text=aggiustadata(rs_fatture("fatture.data"))
				DatiGeneraliDocumento.AppendChild Data
				Set Numero = xmlDoc.createElement("Numero")
				Numero.text=rs_fatture("nfat")
				DatiGeneraliDocumento.AppendChild Numero
				'Set Art73 = xmlDoc.createElement("Art73")
				'Art73.text="SI"
				'DatiGeneraliDocumento.AppendChild Art73
	
			Set DatiOrdineAcquisto = xmlDoc.createElement("DatiOrdineAcquisto")
			DatiGenerali.AppendChild DatiOrdineAcquisto
	
				Set RiferimentoNumeroLinea = xmlDoc.createElement("RiferimentoNumeroLinea")
				RiferimentoNumeroLinea.text="1"
				DatiOrdineAcquisto.AppendChild RiferimentoNumeroLinea
				Set IdDocumento = xmlDoc.createElement("IdDocumento")
				IdDocumento.text=rs_fatture("ordini.idord")
				DatiOrdineAcquisto.AppendChild IdDocumento
				Set Data = xmlDoc.createElement("Data")
				Data.text=aggiustadata(rs_fatture("ordini.data"))
				DatiOrdineAcquisto.AppendChild Data
				Set NumItem = xmlDoc.createElement("NumItem")
				NumItem.text="1"
				DatiOrdineAcquisto.AppendChild NumItem
				'Set CodiceCUP = xmlDoc.createElement("CodiceCUP")
				'CodiceCUP.text="CodiceCUP"
				'DatiOrdineAcquisto.AppendChild CodiceCUP
				'Set CodiceCIG = xmlDoc.createElement("CodiceCIG")
				'CodiceCIG.text="CodiceCIG"
				'DatiOrdineAcquisto.AppendChild CodiceCIG
				
			Set DatiContratto = xmlDoc.createElement("DatiContratto")
			DatiGenerali.AppendChild DatiContratto
	
				Set RiferimentoNumeroLinea = xmlDoc.createElement("RiferimentoNumeroLinea")
				RiferimentoNumeroLinea.text="1"
				DatiContratto.AppendChild RiferimentoNumeroLinea
				Set IdDocumento = xmlDoc.createElement("IdDocumento")
				IdDocumento.text=rs_fatture("ordini.idord")
				DatiContratto.AppendChild IdDocumento
				Set Data = xmlDoc.createElement("Data")
				Data.text=aggiustadata(rs_fatture("ordini.data"))
				DatiContratto.AppendChild Data
				Set NumItem = xmlDoc.createElement("NumItem")
				NumItem.text="1"
				DatiContratto.AppendChild NumItem
				Set CodiceCUP = xmlDoc.createElement("CodiceCUP")
				CodiceCUP.text="CodiceCUP"
				DatiContratto.AppendChild CodiceCUP
				Set CodiceCIG = xmlDoc.createElement("CodiceCIG")
				CodiceCIG.text="CodiceCIG"
				DatiContratto.AppendChild CodiceCIG
				
			Set DatiConvenzione = xmlDoc.createElement("DatiConvenzione")
			DatiGenerali.AppendChild DatiConvenzione
	
				Set RiferimentoNumeroLinea = xmlDoc.createElement("RiferimentoNumeroLinea")
				RiferimentoNumeroLinea.text="1"
				DatiConvenzione.AppendChild RiferimentoNumeroLinea
				Set IdDocumento = xmlDoc.createElement("IdDocumento")
				IdDocumento.text=rs_fatture("ordini.idord")
				DatiConvenzione.AppendChild IdDocumento
				Set Data = xmlDoc.createElement("Data")
				Data.text=aggiustadata(rs_fatture("ordini.data"))
				DatiConvenzione.AppendChild Data
				Set NumItem = xmlDoc.createElement("NumItem")
				NumItem.text="1"
				DatiConvenzione.AppendChild NumItem
				'Set CodiceCUP = xmlDoc.createElement("CodiceCUP")
				'CodiceCUP.text="CodiceCUP"
				'DatiConvenzione.AppendChild CodiceCUP
				'Set CodiceCIG = xmlDoc.createElement("CodiceCIG")
				'CodiceCIG.text="CodiceCIG"
				'DatiConvenzione.AppendChild CodiceCIG
				
			Set DatiRicezione = xmlDoc.createElement("DatiRicezione")
			DatiGenerali.AppendChild DatiRicezione
	
				Set RiferimentoNumeroLinea = xmlDoc.createElement("RiferimentoNumeroLinea")
				RiferimentoNumeroLinea.text="1"
				DatiRicezione.AppendChild RiferimentoNumeroLinea
				Set IdDocumento = xmlDoc.createElement("IdDocumento")
				IdDocumento.text=rs_fatture("ordini.idord")
				DatiRicezione.AppendChild IdDocumento
				Set Data = xmlDoc.createElement("Data")
				Data.text=aggiustadata(rs_fatture("ordini.data"))
				DatiRicezione.AppendChild Data
				Set NumItem = xmlDoc.createElement("NumItem")
				NumItem.text="1"
				DatiRicezione.AppendChild NumItem
				'Set CodiceCUP = xmlDoc.createElement("CodiceCUP")
				'CodiceCUP.text="CodiceCUP"
				'DatiRicezione.AppendChild CodiceCUP
				'Set CodiceCIG = xmlDoc.createElement("CodiceCIG")
				'CodiceCIG.text="CodiceCIG"
				'DatiRicezione.AppendChild CodiceCIG
				
			Set DatiTrasporto = xmlDoc.createElement("DatiTrasporto")
			DatiGenerali.AppendChild DatiTrasporto
	
				Set DatiAnagraficiVettore = xmlDoc.createElement("DatiAnagraficiVettore")
				DatiTrasporto.AppendChild DatiAnagraficiVettore
					
					Set IdFiscaleIVA = xmlDoc.createElement("IdFiscaleIVA")
					DatiAnagraficiVettore.AppendChild IdFiscaleIVA
					
						Set IdPaese = xmlDoc.createElement("IdPaese")
						IdPaese.text="IT"
						IdFiscaleIVA.AppendChild IdPaese
					
						Set IdCodice = xmlDoc.createElement("IdCodice")
						IdCodice.text="0123456"
						IdFiscaleIVA.AppendChild IdCodice
					
				Set Anagrafica = xmlDoc.createElement("Anagrafica")
				DatiAnagraficiVettore.AppendChild Anagrafica
				
					Set Denominazione_ = xmlDoc.createElement("Denominazione")
					Denominazione_.text="Abitidalavoro.it"
					Anagrafica.AppendChild Denominazione_
					
				'Set DataOraConsegna = xmlDoc.createElement("DataOraConsegna")
				
				'DataOraConsegna.text="2012-10-22T16:46:12.000+02:00"
				'DataOraConsegna.text=aggiustadataora(rs_fatture("ordini.data"))
	
				'DatiTrasporto.AppendChild DataOraConsegna
				
		Set DatiBeniServizi = xmlDoc.createElement("DatiBeniServizi")
		FatturaElettronicaBody.AppendChild DatiBeniServizi
		
			
			set rs_ordini_dett=conn.execute("select ordini_dett_esteso.* FROM ordini_dett_esteso where prezzo>0 and idord="&rs_fatture("ordini.idord"))
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
				Set Descrizione = xmlDoc.createElement("Descrizione")
				Descrizione.text=left(rs_ordini_dett("codice_ordine")&" "&rs_ordini_dett("articolo_ordine"),100)
				DettaglioLinee.AppendChild Descrizione
				Set Quantita = xmlDoc.createElement("Quantita")
				Quantita.text=aggiustanumero(rs_ordini_dett("quantita"))
				DettaglioLinee.AppendChild Quantita
				'Set UnitaMisura = xmlDoc.createElement("UnitaMisura")
				'UnitaMisura.text=rs_ordini_dett("um")
				'DettaglioLinee.AppendChild UnitaMisura
				
				prezzoeff=rs_ordini_dett("prezzo")-(rs_ordini_dett("sconto_prodotto")/100)*rs_ordini_dett("prezzo")
				Set PrezzoUnitario = xmlDoc.createElement("PrezzoUnitario")
				PrezzoUnitario.text=aggiustanumero(prezzoeff)
				DettaglioLinee.AppendChild PrezzoUnitario
				Set PrezzoTotale = xmlDoc.createElement("PrezzoTotale")
				PrezzoTotale.text=aggiustanumero(prezzoeff*rs_ordini_dett("quantita"))
				DettaglioLinee.AppendChild PrezzoTotale
				Set AliquotaIVA = xmlDoc.createElement("AliquotaIVA")
				AliquotaIVA.text=aggiustanumero(rs_fatture("Aliquota_IVA"))
				DettaglioLinee.AppendChild AliquotaIVA
				
				totord=totord+round(prezzoeff*rs_ordini_dett("quantita"),2)
				'response.write prezzoeff&" "&totord
				rs_ordini_dett.movenext
				n_linea=n_linea+1

			loop	
			totord=totord+rs_fatture("trasporto")	
			sconto=rs_fatture("ordini.sconto_ordine")
			sconto_totale=sconto_totale+sconto
			if isnull(sconto) then sconto=0	
			totord=round(totord,2)
			if not isnull(rs_fatture("spese_bancarie")) then totord=totord+rs_fatture("spese_bancarie")
			ImponibileImporto_v=totord
			if esenzione_iva(rs_fatture("trattamento_iva")) then
				valore_iva=0
				AliquotaIVA_v=0
			else
				valore_iva=totord*rs_fatture("aliquota_iva")/100
				totord=totord*(1+rs_fatture("aliquota_iva")/100)
				AliquotaIVA_v=rs_fatture("aliquota_iva")
				Imposta_v=totord*(rs_fatture("aliquota_iva")/100)
			end if
			totord=totord-sconto_totale





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
					EsigibilitaIVA.text="D"
				else
					EsigibilitaIVA.text="I"
				end if
				DatiRiepilogo.AppendChild EsigibilitaIVA				
				
		'Set DatiPagamento = xmlDoc.createElement("DatiPagamento")
		'FatturaElettronicaBody.AppendChild DatiPagamento
		
		'	Set CondizioniPagamento = xmlDoc.createElement("CondizioniPagamento")
		'	CondizioniPagamento.text="TP01"
		'	DatiPagamento.AppendChild CondizioniPagamento
		'	Set DettaglioPagamento = xmlDoc.createElement("DettaglioPagamento")
		'	DatiPagamento.AppendChild DettaglioPagamento
		'		Set ModalitaPagamento = xmlDoc.createElement("ModalitaPagamento")
		'		ModalitaPagamento.text="MP01"
		'		DettaglioPagamento.AppendChild ModalitaPagamento
		'		Set DataScadenzaPagamento = xmlDoc.createElement("DataScadenzaPagamento")
		'		DataScadenzaPagamento.text="2012-12-31"
		'		DettaglioPagamento.AppendChild DataScadenzaPagamento
		'		Set ImportoPagamento = xmlDoc.createElement("ImportoPagamento")
		'		ImportoPagamento.text=aggiustanumero(totord)
		'		DettaglioPagamento.AppendChild ImportoPagamento
			
			
				
	set accountEl = Nothing
	set typeEl = Nothing
	
	xmlDoc.Save Server.MapPath("/public/fatture_pa/"&nomefile)
else
	nomefile=rs_fatturepa("nomefile")
	rs_fatturepa.close
	set rs_fatturepa=nothing
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

	anno=year(rs_fatture("fatture.data"))
	mese=month(rs_fatture("fatture.data"))
	giorno=day(rs_fatture("fatture.data"))
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
		aggiustanumero="NULL"
	else
		aggiustanumero=replace(formatnumber(numero,2),",",".")
	end if
end function

%>