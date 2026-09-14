<%
function ControllaPIVA(pi)
if pi = "" or isnull(pi) then
ControllaPIVA = ""
else
'---------------------------------------------------------------------
if Len(pi) <> 11 then
 controllaPIVA = "La lunghezza della partita IVA non &egrave " &_
		"corretta: la partita IVA dovrebbe essere lunga " &_
		"esattamente 11 caratteri."
else
'-----------------------------------------------------------

		Dim objER, result
		 ' istanzia l'oggetto REGULAR EXPRESSION
  		Set objER = New RegExp
		' cerca il pattern in tutta la stringa di input
  		objER.Global = True
		' nessuna differenza fra maiuscole/minuscole
		objER.IgnoreCase = True

 		 '''''''''''''''''''''''''''''''''''''''''''''''''
		  objER.Pattern = "^[0-9]+$"
 		 '''''''''''''''''''''''''''''''''''''''''''''''''

 		 ' verifica la corrispondenza con il pattern
  		result = objER.Test(pi)
		if result <> true then 
		controllaPIVA = "La partita IVA contiene dei caratteri non ammessi: " &_
		"la partita IVA dovrebbe contenere solo cifre."
		Set objER = Nothing
		else 
'------------------------------------------------
Dim s, s1, s2, c, i, char
s1 = 0
	for i = 0 to 9
	i = i + 1 
	char = mid(pi , i , 1 )
		s1 = s1 + asc(char) - asc("0")
'''''''''''''''''''''''''''''''
'controllo dell incremento della variabile
'	response.write(("valore = ")& (asc(char)- asc("0")) & (" s1 = ") & s1 &("<br>") )
''''''''''''''''''''''''''''''''
	next

	for i = 1 to 9 
	i = i + 1 
	char = mid(pi , i , 1 )
		c = 2* ( asc(char) - asc("0"))
			if c > 9 then
			c = c - 9
			s2 = s2 + c
			else
			s2 = s2 + c
			end if
'''''''''''''''''''''''''''''''
'controllo dell incremento della variabile
'	response.write(("valore = ")& (asc(char)- asc("0")) & (" c = ") & c & (" s2 = ") & s2 &("<br>") )
''''''''''''''''''''''''''''''''
			next
			s = s1 + s2
'''''''''''''''''''''''''
'verifica della variabile
'response.Write(s &  ("<br>"))
'''''''''''''''''''''''''
	if( ( 10 - s Mod 10 ) mod 10 <>  asc(Mid(pi, 11, 1)) - asc("0") ) then
		'ControllaPIVA(pi)
		controllaPIVA = "La partita IVA non &egrave; valida: "&_
		"il codice di controllo non corrisponde."
	else
	 controllaPIVA = "" 
	end if

'------------------------------------------------
end if
'------------------------------------------------------------
end if
'---------------------------------------------------------------------
end if 
end function
function ControllaCFnumerico(pi)
if pi = "" or isnull(pi) then
ControllaCFnumerico = ""
else
'---------------------------------------------------------------------
if Len(pi) <> 11 then
 ControllaCFnumerico = "La lunghezza del codice fiscale  non &egrave " &_
		"corretto: il codice fiscale  dovrebbe essere lungo " &_
		"esattamente 11 caratteri."
else
'-----------------------------------------------------------

		Dim objER, result
		 ' istanzia l'oggetto REGULAR EXPRESSION
  		Set objER = New RegExp
		' cerca il pattern in tutta la stringa di input
  		objER.Global = True
		' nessuna differenza fra maiuscole/minuscole
		objER.IgnoreCase = True

 		 '''''''''''''''''''''''''''''''''''''''''''''''''
		  objER.Pattern = "^[0-9]+$"
 		 '''''''''''''''''''''''''''''''''''''''''''''''''

 		 ' verifica la corrispondenza con il pattern
  		result = objER.Test(pi)
		if result <> true then 
		controllaPIVA = "Il codice fiscale contiene dei caratteri non ammessi: " &_
		"Il codice fiscale dovrebbe contenere solo cifre."
		Set objER = Nothing
		else 
'------------------------------------------------
Dim s, s1, s2, c, i, char
s1 = 0
	for i = 0 to 9
	i = i + 1 
	char = mid(pi , i , 1 )
		s1 = s1 + asc(char) - asc("0")
'''''''''''''''''''''''''''''''
'controllo dell incremento della variabile
'	response.write(("valore = ")& (asc(char)- asc("0")) & (" s1 = ") & s1 &("<br>") )
''''''''''''''''''''''''''''''''
	next

	for i = 1 to 9 
	i = i + 1 
	char = mid(pi , i , 1 )
		c = 2* ( asc(char) - asc("0"))
			if c > 9 then
			c = c - 9
			s2 = s2 + c
			else
			s2 = s2 + c
			end if
'''''''''''''''''''''''''''''''
'controllo dell incremento della variabile
'	response.write(("valore = ")& (asc(char)- asc("0")) & (" c = ") & c & (" s2 = ") & s2 &("<br>") )
''''''''''''''''''''''''''''''''
			next
			s = s1 + s2
'''''''''''''''''''''''''
'verifica della variabile
'response.Write(s &  ("<br>"))
'''''''''''''''''''''''''
	if( ( 10 - s Mod 10 ) mod 10 <>  asc(Mid(pi, 11, 1)) - asc("0") ) then
		'ControllaPIVA(pi)
		ControllaCFnumerico = "Il codice fiscale  non &egrave; valido: "&_
		"il codice di controllo non corrisponde."
	else
	 ControllaCFnumerico = "" 
	end if

'------------------------------------------------
end if
'------------------------------------------------------------
end if
'---------------------------------------------------------------------
end if 
end function



function ControllaCF(cf)
	if cf = "" or isnull(cf) then
		controllaCF = ""
else
'------------------------------------------
	if Len(cf) <> 16 and Len(cf)<> 11 then
		controllaCF = "La lunghezza del codice fiscale non &egrave; " &_
				"corretta: il codice fiscale dovrebbe essere " &_ 
				"lungo esattamente 16 o 11 caratteri."
	elseif 	Len(cf)= 11	then
		if isnumeric(cf) then 
			controllaCF =""
		else
			controllaCF = "Il codice fiscale deve contenere solo cifre."
		end if

	else
'--------------------------------------------------
		cf = Ucase(cf)
		Dim objER, result
		 ' istanzia l'oggetto REGULAR EXPRESSION
  		Set objER = New RegExp
		' cerca il pattern in tutta la stringa di input
  		objER.Global = True
		' nessuna differenza fra maiuscole/minuscole
		objER.IgnoreCase = True

 		 '''''''''''''''''''''''''''''''''''''''''''''''''
		  objER.Pattern = "^[\w]+$"
 		 '''''''''''''''''''''''''''''''''''''''''''''''''

 		 ' verifica la corrispondenza con il pattern
  		result = objER.Test(cf)
'response.write(result)
		if result <> true then 
		controllaCF = "Il codice fiscale contiene dei caratteri non validi:" &_
		"i soli caratteri validi sono le lettere e le cifre."
		 Set objER = Nothing
		else 
'-----------------------------------------------------------

	Dim s, c, s1, s2, i
	s1 = 0
	i = 0 
''''''''''''''''''''''''''''''''''''''
'prima versione - meno elegante
'	for i = 1 to 14 
'	i = i + 1 
''''''''''''''''''''''''''''''''''''''
for i = 2 to 14 step 2
		c = Mid(cf, i, 1 )
		if( "0" <= c AND c <= "9" ) then
			s1 = s1 + Asc(c) - Asc("0")
''''''''''''''''''''''''
'controlla il loop
'			response.write("c= "& c & " s1= "& s1 &"<br>")
''''''''''''''''''''''''
		else
			s1 = s1 + Asc(c) - Asc("A")
''''''''''''''''''''''''
'controlla il loop
'			response.write("c= "& c & " s1= "& s1 &"<br>")
''''''''''''''''''''''''
		end if
	next
'''''''''''''''''''''
'controlla la somma delle cifre pari
'	response.write("s1="&s1&"<br>")
''''''''''''''''''''''
	s2 = 0
'''''''''''''''''''''''''''''''
'prima versione - meno elegante
'	for i = 0 to 14
'	i = i + 1 
'''''''''''''''''''''''''''''
for i = 1 to 15 step 2
		c = Mid(cf, i, 1 )
		select Case (c)
		case "0"
		  s2 = s2 + 1
		case "1"
		  s2 = s2 + 0
		case "2"
		  s2 = s2 + 5
		case "3"
		  s2= s2 + 7
		case "4"
		  s2 = s2 + 9
		case "5"
		  s2 = s2 + 13
		case "6"
		  s2 = s2 + 15
		case "7"
		  s2 = s2 + 17
		case "8"
		  s2 = s2 + 19
		case "9"
		  s2 = s2 + 21
		case "A"
		  s2 = s2 + 1
		case "B"
		  s2 = s2 + 0
		case "C"
		  s2 = s2 + 5
		case "D"
		  s2 = s2 + 7
		case "E"
		  s2 = s2 + 9
		case "F"
		  s2 = s2 + 13
		case "G"
		  s2 = s2 + 15
		case "H"
		  s2 = s2 + 17
		case "I"
		  s2 = s2 + 19
		case "J"
		  s2 = s2 + 21
		case "K"
		  s2 = s2 + 2
		case "L"
		  s2 = s2 + 4
		case "M"
		  s2 = s2 + 18
		case "N"
		  s2 = s2 + 20
		case "O"
		  s2 = s2 + 11
		case "P"
		  s2 = s2 + 3
		case "Q"
		  s2 = s2 + 6
		case "R"
		  s2 = s2 + 8
		case "S"
		  s2 = s2 + 12
		case "T"
		  s2 = s2 + 14
		case "U"
		  s2 = s2 + 16
		case "V"
		  s2 = s2 + 10
		case "W"
		  s2 = s2 + 22
		case "X"
		  s2 = s2 + 25
		case "Y"
		  s2 = s2 + 24
		case "Z"
		  s2 = s2 + 23
		End select
''''''''''''''''''''''''
'controlla il loop
'		response.write("c= "& c & " s2= "& s2 &"<br>")		
''''''''''''''''''''''''
		next
		s= s1+s2
''''''''''''''''''''''''
'controlla la somma dispari
'			response.write("s2="&s2&"<br>")
'controlla il totale
'		response.write(s&"<br>")
''''''''''''''''''''''''		
		if chr((s Mod 26) + Asc("A")) <> Mid(cf, 16, 1) then
		controllaCF = "Il codice fiscale non &egrave; corretto: " &_
		"il codice di controllo non corrisponde. "&left(cf,15)&"<b>"&Mid(cf, 16, 1)&"</b> dovrebbe essere: <b>"&chr((s Mod 26) + Asc("A"))&"</b>"
		else
		controllaCF = ""
		end if
'-----------------------------------------------------------
		end if		
'--------------------------------------------------
	end if
'------------------------------------------
end if 
end function
%>

