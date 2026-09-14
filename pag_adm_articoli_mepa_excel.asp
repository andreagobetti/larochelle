<!--#include virtual="/setup.asp" -->
<!--#include virtual="/multidict_cl.asp" -->
<!--#include virtual="/jsonObject.class.asp" -->
<%
if session("idadmin") = "" then call login()

tdebug=false
idcategoria=request("id")




categoria=conn.execute("select mepa_categorie.descrizione from mepa_categorie where id="&idcategoria)(0)


if not tdebug then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", "attachment; filename=MEPA_"&pulisci_nomefile(categoria)&".xls"
else
	Response.ContentType = "text/html"
end if

set rs=conn.execute("select id, label, obbligatorio, formato, valori_default from mepa_campi where idcategoria="&idcategoria)
arrRS=rs.getrows()
colonne=UBound(arrRS, 2)
Set mepa_campi = New MultiDimensionalDictionary 'Create an Instance of the Class
For i = 0 To colonne
		mepa_campi.SetKey = arrRS(0, i) 'id
		mepa_campi.SetField "label", arrRS(1, i)
		mepa_campi.SetField "obbligatorio", arrRS(2, i)
		mepa_campi.SetField "formato", lcase(left(arrRS(3, i),1))
		mepa_campi.SetField "valori_default", arrRS(4, i)
		mepa_campi.Update 'Bind the record and preapre for the next record
		
	
next


%>
<table width="100%" id="elenco" border="1" cellpadding="0" cellspacing="0" >
		<tr bgcolor="#E5E5E5">
		<%
			For Each Record in mepa_campi.Records
				Response.Write("<td><font face=""Tahoma"" ><strong>" &  Server.HTMLEncode( mepa_campi.Item(Record, "label")) &  "</strong></font></td>")
			Next
		%>
		</tr>
	    <%
		
		
		
		
		set JSON = New JSONobject
	    set rs=conn.execute ("select prodotti.*, prodotti_mepa_campi.mepa_campi from prodotti inner join prodotti_mepa_campi on prodotti.idpro = prodotti_mepa_campi.idpro where idcategoria="&idcategoria&"  order by descrizione")
	    do while not rs.eof 
			mepa_campi_JSON=rs("mepa_campi")
			
			'response.write mepa_campi_JSON
			if mepa_campi_JSON="" then 
				mepa_campi_JSON="[]"
			end if
			JSON.Parse(mepa_campi_JSON)
		
		
	    %>
	    <tr>
		    
		    		<%
			For Each Record in mepa_campi.Records
				response.write "<td>"
				valori_default=mepa_campi.Item(Record, "valori_default")
				richiesto=mepa_campi.Item(Record, "obbligatorio")
				if richiesto="S" then
					richiesto="richiesto"
				else
					richiesto=""
				end if
				valore=JSON(""&Record)
				response.write valore

			Next
		%>
		    			
		    
		    
	    </tr>
	    <%
		    rs.MoveNext
		loop
		set rs = Nothing
	    %>
    </table>
