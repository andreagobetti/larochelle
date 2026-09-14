<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->

<%
idord=request("idord")
%>
<form id="backup_spettanze">
		<input name="idord" type="hidden" value="<%=idord%>">
		
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1"   >
          <tr >
            <td  style="border-bottom:1px solid;">Data</td>
            <td  style="border-bottom:1px solid;">Descrizione</td>
            <td  style="border-bottom:1px solid;">Ripristinato</td>
          </tr>
          <%
	      set rs= conn.execute ("select * from rollback where idord="&idord&" order by data ")    
	      do while not rs.eof    
	          
	       %>   
          <tr >
			<td  style="border-bottom:1px solid;">
				<%=rs("data")%>
			</td>
			<td  style="border-bottom:1px solid;">
				<%=rs("descrizione")%>
			</td>
			<td  style="border-bottom:1px solid;">
				<%if rs("eseguito")=1 then %>SI<%end if %>
			</td>
          </tr>
          
          <%
	          rs.MoveNext
	          Loop
	          
	          set rs = Nothing
	          call connclose()
	          %>
        </table>
        </form>
