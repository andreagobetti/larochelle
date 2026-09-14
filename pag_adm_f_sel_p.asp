<!--#include virtual="/setup.asp" -->
<%
id=request("id")
sql="select * FROM prodotti where idpro= " & id
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn, 3, 3
if request.form("oper")<>"" then
	if request.form("oper")="Cancella immagine" then
		rs("fileimg")=""
		rs("dimimg")=""
		rs.update
		rsclose
	elseif request.form("oper")="Immagine non disponibile" then
		rs("fileimg")="nofoto.gif"
		rs("dimimg")=""
		rs.update
		rsclose
	end if	
	response.redirect "pag_adm_artic.asp"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

<!--#include virtual="/sub_head.asp" -->
</head>
<body>
<center>
  <table height="100%" border="0" cellpadding="<%=cellp%>" cellspacing="<%=cells%>" class="tabellacentrale"> 
    <tr align="center" valign="middle"> 
      <td height="30"> <%=titolo_top%> </td>
    </tr>
    <tr> 
      <td valign="top"><center>
    <table border="1" cellpadding="0" cellspacing="0" class="tabella1"> 
      <tr> 
        <td align="center">&nbsp;Codice&nbsp;</td> 
        <td align="left">&nbsp;Articolo&nbsp;</td> 
      </tr> 
      <tr> 
        <td align="center"><b>&nbsp;<%=rs("codice")%>&nbsp;</b></td> 
        <td align="left"><b>&nbsp;<%=rs("articolo")%>&nbsp;</b></td> 
      </tr> 
      <tr> 
        <td colspan="2" align="center"><FORM METHOD="Post" ENCTYPE="multipart/form-data" ACTION="pag_adm_f_upl_p.asp?id=<%=id%>&addvar=<%=request.querystring("addvar")%>"> 
            <p align="center"><font face="Verdana" size="2"> Inserisci il percorso per l'immagine<br> 
              <INPUT TYPE="file" NAME="blob" size="50"> 
              <BR> 
              <INPUT TYPE="submit" NAME="Enter" value="Carica immagine..."> 
              </font> 
          </FORM></td> 
      </tr> 
      <tr> 
        <td colspan="2" align="center"> <form name="form1" method="post" action="<%=questofile%>" style="margin:0;"> 
            <input type="hidden" name="id" value="<%=id%>"> 
            <table border="0" cellpadding="2" cellspacing="0"> 
              <tr align="center"> 
                <td>  
                  <input type="submit" name="oper" value="Cancella immagine" ></td> 
                <td> 
                  <input type="submit" name="oper" value="Utilizza immagine non disponibile" ></td> 
                <td> 
                  <input type="submit" name="oper" value="Nessuna variazione" style="font.size:10px;"></td> 
              </tr> 
            </table> 
          </form>
</td></tr></table></center>
</td>
      <!-- Codice controllo pagina qui-->
      <!-- FINE Codice controllo pagina qui-->
    </tr>
    <tr> 
      <td height="1">&nbsp;</td>
    </tr>
  </table>
</center>
</body>
</html>
<%
rsClose
%>