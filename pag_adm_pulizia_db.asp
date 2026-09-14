<!--#include virtual="/setup.asp" -->
<%
if session("idadmin") = "" then call login()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<!--#include virtual="/sub_head.asp" -->
</head>
<body>
<center>
  <table height="100%" border="0" cellpadding="<%=cellp%>" cellspacing="<%=cells%>" class="tabellacentrale">     <tr align="center" valign="middle" > 
      <td height="30" colspan="3"> <!--#include virtual="/sub_top.asp" -->
 </td>
<tr> 
      <td width="<%=col_1%>" valign="top"> <!-- Colonna SINISTRA INIZIO-->
        <!--#include virtual="/sub_menu.asp" --><!-- Colonna SINISTRA FINE-->
      </td>
      <td width="<%=col_23%>" align="center" valign="top"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tabella1">
          <tr> 
            <td class=titolomenu>&nbsp;Amministra</td>
          </tr>
		  </table>
        <p>
  <!-- Colonna CORPO INIZIO-->
  <%
	sql = "DELETE carrello.* FROM carrello LEFT JOIN prodotti ON carrello.idpro = prodotti.IDpro WHERE (((prodotti.IDpro) Is Null)); "

	conn.execute(sql)

%>Pulizia record orfani carrello
</p>
        <p>&nbsp; </p>        <!-- Colonna CORPO FINE-->
      </td>
    </tr>
    <tr> 
      <td colspan="2" height="1">&nbsp;</td>
    </tr>
  </table>
</center>
</body>
</html>
