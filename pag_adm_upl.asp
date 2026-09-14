<!--#include file="setup.asp" -->
<%
if session("idadmin") = "" then call login()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=application("brwstitle")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

<!--#include file="sub_head.asp" -->
</head>
<body>
<center>
  <table border="0" height="100%" cellpadding="<%=cellp%>" cellspacing="<%=cells%>">    <tr align="center" valign="middle" > 
      <td height="30" colspan="3"> <!--#include file="sub_top.asp" -->
 </td>
<tr> 
      <td width="<%=col_1%>" valign="top"> 
        <!--#include file="sub_menu.asp" -->
      </td>
      <td width="<%=col_23%>" align="center" valign="top"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="tabella1">
          <tr> 
            <td class=titolomenu>&nbsp;Amministra</td>
          </tr>
	    </table>
              
        <!-- Tabella centrale CORPO INIZIO-->
        <table width="100%" border="0" cellspacing="0" cellpadding="2" class="tabella1">
          <tr> 
            <td class="ui-widget-header">&nbsp;Cartella upload </td>
          </tr>
          <tr> 
            <td><DIV id=example><IFRAME style="WIDTH: 100%; height:500;" 
src="/public/dirlisttile.asp" frameBorder=0 
scrolling=yes></IFRAME></DIV> </td>
          </tr>
        </table> 
        <!-- Tabella centrale CORPO FINE-->
      </td>
</tr>
    <tr> 
      <td colspan="2" height="1">&nbsp;</td>
    </tr>
  </table>
</center>
</body>
</html>
