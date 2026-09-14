<!--#include virtual="/setup.asp" -->
<%
response.write nomesito
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento senza titolo</title>
</head>

<body>
<%if false then%>
<form action="https://www.paypal.com/cgi-bin/webscr" method="POST">
 <%else%>
<form action="https://www.sandbox.paypal.com/cgi-bin/webscr" method="POST">
 <%end if%>
<input name="cmd" type="hidden" value="_xclick" />
 <input name="business " type="hidden" value="andico_1362326156_biz@alice.it" />
 <input name="item_name" type="hidden" value="Ordine 97" />
 <input name="item_number" type="hidden" value="97" />
 <input name="amount" type="hidden" value="1" />
 <input name="currency_code" type="hidden" value="EUR" />
 <input name="return" type="hidden" value="" />
 <input name="notify_url" type="hidden" value="http://<%= lcase(nomesito)%>/pag_paypal_ipn.asp" />
 <input name="submit" type="submit" value="» PAY WITH PAYPAL" />
 
</form>

</body>
</html>
