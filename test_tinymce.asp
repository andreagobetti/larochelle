<!--#include virtual="/setup.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<title><%=application("brwstitle")%></title>
<!-- Load jQuery -->
<script type="text/javascript" src="http://www.google.com/jsapi"></script>
<script type="text/javascript">
	google.load("jquery", "1");
</script>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"/>

<!--#include virtual="/sub_head.asp" -->
<script language="javascript" type="text/javascript" src="/tiny_mce/tiny_mce.js"></script>
<script language="javascript" type="text/javascript" src="tiny_mce_conf1.js"></script>
</head>
<body>


 <textarea name="testata" class="mceEditor" cols="40" rows="20"><%=testo%></textarea></td>
          
</body>
</html>
<%
rsClose
%>
