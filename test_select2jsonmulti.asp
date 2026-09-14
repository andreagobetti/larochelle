<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->

<%
Dim Js
Set Js = jsArray()



For Each item In Request.Querystring
    str_add2log= str_add2log&"(" & item & "):" & Request.Querystring(item) & "|"
Next
str_add2log= str_add2log&"<br><b>Form</b>:<br>"
For Each item In Request.Form
   str_add2log= str_add2log&"(" & item & "):" & Request.Form(item) & "<BR />"
Next
response.write str_add2log
%>

<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>settori select2</title>
<link rel="stylesheet" type="text/css" href="Jquery/css/select2.css">

  <style type="text/css" media="all">
    /* fix rtl for demo */
    .chosen-rtl .chosen-drop { left: -9000px; }
  </style>
</head>
<body>
  <form>
<!--#include virtual="/jquery.inc" -->

<script src="Jquery/js/select2.js"></script>
<input type='hidden' name='settori' id='settori' style="width:500px;" value=""/>
<input type="submit" Value="Invia"/>

  <script type="text/javascript">
  <%
  idpro=1038
set rsm=conn.execute("select settori_prodotti.IDpro, settori.idsettore, settori.Nome_Settore FROM settori INNER JOIN settori_prodotti ON settori.idsettore = settori_prodotti.IDsettore where idpro="&idpro)
	do while not rsm.eof
		Set Js(Null) = jsObject()
		js(null)("id")=rsm("idsettore")
		Js(null)("text")=rsm("Nome_Settore")
		rsm.movenext
	loop
	rsm.close
	
%>


  var vara = <%js.Flush%>;
  $(document).ready(function () {
  	$('#settori').select2({
  		tags: true,
  		multiple: true,
  		placeholder: 'Cerca articolo',
  		minimumInputLength: 0,
  		allowClear: true,
  		ajax: {
  			quietMillis: 150,
  			url: "ajax_function.asp?select2=settori",
  			dataType: 'json',
  			data: function (term, page) {
  				return {
  					term: term
  				};
  			},
  			results: function (data) {
  				return {
  					results: data
  				};
  			}
  		}
  		
  	});
  	$("#settori").select2("data",vara);

  	$("#settori").select2("container").find("ul.select2-choices").sortable({
	    containment: 'parent',
	    start: function() { $("#settori").select2("onSortStart"); },
	    update: function() { $("#settori").select2("onSortEnd"); }
	});
  	
  	

$("#settori").select2("container").find("div.select2-drop").append("<hr style='margin:5px'><span>Hi, I am a footer.</span>");
  });
</script>
  </form>
  
  <a href="<%=questofile%>">ricarica</a>
</body>
</html>
