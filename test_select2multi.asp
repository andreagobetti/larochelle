<!--#include virtual="/setup.asp" -->

<%
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
  <title>Test select2</title>
<link rel="stylesheet" type="text/css" href="Jquery/css/select2.css">

  <style type="text/css" media="all">
    /* fix rtl for demo */
    .chosen-rtl .chosen-drop { left: -9000px; }
  </style>
</head>
<body>
  <form>
<script src="http://code.jquery.com/jquery-1.10.1.min.js"></script>
<script src="Jquery/js/select2.js"></script>
<%val=",2,"%>
  <select name="tipologia"  id="tipologia" multiple style="width:300px;">
  <%
set rsm=conn.execute("select * from utenti_tipologia order by tipologia;")
do while not rsm.eof
%>
            <option value="<%=rsm("id")%>" <%if instr(val,","&rsm("id")&",")>0 then response.write " selected"%>><%=rsm("tipologia")%></option>
            <%
rsm.movenext
loop
rsm.close
%>
          </select>
  <script type="text/javascript">
var form;
var termine;
  function aggiungi(){
	 	$.get("pag_adm_tipologia.asp?ajax=true&tipologia="+termine, function(result){
		  if (result!="")
			{
				$(form).append($('<option>', {value:result, text: termine}));
				 
				 $(form).trigger("change");
				 var data=$(form).select2("data");
				 data.push({"id":result,"text":termine});
				 $(form).select2("data", data); 
			}
			else
			{
//				toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
//				toastr.error('Errore in aggiornamento prenotazione ordine');
			}
		});	

	 
	return true;
	}
  $(document).ready(function () {
	form=  $('#tipologia');
  	$(form).select2({
	formatNoMatches:function(term){
			termine=term;
            return "Nessuna corrispondenza trovata, <button onclick='return aggiungi();'>aggiungi "+term+"  </button>";
        }}
	);
	

  	
  });
</script> 
<input type="submit" value="invia"/>
  </form>
</body>
</html>
