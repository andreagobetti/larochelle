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
<input type='hidden' name='articolo' id='test' style="width:300px;"/><input type='hidden' name='vara' id='vara' style="width:300px;"/>
<input type='hidden' name='varb' id='varb' style="width:300px;"/><input type="submit" Value="Invia"/>

  <script type="text/javascript">


  var vara = [];
  $(document).ready(function () {
  	$('#test').select2({
  		placeholder: 'Cerca articolo',
  		minimumInputLength: 3,
  		allowClear: true,
  		ajax: {
  			quietMillis: 150,
  			url: "ajax_function.asp?select2=articoli_sel2",
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
		$("#test").on("select2-selecting", function(e) {
    //var theID = $(test).val(); // works
    //var theSelection = $(test).filter(':selected').text(); // doesn't work
    var theID = e.object.id;
    var theSelection = e.object.text;
    alert(theID);
    alert(theSelection);
});
$("#test").select2("container").find("div.select2-drop").append("<hr style='margin:5px'><span>Hi, I am a footer.</span>");
  });
</script>
  </form>
</body>
</html>
