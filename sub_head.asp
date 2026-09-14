<meta name="Copyright" content="©2016 Andrea Spotorno">
<meta name="author" content="Andrea Spotorno">
<%=Application("head")%>
<!--#include virtual="/jquery.inc" -->
<link rel="stylesheet" type="text/css" href="pag_adm.css?<%=ver_js_css%>" title="stile">
<link rel="stylesheet" type="text/css" href="/config/pag_adm_custom.css?v=1" title="stile">
<%if session("idadmin")<>"" then%>
<script src="https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.js" type="text/javascript"></script>
<link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.css" />
<link rel="stylesheet" href="jquery/css/superfish.css" media="screen">
<link rel="stylesheet" type="text/css" href="Jquery/css/select2.css">
<script src="js/console-shim.min.js"></script>

<script src="Jquery/js/hoverIntent.js" type="text/javascript"></script>
<script src="Jquery/js/superfish.js" type="text/javascript"></script>
<script type="text/javascript" src="Jquery/js/jquery.sticky-kit.min.js"></script>
<script src="Jquery/js/jquery.ui-contextmenu.min.js" type="text/javascript"></script>
<script type="text/javascript" src="pag_adm.js?<%=ver_js_css%>"></script>
<script type="text/javascript">
	var posizione_errore="";
// Superfish
$(function(){$('.sf-menu').superfish({speed:'fast',speedOut:'fast',delay:400});$("ul.sf-menu li").addClass("ui-state-default");$("ul.sf-menu li").hover(function () { $(this).addClass('ui-state-hover'); },function () { $(this).removeClass('ui-state-hover'); });});


	window.onerror = function (msg, url, line, column, errorObj) {
			//if (Math.random() > 0.1) return; // only log some errors
			var txt="window.onerror ADMIN, Message : " + msg+"<br>url: " + url+"<br>Line number: " + line+"<br>column: "+column+"<br>errorobj: "+errorObj+"<br>Useragent: "+navigator.userAgent+"<br>Posizione errore: "+posizione_errore;
				$(function() {
					$.ajax({
						cache: false,
						url     : "searcher.asp",
						type    : "post",
						data	: "txt_errore="+ encodeURIComponent(txt)

						});

				});
				<%if utente_andrea then %>
				alert(txt);
				<%end if %>
            }

</script>
<style>
.titolo_admin
{
	HEIGHT: 24px;
	padding-top: 5px;
	padding-left: 5px;
	margin-top:0px;
}
.titolo_admin a
{
	color:#F00;
	FONT-WEIGHT: bold;
	font-size: 14px;
}
#barra_fissa{z-index: 9999;}
.ui-autocomplete.ui-widget {
  font-size: 12px;
}

/* Jquery TAB Modifica dei margini*/
.ui-tabs .ui-tabs-panel {
    padding: 0 0 0 0;
}
#first_tab{
	font-weight: bold;
	color: red;
}
</style>
<%if session("visualizza_messaggio")<>"" then
dim toastr,n_toastr

toastr=split(session("visualizza_messaggio"),"||")
n_toastr=ubound(toastr)

%>

<script>


<%="//"&n_toastr%>
$(function() {
		toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0",  "extendedTimeOut": "0","closeButton": false};
		<%for n=0 to n_toastr%>
		<%="//for n="&n%>
		<%
		toast=split(toastr(n),"|")

		if n<n_toastr then%>
		toastr.<%=toast(0)%>("<%=toast(1)%>", "<%=time()%>");
		<%else


				%>



		toastr.<%=toast(0)%>("<%=toast(1)%>", "<%=time()%>",
		    {onclick: function() {
				console.log ("Richiesta reset_messaggio");

		    $.ajax({
		    			url     : "ajax_function.asp?oper=reset_messaggio",
		    			type    : "post",
		    			dataType: 'json',
		    			cache: false,
		    			success: function(data){
			    			console.log(data.Message);

		    			}
		    			,error:function(xhr, textStatus, error){
		    			      console.log("xhr.statusText:"+xhr.statusText);
		    			      console.log("xhr.responseText:"+xhr.responseText);
		    			      console.log("textStatus:"+textStatus);
		    			      console.log("error:"+error);
		    				  }
		    		});

		    }}
		);
		<%end if

			next
		%>

});
</script>

<%end if%>
<%end if%>
