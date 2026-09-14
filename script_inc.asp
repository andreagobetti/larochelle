	<script src="js/console-shim.min.js"></script>
	<script src="js/bootstrap.min.js"></script>
    <script src="js/smoothscroll.js"></script>
	<script src="js/jquery.debouncedresize.js"></script>
    <script src="js/retina.min.js"></script>
    <script src="js/jquery.placeholder.js"></script>
    <script src="js/jquery.hoverIntent.min.js"></script>
    <%if false then%>
	<script src="js/twitter/jquery.tweet.min.js"></script>
	<%end if%>
	<script src="js/jquery.flexslider-min.js"></script>
    <script src="js/owl.carousel.min.js"></script>
	<script src="js/jflickrfeed.min.js"></script>
	<script src="js/jquery.prettyPhoto.js"></script>
	<script src="js/jquery-ui.js"></script>
	<%if page_script="product" then%>
	<script src="js/jquery.fitvids.js"></script>
    <script src="js/jquery.elastislide.js"></script>
    <script src="js/jquery.elevateZoom.min.js"></script>
    <%end if%>
	<%if page_script="default" then%>
    <script src="js/jquery.themepunch.plugins.min.js"></script>
    <script src="js/jquery.themepunch.revolution.min.js"></script>
    <%end if%>
    <%if js_toastr then%>
		<script src="Jquery/js/toastr.min.js" type="text/javascript"></script>
	<%end if%>
	<%if css_form then%>        
    	<script src="js/bootstrap-switch.min.js"></script>
	<%end if%>
		
	<%if utente_admin then %>
	<script src="Jquery/js/hoverIntent.js" type="text/javascript"></script>
	<script src="Jquery/js/superfish.js" type="text/javascript"></script>
	<%end if %>
	
	<script src="Jquery/Js/jquery.cookiebar.js"></script>
	<script src="js/main.js"></script>
	<script>
	
	
	window.onerror = function (msg, url, line) {
		//if (Math.random() > 0.1) return; // only log some errors
		var skip="";
		if (line==0 || line=="0")
		{
			skip="Line 0";
		}
		else if (sessionStorage.getItem("error") != line){
			sessionStorage.setItem("sessionStorage","memorizzo");
			sessionStorage.setItem("error",line)
			skip="report"
		}
		else
		{
			sessionStorage.setItem("sessionStorage","memorizzato");
		}

		$(function() {
			if(skip=="report"){
					var txt="window.onerror: "+location.href+" Message : " + msg+"<br>url : " + url+"<br>Line number : " + line+"<br>Useragent : "+navigator.userAgent+"<br>Skip : "+skip+"<br>sessionStorage : "+sessionStorage.getItem("sessionStorage");

					$.ajax({
						cache: false,
						url     : "searcher.asp",
						type    : "post",
						data	: "txt_errore="+ encodeURIComponent(txt)
						
						});
			}
				
		});
	}
    $(function() {

        $("#cerca").autocomplete({
			source: "searcher.asp?cosa=articolihome",
			cache: false ,
	        minLength: 2,
			maxCacheLength:0,
			open: function () {
		        $(this).data("uiAutocomplete").menu.element.css('width', '').addClass("dropdown-menu  col-md-3 col-sm-6 col-xs-12");
		    },
			select: function( event, ui ) {
				if (ui.item.tipo=="P"){
					document.location.href="product.asp?P=x&idpro="+ui.item.id
				}
				if (ui.item.tipo=="S"){
					document.location.href="category.asp?idsettore="+ui.item.id
					}
				if (ui.item.tipo=="T"){
					$(".quick-search-form").submit();
					}
				}
	
	    }).data( "ui-autocomplete" )._renderItem = function( ul, item ) {
	<%if true then%>
			switch(item.tipo){
			case "D":
				var inner_html = '<div class="list_item_container"><strong>' + item.articolo + '</strong></div>';
				break;	
			case "P":
			 var inner_html = '<a><div class="list_item_container"><div class="item-code">' + item.codice + '</div><div class="item-code">' + item.articolo + '</div></div></a>';
			 break;
 			case "S":
			 var inner_html = '<a><div class="list_item_container"><div class="item-code">' + item.codice + '</div><div class="item-code">' + item.articolo + '</div></div></a>';
 			case "T":
			 var inner_html = '<a><div class="list_item_container"><strong>' + item.articolo + '</strong></div></a>';
			}
			 
			 
			 <%else%>
	       var inner_html = '<a><div class="list_item_container"><div class="image"><img src="send_img.asp?s=100&imgprod=' + item.img + '"></div><div class="label">' + item.codice + '</div><div class="description">' + item.articolo + '</div></div></a>';<%end if%>
	        return $( "<li></li>" )
	            .data( "item.autocomplete", item )
	            .append(inner_html)
	            .appendTo( ul );
	    };

	$(".addfavorite").click(function(e) {
				e.preventDefault();
				var this_obj=$(this);
				var idpro=this_obj.closest(".item").attr("id").substring(6);
				console.log("idpro: "+idpro);
				$.ajax({
							url     : "searcher.asp",
							type    : "post",
							dataType: 'json',
							data	: "oper=favouritetrue&idpro=" + idpro,
							success: function(data){
								if (data.success==true)
								{
									if (data.action=="add")
									{	
										console.log(data.Message);
										$(this_obj).addClass("active");
										//$("#alert-add-remove").hide();
										//$("#alert-add-pre").slideDown();
									}
									else
									{
										console.log("errore:"+data.Message);
									}
								}
								else
								{
									alert(data.Message);
								}
								
							
							}
							,error:function(xhr, textStatus, error){
							      console.log("xhr.statusText:"+xhr.statusText);
							      console.log("xhr.responseText:"+xhr.responseText);
							      console.log("textStatus:"+textStatus);
							      console.log("error:"+error);
								  }
						});
				
			});
	<%if footer_type="newsletter-container" then %>
	$("#submitnewsletter").click(function(e) {
				e.preventDefault();
				var email=$("#emailnewsletter").val();
				if (email==""){
					console.log("Email mancante");
					return true;
				}
				$.ajax({
							url     : "searcher.asp",
							type    : "post",
							dataType: 'json',
							data	: "oper=addnewsletter&email=" + email,
							success: function(data){
								console.log(data);
								if (data.success==true)
								{
									$("#register-newsletter").html("Email per conferma attivazione newsletter inviata");
								}
								else
								{
									$("#register-newsletter").append("<p>"+data.Message+"</p>");

								}
								
							
							}
							,error:function(xhr, textStatus, error){
							      console.log("xhr.statusText:"+xhr.statusText);
							      console.log("xhr.responseText:"+xhr.responseText);
							      console.log("textStatus:"+textStatus);
							      console.log("error:"+error);
								  }
						});
				
			});
	<%end if %>
			
	    $.cookieBar({
		    fixed: true,
		    bottom: true,
		    message: "Per offrirti il miglior servizio possibile questo sito utilizza cookies. Continuando la navigazione nel sito autorizzi l'uso dei cookies. ",
		    acceptText: 'Accetta',
		    policyButton: true,
		    policyText: 'Ulteriori informazioni',
		  policyURL: '/cookies.asp'
		});

    });
    </script>
    	
    	<script>
	    	

	    	
	    	</script>

    <%if false then %>
    <!-- Begin Cookie Consent plugin by Silktide - http://silktide.com/cookieconsent -->
<script type="text/javascript">
    window.cookieconsent_options = {"message":"Per offrirti il miglior servizio possibile questo sito utilizza cookies. Continuando la navigazione nel sito autorizzi l'uso dei cookies.","dismiss":"Accetta","learnMore":"Ulteriori informazioni","link":"cookies.asp","theme":"light-bottom"};
</script>
<%end if%>



	