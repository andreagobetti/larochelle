<%
	
tipo_allegato=request.form("tipo_allegato")
id_tipo_allegato=request.form("id_tipo_allegato")
%>
<form action="/upload" class="dropzone" drop-zone="" id="file-dropzone"></form>
        
<script type='text/javascript'>
    var tipo_allegato=<%=tipo_allegato%>;
    var id_tipo_allegato=<%=id_tipo_allegato%>;
    
    var jsUrl='https://cdnjs.cloudflare.com/ajax/libs/dropzone/4.3.0/min/dropzone.min.js';
    var cssUrl='https://cdnjs.cloudflare.com/ajax/libs/dropzone/4.3.0/min/dropzone.min.css';
    
    
    var link = document.createElement("link");
	link.href = cssUrl;
	link.type = "text/css";
	link.rel = "stylesheet";
	document.getElementsByTagName("head")[0].appendChild(link);
    //DropzoneJS snippet - js

	$.getScript(jsUrl,function(){
	  // instantiate the uploader
	  $('#file-dropzone').dropzone({ 
	    url: "sub_upload_new.asp?tipo_allegato="+tipo_allegato+"&id_tipo_allegato="+id_tipo_allegato+"&from=dropzone",
	    maxFilesize: 100,
	    paramName: "uploadfile",
	    maxThumbnailFilesize: 5,
	    dictDefaultMessage: '<b>Trascina il file qui per eseguire l\'upload</b><br>oppure fai click',
	    dictFallbackMessage: 'Clicca per aggiungere un file',
	    init: function() {
	      
	      this.on('success', function(file, json) {
		      
		      
		      aggiorna_elenco_files();
		      
	      });
	      
	      this.on('addedfile', function(file) {
	        
	      });
	      this.on("sending", function(file, xhr, formData) {
			  // Will send the filesize along with the file as POST data.
			  formData.append("tipo_allegato", 3);
			  console.log("sending");
			});
	      
	      this.on('drop', function(file) {
	      }); 
	    }
	  });
	});
	
	
	function aggiorna_elenco_files(){
		
        $.ajax({
            url: "sub_allegati_ajax.asp",
            type: "post",
            cache: false,
            data: { tipo_allegato: tipo_allegato,id_tipo_allegato: idord, noscript:'noscript'},
            success: function(data) {
	            $("#elenco_allegati").replaceWith(data);
            },
            error: function(xhr, textStatus, error) {
                toastr.options = {
                    "positionClass": "toast-bottom-right",
                    "timeOut": "0",
                    "extendedTimeOut": "0",
                    "closeButton": true
                };
                toastr.error('Errore nel caricamento');


            }
        });

		
		
	}
	
	
	

$(document).ready(function() {});
        
</script>
        
        
        
