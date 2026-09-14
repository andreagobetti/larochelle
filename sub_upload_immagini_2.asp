    <div id="img_upload">
    <!-- The fileupload-buttonbar contains buttons to add/delete files and start/cancel the upload -->
    <div class="fileupload-buttonbar">
        <div class="fileupload-buttons">
            <!-- The fileinput-button span is used to style the file input field as button -->
            <span class="fileinput-button">
                <span>Aggiungi immagini...</span>
                <input type="file" name="files[]" multiple>
            </span>
            <button type="submit" class="start">Carica tutti</button>
        </div>
        <!-- The global progress state -->
        <div class="fileupload-progress fade" style="display:none">
            <!-- The global progress bar -->
            <div class="progress" role="progressbar" aria-valuemin="0" aria-valuemax="100"></div>
            <!-- The extended global progress state -->
            <div class="progress-extended">&nbsp;</div>
        </div>
    </div>
    <!-- The table listing the files available for upload/download -->
    <table role="presentation" width="100%"><tbody class="files"></tbody></table>
    </div>
<br>

<!-- The template to display files available for upload -->
<script id="template-upload" type="text/x-tmpl">
{% for (var i=0, file; file=o.files[i]; i++) { %}
    <tr class="template-upload fade">
        <td>
            <span class="preview"></span>
        </td>
        <td>
            <p class="name">{%=file.name%}</p>
            <strong class="error"></strong>
        </td>
        <td>
            <p class="size">Processing...</p>
            <div class="progress"></div>
        </td>
        <td>
            {% if (!i && !o.options.autoUpload) { %}
                <button class="start" disabled>Carica singolo</button>
            {% } %}
            {% if (!i) { %}
                <button class="cancel">Cancel</button>
            {% } %}
        </td>
    </tr>
{% } %}
</script>
<!-- The template to display files available for download -->
<script id="template-download" type="text/x-tmpl">
{% for (var i=0, file; file=o.files[i]; i++) { %}
    <tr class="template-download fade">
        <td>
            <span class="preview">
                {% if (file.thumbnailUrl) { %}
                    <a href="{%=file.url%}" title="{%=file.name%}" download="{%=file.name%}" data-gallery><img src="{%=file.thumbnailUrl%}"></a>
                {% } %}
            </span>
        </td>
        <td>
            <p class="name">
                <a href="{%=file.url%}" title="{%=file.name%}" download="{%=file.name%}" {%=file.thumbnailUrl?'data-gallery':''%}>{%=file.name%}</a>
            </p>
            {% if (file.error) { %}
                <div><span class="error red">Errore</span> {%=file.error%}</div>
            {% } %}
        </td>
        <td>
            <span class="size">{%=o.formatfilesize(file.size)%}</span>
        </td>
        <td>
            <button class="delete" data-type="{%=file.deleteType%}" data-url="{%=file.deleteUrl%}"{% if (file.deleteWithCredentials) { %} data-xhr-fields='{"withCredentials":true}'{% } %}>Delete</button>
            <input type="checkbox" name="delete" value="1" class="toggle">
        </td>
    </tr>
{% } %}
</script>
<!-- The Templates plugin is included to render the upload/download listings -->
<script src="//blueimp.github.io/JavaScript-Templates/js/tmpl.min.js"></script>
<!-- The Load Image plugin is included for the preview images and image resizing functionality -->
<script src="//blueimp.github.io/JavaScript-Load-Image/js/load-image.all.min.js"></script>
<!-- The Canvas to Blob plugin is included for image resizing functionality -->
<script src="//blueimp.github.io/JavaScript-Canvas-to-Blob/js/canvas-to-blob.min.js"></script>
<!-- blueimp Gallery script -->
<script src="//blueimp.github.io/Gallery/js/jquery.blueimp-gallery.min.js"></script>
<!-- The Iframe Transport is required for browsers without support for XHR file uploads -->
<script src="Jquery/JQuery-File-Upload/jquery.iframe-transport.js"></script>
<!-- The basic File Upload plugin -->
<script src="Jquery/JQuery-File-Upload/jquery.fileupload.js"></script>
<!-- The File Upload processing plugin -->
<script src="Jquery/JQuery-File-Upload/jquery.fileupload-process.js"></script>
<!-- The File Upload image preview & resize plugin -->
<script src="Jquery/JQuery-File-Upload/jquery.fileupload-image.js"></script>
<!-- The File Upload audio preview plugin -->
<script src="Jquery/JQuery-File-Upload/jquery.fileupload-audio.js"></script>
<!-- The File Upload video preview plugin -->
<script src="Jquery/JQuery-File-Upload/jquery.fileupload-video.js"></script>
<!-- The File Upload validation plugin -->
<script src="Jquery/JQuery-File-Upload/jquery.fileupload-validate.js"></script>
<!-- The File Upload user interface plugin -->
<script src="Jquery/JQuery-File-Upload/jquery.fileupload-ui.js"></script>
<!-- The File Upload jQuery UI plugin -->
<script src="Jquery/JQuery-File-Upload/jquery.fileupload-jquery-ui.js"></script>
<!-- The main application script -->
<script src="Jquery/JQuery-File-Upload/main.js"></script>
<!--<script src="Jquery/JQuery-File-Upload/main.js"></script>
-->
<script>
$(function () {
	
	
	$('#img_upload').fileupload({
	//uploadTemplateId: 'template-upload',
	//downloadTemplateId: 'template-download',
	previewAsCanvas: false,
	acceptFileTypes:  /(\.|\/)(gif|jpe?g|png)$/i,
	url:'/sub_upload_new.asp?tipo_allegato=<%=tipo_allegato%>&id_tipo_allegato=<%=id_tipo_allegato%>&tipo_response=new'
	}).bind('fileuploaddone', function (e, data) {
		//alert("Immagine caricata");
		$.ajaxSetup({ cache: false });
		$.get("ajax_function.asp?oper=aggiorna_elenco_immagini&tipo_allegato=<%=tipo_allegato%>&id_tipo_allegato=<%=id_tipo_allegato%>&idpro=<%=idpro%>", function(result){
			$("#sub_allegati_img").html(result);
		 });	
		 return true;
	});


	});
</script>
Form: 04_11_2014<br>Upload: sub_upload_new.asp

<!-- The XDomainRequest Transport is included for cross-domain file deletion for IE 8 and IE 9 -->
<!--[if (gte IE 8)&(lt IE 10)]>
<script src="js/cors/jquery.xdr-transport.js"></script>
<![endif]-->

