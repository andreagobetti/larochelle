<%
if tipo_allegato="" then tipo_allegato=request("tipo_allegato")
if id_tipo_allegato="" then id_tipo_allegato=request("tipo_allegato")

%>
<!-- The file upload form used as target for the file upload widget -->
<%
	if tipo_allegato=3 then %>
	<form id="fileupload"  method="POST" enctype="multipart/form-data">
	<%else %>
    <div id="fileupload">
	<%end if %>
    <!-- Redirect browsers with JavaScript disabled to the origin page -->
    <noscript><input type="hidden" name="redirect" value="https://blueimp.github.io/jQuery-File-Upload/"></noscript>
    <!-- The fileupload-buttonbar contains buttons to add/delete files and start/cancel the upload -->
    <div class="fileupload-buttonbar">
        <div class="fileupload-buttons">
            <!-- The fileinput-button span is used to style the file input field as button -->
            <span class="fileinput-button">
                <span>Add files...</span>
                <input type="file" name="files[]" multiple>
            </span>
            <!-- The global file processing state -->
            <span class="fileupload-process"></span>
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
    <table role="presentation"><tbody class="files"></tbody></table>
<%
	if tipo_allegato=3 then %>
    </form>
	<%else %>
    </div>
	<%end if %>
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
                <button class="start" disabled>Start</button>
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
                <div><span class="error">Error</span> {%=file.error%}</div>
            {% } %}
        </td>
        <td>
            <span class="size">{%=o.formatFileSize(file.size)%}</span>
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
<script>
$(function () {
	
	
	$('#fileupload').fileupload({
	//uploadTemplateId: 'template-upload',
	//downloadTemplateId: 'template-download',
	previewAsCanvas: false,
	autoUpload:true,
	//acceptFileTypes:  /(\.|\/)(gif|jpe?g|png)$/i,
	acceptFileTypes: /^application\/(pdf|msword)$|^doc$|^docx$/i,
	url:'/sub_upload_new.asp?tipo_allegato=<%=tipo_allegato%>&id_tipo_allegato=<%=id_tipo_allegato%>&tipo_response=new'
	}).bind('fileuploaddone', function (e, data) {
		//alert("Immagine caricata");
		$.ajaxSetup({ cache: false });
		$.get("ajax_function.asp?oper=aggiorna_elenco_allegati&tipo_allegato=<%=tipo_allegato%>&id_tipo_allegato=<%=id_tipo_allegato%>&idpro=<%=idpro%>", function(result){
			$("#sub_allegati").html(result);
		 });	
		 return true;
	});


	});
</script>
File accettati: PDF DOC DOCX XLS XLSX
<%if utente_andrea then %><br>[soloio]
Form: 30_05_2016<br>Upload: sub_upload_new.asp[/soloio]
<%end if%>

