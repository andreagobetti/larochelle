<%
id_tipo_allegato=request("id_tipo_allegato")
tipo_allegato=request("tipo_allegato")
%>
<!DOCTYPE HTML>
<!--
/*
 * jQuery File Upload Plugin jQuery UI Demo 9.0.0
 * https://github.com/blueimp/jQuery-File-Upload
 *
 * Copyright 2013, Sebastian Tschan
 * https://blueimp.net
 *
 * Licensed under the MIT license:
 * http://www.opensource.org/licenses/MIT
 */
-->
<html lang="it">
<head>
<!-- Force latest IE rendering engine or ChromeFrame if installed -->
<!--[if IE]>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<![endif]-->
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<!-- jQuery UI styles -->
<link rel="stylesheet" href="/jquery/css/jquery.ui.all.css"/>
<style>
/* Adjust the jQuery UI widget font-size: */
.ui-widget {
    font-size: 0.75em;
}
</style>
<!-- CSS to style the file input field as button and adjust the Bootstrap progress bars -->
<link rel="stylesheet" href="Jquery/JQuery-File-Upload/css/jquery.fileupload.css">
<link rel="stylesheet" href="Jquery/JQuery-File-Upload/css/jquery.fileupload-ui.css">
<!-- CSS adjustments for browsers with JavaScript disabled -->
<noscript><link rel="stylesheet" href="Jquery/JQuery-File-Upload/css/jquery.fileupload-noscript.css"></noscript>
<noscript><link rel="stylesheet" href="Jquery/JQuery-File-Upload/css/jquery.fileupload-ui-noscript.css"></noscript>
</head>
<body>
pag_upload_allegati_new
<!--<div class="ui-widget-header ui-corner-all titolo_admin">Upload allegati</div>
-->
<!-- The file upload form used as target for the file upload widget -->
<form id="fileupload"  method="POST" enctype="multipart/form-data">
    <!-- The fileupload-buttonbar contains buttons to add/delete files and start/cancel the upload -->
    <div class="fileupload-buttonbar">
        <div class="fileupload-buttons">
            <!-- The fileinput-button span is used to style the file input field as button -->
            <span class="fileinput-button">
                <span>Aggiungi files...</span>
                <input type="file" name="files[]" multiple>
            </span>
            <button type="submit" class="start">Carica files</button>
            <button type="reset" class="cancel">Annulla caricamento</button>
            <button type="button" class="delete">Cancella</button>
            <input type="checkbox" class="toggle">
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
    <!-- The table listing the files available for upload/download -->
    <table role="presentation" width="100%"><tbody class="files"></tbody></table>
        </div>

</form>
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
            {% if (!i) { %}
                <button class="cancel">Rimuovi</button>
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
            <span class="size">{%=o.formatfilesize(file.size)%}</span>
        </td>
        <td>
            <button class="delete" data-type="{%=file.deleteType%}" data-url="{%=file.deleteUrl%}"{% if (file.deleteWithCredentials) { %} data-xhr-fields='{"withCredentials":true}'{% } %}>Delete</button>
            <input type="checkbox" name="delete" value="1" class="toggle">
        </td>
    </tr>
{% } %}
</script>
<script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
<script src="//ajax.googleapis.com/ajax/libs/jqueryui/1.10.3/jquery-ui.min.js"></script>
<!-- The Templates plugin is included to render the upload/download listings -->
<script src="http://blueimp.github.io/JavaScript-Templates/js/tmpl.min.js"></script>
<!-- The Load Image plugin is included for the preview images and image resizing functionality -->
<script src="http://blueimp.github.io/JavaScript-Load-Image/js/load-image.min.js"></script>
<!-- The Canvas to Blob plugin is included for image resizing functionality -->
<script src="http://blueimp.github.io/JavaScript-Canvas-to-Blob/js/canvas-to-blob.min.js"></script>
<!-- blueimp Gallery script -->
<script src="http://blueimp.github.io/Gallery/js/jquery.blueimp-gallery.min.js"></script>
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
<!--<script src="js/main.js"></script>
-->
        <script>
$(function () {


$('#fileupload').fileupload({

acceptFileTypes:  /(\.|\/)(gif|jpe?g|png|pdf|doc?x|xls?x)$/i,
url:'/sub_upload_new.asp?tipo_allegato=<%=tipo_allegato%>&id_tipo_allegato=<%=id_tipo_allegato%>&tipo_response=new'
});


	});
</script>
<!-- The XDomainRequest Transport is included for cross-domain file deletion for IE 8 and IE 9 -->
<!--[if (gte IE 8)&(lt IE 10)]>
<script src="js/cors/jquery.xdr-transport.js"></script>
<![endif]-->
</body> 
</html>
