<!--#include virtual="/setup.asp" -->
<%

if request("send-contact")<>"" then
	txt_errore=""
	'Controllo validazione form eseguita
	if request.form("form_valido")="xx" then
		txt_errore="Errore"
		add2log "Bloccato modulo manuali per validazione form bypassata"&vbcrlf&queryeform()&"<br>IP:"&Request.ServerVariables("REMOTE_ADDR"),3
	end if
	
	if txt_errore="" then
		if request.form("passmanuali")="" or request.form("passmanuali")<>pw_manuali then
			txt_errore="Password errata"
		end if
	end if
	if txt_errore="" then
		session("passmanuali")=true
	end if
end if
if session("idadmin")<>"" then
	css_FileUpload=true
	js_toastr=true
end if
meta_title="listini"
%>
<!--#include virtual="/config/header_inc.asp" -->

        <section id="content">
        	<div id="breadcrumb-container">
        		<div class="container">
					<ul class="breadcrumb">
						<li><a href="/">Home</a></li>
						<li class="active"><%=traduci("listini")%></li>
					</ul>
        		</div>
        	</div>
        	<div class="container">
        		<div class="row">
        			<div class="col-md-12">
						<header class="content-title">
							<h1 class="title"><%=traduci("listini")%></h1>
							<%if session("idadmin")="" then %>
							<p class="title-desc"><%=traduci("listinitxt")%></p>
							<%else %>
							<p><%=traduci("downtxt")%></p>
							<%end if %>
						</header>
						<%if session("idadmin")<>"" then%>
						<ul class="product-details-list">
						<%
						sql="select * from files where cosa=7"
						set rs_files=conn.execute(sql)
						do until rs_files.EOF%>
							<li id="id_file_<%=rs_files("idfiles")%>">
								<a href="<%=rs_files("path")&rs_files("filename")%>" target="blank"><%=rs_files("descrizione")%></a>
								<%if session("idadmin")<>"" then%>
								<span class="file">
								<a class="edit_img btn btn-info btn-xs" title="Modifica nome"><i class="fa fa-pencil"></i></a>	
								<a class="delete_img btn btn-danger btn-xs" title="Elimina"><i class="fa fa-trash-o"></i></a>
								</span><%end if%>
							</li>
							
							
							
							<%
							rs_files.movenext
							loop
							rs_files.close
							%>
						
						</ul>
						<%end if%>
        				
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
				<%if session("idadmin")<>"" then%>
				<div class="row">
        			<div class="col-md-12">

        		    <!-- The fileinput-button span is used to style the file input field as button -->
				    <span class="btn btn-success fileinput-button">
				        <i class="glyphicon glyphicon-plus"></i>
				        <span>Aggiungi files...</span>
				        <!-- The file input field used as target for the file upload widget -->
				        <input id="fileupload" type="file" name="files[]" multiple>
				    </span>
				    <br>
				    <br>
				    <!-- The global progress bar -->
				    <div id="progress" class="progress">
				        <div class="progress-bar progress-bar-success"></div>
				    </div>
				    <!-- The container for the uploaded files -->
				    <div id="files" class="files"></div>
        			</div><!-- End .col-md-12 -->
        		</div><!-- End .row -->
				<%end if%>
			</div><!-- End .container -->
        
        </section><!-- End #content -->
	<!--#include virtual="/footer_inc.asp" -->
       
    <!-- END -->
	<!--#include virtual="/script_inc.asp" -->

	<script src="js/Jqueryvalidate/jquery.validate.min.js"></script>
	<script src="js/Jqueryvalidate/localization/messages_it.min.js"></script>
	<script>
	$(document).ready(function() {
		$("#contact-form").validate({
			rules: {
				passmanuali: {required:true,minlength:3}
			},
			errorElement: "span",
			errorClass: "help-block",
			highlight: function (element, errorClass, validClass) {
				$(element).closest('.input-group').addClass('has-error');
			},
			unhighlight: function (element, errorClass, validClass) {
				$(element).closest('.input-group').removeClass('has-error');
			},
			errorPlacement: function (error, element) {
				if (element.parent('.input-group').length || element.prop('type') === 'checkbox' || element.prop('type') === 'radio') {
					error.insertAfter(element.parent());
				} else {
					error.insertAfter(element);
				}
			},
			submitHandler: function(form){
				$('#form_valido').val('si');
				form.submit();
			}
		});

	});
	</script>
	<%if session("idadmin") then%>

	<div class="modal fade" id="modal-form" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
	  <div class="modal-dialog">
		<div class="modal-content">
		  <div class="modal-header">
			<button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">&times;</span><span class="sr-only">Close</span></button>
			<h4 class="modal-title" id="exampleModalLabel">Modifica descrizione file</h4>
		  </div>
		  <div class="modal-body">
			<form role="form">
			  <div class="form-group">
				<label for="recipient-name" class="control-label">Descrizione file</label>
				<textarea name="descrizione_file" id="descrizione_file" value="" class="form-control"  ></textarea>
			  <input type="hidden" name="dialog_idfile" id="dialog_idfile" value="" >
			  </div>
			</form>
		  </div>
		  <div class="modal-footer">
			<button type="button" class="btn btn-primary" id="modifica">Salva modifiche</button>
			<button type="button" class="btn btn-default" data-dismiss="modal">Chiudi</button>
		  </div>
		</div>
	  </div>
	</div>
    <div class="modal fade" id="confirm-delete" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
            
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                    <h4 class="modal-title" id="myModalLabel">Conferma eliminazione</h4>
                </div>
            
                <div class="modal-body">
                    <p>Sei sicuro di voler procedere?</p>
                    <p>La procedura è irreversibile.</p>
                    <p class="debug-url"></p>
                </div>
                
                <div class="modal-footer">
                    <a href="#" class="btn btn-danger danger" id="delete">Elimina</a>
					<button type="button" class="btn btn-default" data-dismiss="modal">Annulla</button>                    
                </div>
            </div>
        </div>
    </div>

	<div id="dialog-confirm" style="display:none;">
	  Vuoi eliminare l'immagine?
	</div>
	<script>
	</script>

	
	<!-- The jQuery UI widget factory, can be omitted if jQuery UI is already included -->
	<script src="jquery/JQuery-File-Upload/vendor/jquery.ui.widget.js"></script>
	<!-- The Load Image plugin is included for the preview images and image resizing functionality -->
	<script src="//blueimp.github.io/JavaScript-Load-Image/js/load-image.all.min.js"></script>
	<!-- The Canvas to Blob plugin is included for image resizing functionality -->
	<script src="//blueimp.github.io/JavaScript-Canvas-to-Blob/js/canvas-to-blob.min.js"></script>
	<%if false then%>
	<!-- Bootstrap JS is not required, but included for the responsive demo navigation -->
	<script src="//netdna.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>
	<%end if%>
	<!-- The Iframe Transport is required for browsers without support for XHR file uploads -->
	<script src="jquery/JQuery-File-Upload/jquery.iframe-transport.js"></script>
	<!-- The basic File Upload plugin -->
	<script src="jquery/JQuery-File-Upload/jquery.fileupload.js"></script>
	<!-- The File Upload processing plugin -->
	<script src="jquery/JQuery-File-Upload/jquery.fileupload-process.js"></script>
	<!-- The File Upload image preview & resize plugin -->
	<script src="jquery/JQuery-File-Upload/jquery.fileupload-image.js"></script>
	<!-- The File Upload audio preview plugin -->
	<script src="jquery/JQuery-File-Upload/jquery.fileupload-audio.js"></script>
	<!-- The File Upload video preview plugin -->
	<script src="jquery/JQuery-File-Upload/jquery.fileupload-video.js"></script>
	<!-- The File Upload validation plugin -->
	<script src="jquery/JQuery-File-Upload/jquery.fileupload-validate.js"></script>
	<script>
	/*jslint unparam: true, regexp: true */
	/*global window, $ */
	
	$(function () {
		var id_file;
		$.ajaxSetup({
			cache: false
		});

	
	
	
	    'use strict';
	    // Change this to the location of your server-side upload handler:
	    //var url = window.location.hostname === 'blueimp.github.io' ?
	                //'//jquery-file-upload.appspot.com/' : 'server/php/',
	        uploadButton = $('<button/>')
	            .addClass('btn btn-primary')
	            .prop('disabled', true)
	            .text('Processing...')
	            .on('click', function () {
	                var $this = $(this),
	                    data = $this.data();
	                $this
	                    .off('click')
	                    .text('Abort')
	                    .on('click', function () {
	                        $this.remove();
	                        data.abort();
	                    });
	                data.submit().always(function () {
	                    $this.remove();
	                });
	            });
	    $('#fileupload').fileupload({
			url:'/sub_upload_new.asp?tipo_allegato=7&id_tipo_allegato=0&tipo_response=new',
	       dataType: 'json',
	        autoUpload: false,
	        acceptFileTypes: /(\.|\/)(gif|jpe?g|png|pdf)$/i,
	        maxfilesize: 15000000, // 5 MB
	        // Enable image resizing, except for Android and Opera,
	        // which actually support image resizing, but fail to
	        // send Blob objects via XHR requests:
	        disableImageResize: /Android(?!.*Chrome)|Opera/
	            .test(window.navigator.userAgent),
	        previewMaxWidth: 100,
	        previewMaxHeight: 100,
	        previewCrop: true
	    }).on('fileuploadadd', function (e, data) {
	        data.context = $('<div/>').appendTo('#files');
	        $.each(data.files, function (index, file) {
	            var node = $('<p/>')
	                    .append($('<span/>').text(file.name));
	            if (!index) {
	                node
	                    .append('<br>')
	                    .append(uploadButton.clone(true).data(data));
	            }
	            node.appendTo(data.context);
	        });
	    }).on('fileuploadprocessalways', function (e, data) {
	        var index = data.index,
	            file = data.files[index],
	            node = $(data.context.children()[index]);
	        if (file.preview) {
	            node
	                .prepend('<br>')
	                .prepend(file.preview);
	        }
	        if (file.error) {
	            node
	                .append('<br>')
	                .append($('<span class="text-danger"/>').text(file.error));
	        }
	        if (index + 1 === data.files.length) {
	            data.context.find('button')
	                .text('Upload')
	                .prop('disabled', !!data.files.error);
	        }
	    }).on('fileuploadprogressall', function (e, data) {
	        var progress = parseInt(data.loaded / data.total * 100, 10);
	        $('#progress .progress-bar').css(
	            'width',
	            progress + '%'
	        );
	    }).on('fileuploaddone', function (e, data) {
	        $.each(data.result.files, function (index, file) {
	        	id_file=file.idfile;
	        	console.log("idfile:"+id_file);
	        	$("#descrizione_file").val(file.name);
	        	
	        	$( "#modal-form" ).modal("show");
	            if (file.url) {
	                var link = $('<a>')
	                    .attr('target', '_blank')
	                    .prop('href', file.url);
	                $(data.context.children()[index])
	                    .wrap(link);
	            } else if (file.error) {
	                var error = $('<span class="text-danger"/>').text(file.error);
	                $(data.context.children()[index])
	                    .append('<br>')
	                    .append(error);
	            }
	        });
	    }).on('fileuploadfail', function (e, data) {
	        $.each(data.files, function (index) {
	            var error = $('<span class="text-danger"/>').text('Errore nel caricamento.');
	            $(data.context.children()[index])
	                .append('<br>')
	                .append(error);
	        });
	    }).prop('disabled', !$.support.fileInput)
	        .parent().addClass($.support.fileInput ? undefined : 'disabled');

//Parte mia

		var form;
		$(".edit_img").on("click", function (e) {
		 	id_file=$(this).closest("li").attr("id").substring(8);
			console.log("idfile:"+id_file);
			$.get("ajax_function.asp?oper=load_descrizione_img&id_file=" + id_file, function (result) {
			//alert("qui");
			if (result.success==true)
			{
				//toastr.options = {"positionClass": "toast-bottom-right"};
				//toastr.success(result.Message);
				
				$("#descrizione_file").val(result.descrizione_img);
				$("#modal-form").modal( "show" );
			}
			else
			{
				toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
				toastr.error(result.Message);
			}
			});	
		});

	$("#modifica").on("click", function (e) {
		var descrizione_file=$( "#descrizione_file" ).val();
		var dati=encodeURI("oper=save_descrizione_img&descrizione_img="+descrizione_file+"&id_file="+id_file);
		console.log("salvo descrizione idfile:"+id_file);
		$.ajax({
			url     : "ajax_function.asp",
			type    : "post",
			dataType: 'json',
			data	: dati,
			success: function(data){
				//alert(data.Message);
				
				toastr.options = {"positionClass": "toast-bottom-right"};
				toastr.success(data.Message);
				$("#modal-form").modal( "hide" );
			}
			,error:function(xhr, textStatus, error){
			      console.log("xhr.statusText:"+xhr.statusText);
			      console.log("xhr.responseText:"+xhr.responseText);
			      console.log("textStatus:"+textStatus);
			      console.log("error:"+error);
				  }
		});

	
	
	});

			
	$(".delete_img").on("click", function (e) {
		id_file=$(this).closest("li").attr("id").substring(8);
		console.log ("idfile:"+id_file);
		$("#confirm-delete").modal("show");
	});
				
	$("#delete").on("click", function (e) {
	        
        $.ajaxSetup({ cache: false });

		$.get("ajax_function.asp?oper=elimina_file&id_file="+id_file, function(result){
			//alert("qui");
			if (result.success==true)
			{
				toastr.options = {"positionClass": "toast-bottom-right"};
				toastr.success(result.Message);
				$("#id_file_"+id_file).remove();
				 $("#confirm-delete").modal("hide");
			}
			else
			{
				toastr.options = {"positionClass": "toast-bottom-right","timeOut": "0","extendedTimeOut": "0", "closeButton": true};
				toastr.error(result.Message);
			}
		});	
    });
});
</script>
	
	
	<%end if%>
	

    </body>
</html>