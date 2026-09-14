/* Configurazione 1 estesa */
tinyMCE.init({
	// file_browser_callback: 'openKCFinder',
	forced_root_block:'',
	language : 'it',
	selector:"textarea.mceEditor",
	statusbar: false, 
	toolbar_items_size: 'small',
	auto_focus : false,
	plugins: [ 
	"advlist autolink lists link image charmap print preview hr anchor pagebreak",
	"searchreplace wordcount visualblocks visualchars code fullscreen",
	"insertdatetime media nonbreaking save table contextmenu directionality",
	"paste textcolor  textcolor"
	],
	toolbar1: "undo redo | styleselect fontsizeselect | bold italic forecolor | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent | link ",
	image_advtab: true ,
	
    font_size_style_values: "12px,13px,14px,16px,18px,20px",
    extended_valid_elements: "span[class]"
	
	
	//file_browser_callback: "responsivefilemanager"
	//   open_manager_upload_path: 'uploads/'
});
	
