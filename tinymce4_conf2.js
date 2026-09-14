/* Configurazione 1 estesa */
tinyMCE.init({
	// file_browser_callback: 'openKCFinder',
	language : 'it',
	selector:"textarea.mceEditor",
	//inline: true,
	statusbar: false, 
	menubar:false,
	toolbar_items_size: 'small',
	auto_focus : false,
	plugins: [ 
	"advlist autolink lists link image charmap print preview hr anchor pagebreak",
	"searchreplace wordcount visualblocks visualchars code fullscreen",
	"insertdatetime media nonbreaking save table contextmenu directionality",
	"paste textcolor responsivefilemanager youtube textcolor"
	],
	toolbar1: " undo redo |  fontsizeselect | bold italic forecolor | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent ",
	image_advtab: true ,
	
    font_size_style_values: "12px,13px,14px,16px,18px,20px",
	
	external_filemanager_path:"/filemanager123/",
	filemanager_title:"Responsive Filemanager" ,
	external_plugins: { "filemanager" : "/filemanager/plugin.js"},
	setup : function(ed)
            {
                // set the editor font size
                ed.on('init', function() {
				   this.getDoc().body.style.fontSize = '12pt';
                });
            }
	//file_browser_callback: "responsivefilemanager"
	//   open_manager_upload_path: 'uploads/'
});
	
