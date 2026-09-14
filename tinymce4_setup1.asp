
	forced_root_block : "",
	force_br_newlines : false,
	force_p_newlines : false, //Evita tags p
	theme: 'modern',
	language : 'it',
	statusbar: false, 
	toolbar_items_size: 'small',
	auto_focus : false,
	style_formats: [
        {title: 'Titolo sezione', block: 'h2', classes: 'title'},
        {title: 'Sottotitolo', block: 'h3', classes: 'sub-title'},
        {title: 'Testo', block: 'p', classes: 'title-desc'}],
	plugins: [ 
	"advlist autolink lists link image charmap print preview hr anchor pagebreak",
	"searchreplace wordcount visualblocks visualchars code fullscreen",
	"insertdatetime media nonbreaking save table contextmenu directionality",
	"paste textcolor responsivefilemanager youtube textcolor codemirror"
	],
	toolbar1: "insertfile undo redo | styleselect fontsizeselect | bold italic forecolor | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent | link image youtube | code",
	image_advtab: true ,
	
    font_size_style_values: "12px,13px,14px,16px,18px,20px",
	
	external_filemanager_path:"/filemanager123/",
	filemanager_title:"Responsive Filemanager" ,
	external_plugins: { "filemanager" : "/filemanager123/plugin.min.js"},
	setup : function(ed) {
                  ed.on('change', function(e) {
	                  tinymce.triggerSave();
			        enableSaveBtn(ed.getElement());
                     //console.log('the event object '+e);
                     //console.log('the editor object '+ed.getElement());
                     //console.log('the content '+ed.getContent());
                  });
            },
    codemirror: {
	    indentOnInit: true, // Whether or not to indent code on init. 
	    path: 'CodeMirror', // Path to CodeMirror distribution
	    config: {           // CodeMirror config object
	       //mode: 'application/x-httpd-php',
	       lineNumbers: true,
	       indentWithTabs:true
	    },
	    jsfiles: [          // Additional JS files to load
	       'mode/clike/clike.js',
	       'mode/php/php.js'
	    ]
	}
	
