<!DOCTYPE html>



 
 
 
	<head>
		<meta charset="utf-8" />
		<title>Orbit Demo</title>
		
		<!-- Attach our CSS -->
	  	
		<!-- Attach necessary JS -->
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
<link rel="stylesheet" type="text/css" href="Jquery/js/orbit/orbit-1.2.3.css">
<script src="Jquery/js/orbit/jquery.orbit-1.2.3.min.js"></script>
		
			<!--[if IE]>
			     <style type="text/css">
			         .timer { display: none !important; }
			         div.caption { background:transparent; filter:progid:DXImageTransform.Microsoft.gradient(startColorstr=#99000000,endColorstr=#99000000);zoom: 1; }
			    </style>
			<![endif]-->
		
		<!-- Run the plugin -->
		<script type="text/javascript">
			$(window).load(function() {
$('#featured').orbit({
     animation: 'fade',                  // fade, horizontal-slide, vertical-slide, horizontal-push
     animationSpeed: 800,                // how fast animtions are
     timer: false, 			 // true or false to have the timer
     advanceSpeed: 4000, 		 // if timer is enabled, time between transitions 
     pauseOnHover: false, 		 // if you hover pauses the slider
     startClockOnMouseOut: false, 	 // if clock should start on MouseOut
     startClockOnMouseOutAfter: 1000, 	 // how long after MouseOut should the timer start again
     directionalNav: false, 		 // manual advancing directional navs
     captions: true, 			 // do you want captions?
     captionAnimation: 'fade', 		 // fade, slideOpen, none
     captionAnimationSpeed: 800, 	 // if so how quickly should they animate in
     bullets: true,			 // true or false to activate the bullet navigation
     bulletThumbs: false,		 // thumbnails for the bullets
     bulletThumbLocation: '',		 // location from this file where thumbs will be
     afterSlideChange: function(){} 	 // empty function 
});			});
		</script>
		
	</head>
	<body>
	
	
	
	
	
<!-- =======================================

THE ACTUAL ORBIT SLIDER CONTENT 

======================================= -->
		<div id="featured"> 
            <img src="http://farm6.static.flickr.com/5270/5627221570_afdd85f16a_z.jpg" alt="" title="Light Trails" />
            <img src="http://farm6.static.flickr.com/5146/5627204218_b83b2d25d6_z.jpg" alt="" title="Bokeh" />
            <img src="http://farm6.static.flickr.com/5181/5626622843_783739c864_z.jpg" alt="" title="Blossoms" />
		</div>
		
		
		
	</body>
</html>




  
