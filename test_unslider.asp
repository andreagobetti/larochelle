<HTML>
<HEAD>
<style>


.banner {
	position: relative;
/*	width: 100%;
*/	overflow: auto;
	
/*	font-size: 18px;
	line-height: 24px;
*/	text-align: center;
	
	color: rgba(255,255,255,.6);
	text-shadow: 0 0 1px rgba(0,0,0,.05), 0 1px 2px rgba(0,0,0,.3);
	
	background: #fff;
/*	box-shadow: 0 1px 2px rgba(0,0,0,.25);
*/}
	.banner ul {
		list-style: none;
/*		width: 300%;
*/	}
	.banner ul li {
		display: block;
		float: left;
/*		width: 33%;
*/		
		min-height: 350px;
		
/*		-webkit-background-size: 100% 100%;*/
		-moz-background-size: 100% 100%;
		-o-background-size: 100% 100%;
		-ms-background-size: 100% 100%;
/*		background-size: 100% 100%;*/
		
		box-shadow: inset 0 -3px 6px rgba(0,0,0,.1);
	}
	
	.banner h1, .banner h2 {
		font-size: 40px;
		line-height: 52px;
		
		color: #fff;
	}
	
	.banner .btn {
		display: inline-block;
		margin: 25px 0 0;
		padding: 9px 22px 7px;
		clear: both;
		
		color: #fff;
		font-size: 12px;
		font-weight: bold;
		text-transform: uppercase;
		text-decoration: none;
		
		border: 2px solid rgba(255,255,255,.4);
		border-radius: 5px;
	}
		.banner .btn:hover {
			background: rgba(255,255,255,.05);
		}
		.banner .btn:active {
			-webkit-filter: drop-shadow(0 -1px 2px rgba(0,0,0,.5));
			-moz-filter: drop-shadow(0 -1px 2px rgba(0,0,0,.5));
			-ms-filter: drop-shadow(0 -1px 2px rgba(0,0,0,.5));
			-o-filter: drop-shadow(0 -1px 2px rgba(0,0,0,.5));
			filter: drop-shadow(0 -1px 2px rgba(0,0,0,.5));
		}
		
	.banner .btn, .banner .dot {
		-webkit-filter: drop-shadow(0 1px 2px rgba(0,0,0,.3));
		-moz-filter: drop-shadow(0 1px 2px rgba(0,0,0,.3));
		-ms-filter: drop-shadow(0 1px 2px rgba(0,0,0,.3));
		-o-filter: drop-shadow(0 1px 2px rgba(0,0,0,.3));
		filter: drop-shadow(0 1px 2px rgba(0,0,0,.3));
	}
	
	.banner .dots {
		position: absolute;
		left: 0;
		right: 0;
		bottom: 20px;
	}
		.banner .dots li {
			display: inline-block;
			width: 10px;
			height: 10px;
			margin: 0 4px;
			
			text-indent: -999em;
			
			border: 2px solid #fff;
			border-radius: 6px;
			
			cursor: pointer;
			opacity: .4;
			
			-webkit-transition: background .5s, opacity .5s;
			-moz-transition: background .5s, opacity .5s;
			transition: background .5s, opacity .5s;
		}
			.banner .dots li.active {
				background: #fff;
				opacity: 1;
			}
		



		</style>
<script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
<script type="text/javascript" src="Jquery/js/unslider.js"></script>
<script>
$(function() {
	$('.banner').unslider({
		speed: 500,               //  The speed to animate each slide (in milliseconds)
		delay: 3000,              //  The delay between slide animations (in milliseconds)
		keys: true,               //  Enable keyboard (left, right) arrow shortcuts
		dots: true,               //  Display dot navigation
		fluid: true              //  Support responsive design. May break non-responsive designs
	});
});
</script>
</HEAD>
<BODY>
     <div id="slideshow" class="banner" style="width:300px;">
        <ul>
          <li>
            <img src="http://farm6.static.flickr.com/5270/5627221570_afdd85f16a_z.jpg" alt="" title="Light Trails" />
          </li>
          
          <li>
            <img src="http://farm6.static.flickr.com/5146/5627204218_b83b2d25d6_z.jpg" alt="" title="Bokeh" />
          </li>
          
          <li>           
            <img src="http://farm6.static.flickr.com/5181/5626622843_783739c864_z.jpg" alt="" title="Blossoms" />
          </li>
          
          <li>           
            <img src="http://farm6.static.flickr.com/5183/5627213996_915aa49939_z.jpg" alt="" title="Funky Painting" />
          </li>
          
          <li>           
            <img src="http://farm6.static.flickr.com/5182/5626649425_fde8610329_z.jpg" alt="" title="Vintage Chandelier" />
          </li>	                         
        </ul>
      </div>   
      </BODY>
</HTML>