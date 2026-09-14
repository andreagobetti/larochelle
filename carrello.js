$(document).ready(function()
{
    $('.quantity-input-up').click( function() {
    	var cartID=$(this).attr('id').replace('cart_quantity_up_', '');
    	console.log("cartID:"+cartID);
    	actQuantity=parseInt($("#quantity_"+cartID).val());
    	
    	console.log("actQuantity:"+actQuantity);
    	$("#quantity_"+cartID).val(actQuantity+1);
    	update_cart();
    });
    $('.quantity-input-down').click( function() {
    	var cartID=$(this).attr('id').replace('cart_quantity_down_', '');
    	console.log("cartID:"+cartID);
    	actQuantity=parseInt($("#quantity_"+cartID).val());
    	
    	console.log("actQuantity:"+actQuantity);
    	actQuantity=actQuantity-1;
    	
    	if (actQuantity<1){actQuantity=1}
    	$("#quantity_"+cartID).val(actQuantity);

    	update_cart();
    	
    });
    //$('.cart_quantity_delete').click( function() {
    //    deleteProductFromSummary($(this).attr('id'));
    //    return false;
    //});
    $('.cart_quantity_input').typeWatch({
        highlight: true,
        wait: 600,
        captureLength: 0,
        callback: function(val) {
        console.log("typeWatch");
        update_cart();
            //updateQty(val, true, this.el);
        }
    });
    $("#update_cart").click(function(){
	    $("#form1").submit();
    });

    function update_cart(){
	    $("#update_cart").fadeIn("slow");
    	$(".nascondi").hide();

	    
    }
});
/*
 *	TypeWatch 3.0
 *	
 *	Original by Denny Ferrassoli
 *	Refactored by Charles Christolini
 *	Revised by Kai Schlamp
 *
 *	Examples/Docs: www.dennydotnet.com
 *	
 *  Copyright(c) 2007 Denny Ferrassoli - DennyDotNet.com
 *  Copyright(c) 2008 Charles Christolini - BinaryPie.com
 *  Copyright(c) 2010 Kai Schlamp - medihack.org
 *  
 *  Dual licensed under the MIT and GPL licenses:
 *  http://www.opensource.org/licenses/mit-license.php
 *  http://www.gnu.org/licenses/gpl.html
*/

(function(jQuery) {
	jQuery.fn.typeWatch = function(o){
		// Options
		var options = jQuery.extend({
			wait : 500,
			callback : function() { },
			highlight : true,
			captureLength : 2,
			fireOnEmpty : false,
		}, o);
		
		function checkElement(timer, override) {
			var elTxt = jQuery(timer.el).val();
		
			if ((elTxt.length >= options.captureLength && elTxt.toUpperCase() != timer.text)
			|| (override && elTxt.length >= options.captureLength)
			|| (options.fireOnEmpty && elTxt.length == 0 && timer.text)) {
				timer.text = elTxt.toUpperCase();
				timer.cb(elTxt);
			}
		}
		
		function watchElement(elem) {			
			// Must be text or textarea
			if (elem.type.toUpperCase() == "TEXT" || elem.nodeName.toUpperCase() == "TEXTAREA") {

				// Allocate timer element
				var timer = {
					timer : null, 
					text : jQuery(elem).val().toUpperCase(),
					cb : options.callback, 
					el : elem, 
					wait : options.wait
				};

				// Set focus action (highlight)
				if (options.highlight) {
					jQuery(elem).focus(
						function() {
							this.select();
						});
				}

				// Key watcher / clear and reset the timer
				var startWatch = function(evt) {
					var timerWait = timer.wait;
					var overrideBool = false;

                    // If enter is pressed then diretly execute the callback
					if (evt.keyCode == 13 && this.type.toUpperCase() == "TEXT") {
						timerWait = 1;
						overrideBool = true;
					}
					
					var timerCallbackFx = function()
					{
						checkElement(timer, overrideBool)
					}
					
					// Clear timer					
					clearTimeout(timer.timer);
					timer.timer = setTimeout(timerCallbackFx, timerWait);									
				};
				
				jQuery(elem).keyup(startWatch);
			}
		}
		
		// Watch Each Element
		return this.each(function(index){
			watchElement(this);
		});
	};
})(jQuery);
