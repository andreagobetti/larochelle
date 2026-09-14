
<%
	
If Err.Number <> 0 Then
	add2log "Errore in "&questofile&vbcrlf&Err.Description&queryeform(),3
  
End If
	
	
	
	if (IsObject(conn)) then
		if not ( conn is Nothing) then
			conn.Close
			set conn=Nothing
		end if
	end if
	call CheckConnChiusa()

	%>
<!--#include virtual="/config/footer_conf.asp" -->
        <footer id="footer">
	        <%if footer_type="twitterfeed-container" then%>
        	<div id="twitterfeed-container">
        		<div class="container">
        			<div class="row">
        				
        				<div class="twitterfeed col-md-12">
        					<div class="twitter-icon"><i class="fa fa-twitter"></i></div><!-- End .twitter-icon -->
        					<div class="row">
        						<div class="col-md-10 col-sm-10 col-xs-10 col-md-offset-1 col-sm-offset-1 col-xs-offset-1">
        							<div class="twitter_feed flexslider"></div>
        						</div>
        					</div>
        					
        				</div><!-- End .twiitterfeed .col-md-12 -->
        				
        			</div><!-- End .row -->
    			</div><!-- End .container -->
        	</div><!-- End #twitterfeed-container -->
        	<%end if%>
	        <%if footer_type="newsletter-container" then%>
            <div id="newsletter-container">
                <div class="container">
                    <div class="row">
                        <div class="col-md-12 col-sm-12 col-xs-12 clearfix">
                        <h3><%=traduci("newsltxt")%></h3>
                            <form id="register-newsletter" style="color:#000000;">
                            <input type="text" name="emailnewsletter" id="emailnewsletter" required="email" placeholder="<%=traduci("newslpl")%>">
                            <input type="submit" class="btn btn-custom-3" value="<%=traduci("newslbtn")%>" id="submitnewsletter">
                            </form>
                        </div><!--End  .col-md-6 -->
                        
                    </div><!-- End .row -->
                </div><!-- End .container -->
            </div><!-- End #newsletter-container -->
        	<%end if%>
        	<%call inner_footer()%>
        	<%call bottom_footer()%>
        	
        </footer><!-- End #footer -->
    </div><!-- End #wrapper -->
    <a href="#" id="scroll-top" title="Scroll to Top"><i class="fa fa-angle-up"></i></a><!-- End #scroll-top -->
<%
set dizionario=Nothing    
	
%>