<%
	set rs_news=conn.execute ("select * from news where tipo='ln' order by ordine")
	if not rs_news.eof then
	
	
%>
  
  
                        <div id="latestnews-slider-container" class="carousel-wrapper">
                            <header class="content-title">
                                <div class="title-bg">
                                    <h2 class="title"><%=traduci("lnews")%></h2>
                                </div><!-- End .title-bg -->
                            </header>
                            <div class="carousel-controls">
                                <div id="latestnews-slider-prev" class="carousel-btn carousel-btn-prev">
                                </div><!-- End .carousel-prev -->
                                <div id="latestnews-slider-next" class="carousel-btn carousel-btn-next carousel-space">
                                </div><!-- End .carousel-next -->
                            </div><!-- End .carousel-controllers -->
                            <div class="sm-margin"></div><!-- space -->
                            <div class="row">
                                <ul class="latestnews-slider owl-carousel">
	                                <%do while not rs_news.eof%>
	                                
                                    <li>
                                    <%if false then%>
                                        <a href="single.html">
                                            <figure class="latestnews-media-container">
                                                <img src="images/blog/post6-small.jpg" alt="lats post" class="img-responsive">
                                            </figure>
                                        </a>
                                    <%end if%>
                                        <h3><a href="single-page.asp?id=<%=rs_news("idnews")%>&chk=<%=chkDataUser(rs_news("data"))%>"><%=rs_news("titolo")%></a></h3>
                                        <p><%=rs_news("testo")%></p>
                                        <div class="latestnews-meta-container">
                                            <div class="pull-left">
                                                <a href="#">Read More...</a>
                                            </div><!-- End .pull-left -->
                                            <div class="pull-right">
                                                <%=formatdatetime(rs_news("data"),2)%>
                                            </div><!-- End .pull-right -->
                                        </div><!-- End .latest-posts-meta-container -->
                                    </li>
                                    <%
	                                    rs_news.moveNext
										loop
	                                %>
	                                
                                    <%if false then%>
                                    <li>
                                        <a href="single.html">
                                            <figure class="latestnews-media-container">
                                                <img src="images/blog/post6-small.jpg" alt="lats post" class="img-responsive">
                                            </figure>
                                        </a>
                                        <h3><a href="single.html">35% Discount on second purchase!</a></h3>
                                        <p>Sed blandit nulla nec nunc ullamcorper tristique. Maurisadipiscing cursus ante ultricies dictum sed lobortis. Nulla iaculis auctor libero, varius adipiscing sapien bibendum vel. In placerat arcu.</p>
                                        <div class="latestnews-meta-container clearfix">
                                            <div class="pull-left">
                                                <a href="#">Read More...</a>
                                            </div><!-- End .pull-left -->
                                            <div class="pull-right">
                                                12.05.2013
                                            </div><!-- End .pull-right -->
                                        </div><!-- End .latest-posts-meta-container -->
                                    </li>
                                    <li>
                                        <a href="single.html">
                                            <figure class="latestnews-media-container">
                                                <img src="images/blog/post7-small.jpg" alt="lats post" class="img-responsive">
                                            </figure>
                                        </a>
                                        <h3><a href="single.html">New Arrivals.</a></h3>
                                        <p>Aiquam mauris libero, suscipit sed ornare ac, suscipit non felis. Fusce sit amet orci justo, a ultrices urna cursus. Suspendisse mauris nibh, tristique eget consectetur a fermentum.</p>
                                        <div class="latestnews-meta-container clearfix">
                                            <div class="pull-left">
                                                <a href="#">Read More...</a>
                                            </div><!-- End .pull-left -->
                                            <div class="pull-right">
                                                10.05.2013
                                            </div><!-- End .pull-right -->
                                        </div><!-- End .latest-posts-meta-container -->
                                    </li>
                                    <%end if%>
                                </ul>
                            </div><!-- End .row -->
                        </div><!-- End .latestnews-slider-container -->
    <%end if%>
