<!--#include virtual="/setup.asp" -->
<!--#include file="JSON_latest.asp"-->
<%
if request("oper")="pubblici" then
	
	call mostra_tabella("pubblici")
	call connclose()
	response.End
	
end if
if request("oper")="privati" then
	
	call mostra_tabella("privati")
	call connclose()
	response.End
	
end if
	
%>

<!DOCTYPE html>
<!--[if IE 8]> <html lang="en" class="ie8 no-js"> <![endif]-->
<!--[if IE 9]> <html lang="en" class="ie9 no-js"> <![endif]-->
<!--[if !IE]><!-->
<html lang="en">
    <!--<![endif]-->
    <!-- BEGIN HEAD -->

    <head>
		<!--#include virtual="/met_head.asp" -->
    </head>
    <!-- END HEAD -->

    <body class="page-container-bg-solid page-header-menu-fixed">
        <div class="page-wrapper">
            <div class="page-wrapper-row">
                <div class="page-wrapper-top">
                    <!-- BEGIN HEADER -->
                    <div class="page-header">
	                    
                        <!-- BEGIN HEADER TOP -->
                        <div class="page-header-top">
                            <div class="container">
                                <!-- BEGIN LOGO -->
                                <div class="page-logo">
                                    <a href="index.html">
                                        
                                    </a>
                                </div>
                                <!-- END LOGO -->
	                   		<!-- #include virtual="/met_top_navigation.asp" -->
                            </div>
                        </div>
                        <!-- END HEADER TOP -->
                   		<!--#include virtual="/met_header_menu.asp" -->
 
                        
                    </div>
                    <!-- END HEADER -->
                </div>
            </div>
            <div class="page-wrapper-row full-height">
                <div class="page-wrapper-middle">
                    <!-- BEGIN CONTAINER -->
                    <div class="page-container">
                        <!-- BEGIN CONTENT -->
                        <div class="page-content-wrapper">
                            <!-- BEGIN CONTENT BODY -->
                            <!-- BEGIN PAGE HEAD-->
                            <div class="page-head">
                                <div class="container">
	                                
                                    <!-- BEGIN PAGE TITLE -->
                                    <div class="page-title">
                                        <h1>Appunti</h1>
                                    </div>
                                    <!-- END PAGE TITLE -->
                                        <div class="btn-group btn-theme-panel pull-right">
                                            <a href="javascript:;" class="btn" >
												<i class="fa fa-plus" aria-hidden="true"></i>
                                            </a>
                                        </div>

			                   		<!--#include virtual="/met_page_toolbar.asp" -->
                                </div>
                            </div>
                            <!-- END PAGE HEAD-->
                            <!-- BEGIN PAGE CONTENT BODY -->
                            <div class="page-content">
                                <div class="container">
			                   		<!-- include virtual="/met_breadcrumbs.asp" -->
	                                
                                    
                                    
                                    
                                    <!-- BEGIN PAGE CONTENT INNER -->
                                    <div class="page-content-inner">
	                                    
										<div class="portlet light ">
											<% if false then %>
                                            <div class="portlet-title">
                                                <div class="caption">
                                                    <i class="icon-social-dribbble font-purple-soft"></i>
                                                    <span class="caption-subject font-purple-soft bold uppercase">Default Tabs</span>
                                                </div>
                                                <div class="actions">
                                                    <a class="btn btn-circle btn-icon-only btn-default" href="javascript:;">
                                                        <i class="icon-cloud-upload"></i>
                                                    </a>
                                                    <a class="btn btn-circle btn-icon-only btn-default" href="javascript:;">
                                                        <i class="icon-wrench"></i>
                                                    </a>
                                                    <a class="btn btn-circle btn-icon-only btn-default" href="javascript:;">
                                                        <i class="icon-trash"></i>
                                                    </a>
                                                </div>
                                            </div>
                                            <% end if %>
                                            <div class="portlet-body">
	                                            <div class="tabbable-custom" style="overflow: visible;">
                                                <ul class="nav nav-tabs" id="myTabs">
                                                    <li class="active">
                                                        <a href="#tab_1_1" data-toggle="tab" aria-expanded="true" data-url="<%=questofile%>?oper=pubblici">Pubblici </a>
                                                    </li>
                                                    <li class="">
                                                        <a href="#tab_1_2" data-toggle="tab" aria-expanded="true" data-url="<%=questofile%>?oper=privati">Privati</a>
                                                    </li>
                                                </ul>
                                                <div class="tab-content">
                                                    <div class="tab-pane active" id="tab_1_1">
	                                                    Loading...
                                                    </div>
                                                    <div class="tab-pane" id="tab_1_2">
	                                                    Loading...
                                                    </div>
                                                </div>
	                                            </div>
                                            </div>
                                        </div>

                                    </div>
                                    <!-- END PAGE CONTENT INNER -->
                                </div>
                            </div>
                            <!-- END PAGE CONTENT BODY -->
                            <!-- END CONTENT BODY -->
                        </div>
                        <!-- END CONTENT -->
						<!-- include virtual="/met_quick_sidebar.asp" -->
                        
                    </div>
                    <!-- END CONTAINER -->
                </div>
            </div>
            <div class="page-wrapper-row">
                <div class="page-wrapper-bottom">
                    <!-- BEGIN FOOTER -->
                    
					<!-- include virtual="/met_pre_footer.asp" -->
					<!--#include virtual="/met_footer.asp" -->
                    <!-- END FOOTER -->
                </div>
            </div>
        </div>
        <!-- BEGIN QUICK NAV -->
        <nav class="quick-nav">
            <a class="quick-nav-trigger" href="#0">
                <span aria-hidden="true"></span>
            </a>
            <ul>
                <li>
                    <a href="https://themeforest.net/item/metronic-responsive-admin-dashboard-template/4021469?ref=keenthemes" target="_blank" class="active">
                        <span>Purchase Metronic</span>
                        <i class="icon-basket"></i>
                    </a>
                </li>
                <li>
                    <a href="https://themeforest.net/item/metronic-responsive-admin-dashboard-template/reviews/4021469?ref=keenthemes" target="_blank">
                        <span>Customer Reviews</span>
                        <i class="icon-users"></i>
                    </a>
                </li>
                <li>
                    <a href="http://keenthemes.com/showcast/" target="_blank">
                        <span>Showcase</span>
                        <i class="icon-user"></i>
                    </a>
                </li>
                <li>
                    <a href="http://keenthemes.com/metronic-theme/changelog/" target="_blank">
                        <span>Changelog</span>
                        <i class="icon-graph"></i>
                    </a>
                </li>
            </ul>
            <span aria-hidden="true" class="quick-nav-bg"></span>
        </nav>
        <div class="quick-nav-overlay"></div>
        <!-- END QUICK NAV -->
        
		<!--#include virtual="/met_scripts.asp" -->
		
		<script>
		$(function() {
			$('#myTabs a').click(function (e) {
				e.preventDefault();
			  
				var url = $(this).attr("data-url");
			  	var href = this.hash;
			  	var pane = $(this);
				console.log("data");
				// ajax load from data-url
				$(href).load(url,function(result){      
				    pane.tab('show');
				});
			});
			
			// load first tab content
			var obj=$("#myTabs .active a");
			var url=$(obj).attr("data-url");
			var pane=$(obj).attr("href");
			console.log("url:"+url+" pane:"+pane);
			$(pane).load(url,function(result){      
			    obj.tab('show');
			});
			


		});	
			
		</script>
		
		
    </body>

</html>

<%
sub mostra_tabella(quali)
	
 %>
	
	
														<div class="table-scrollable" style="overflow: visible;">
                                                            <table class="table table-hover">
                                                                <thead>
                                                                    <tr>
                                                                        <th> Data </th>
                                                                        <th> Testo </th>
                                                                        <th> <span class="pull-right">Azioni</span> </th>
                                                                    </tr>
                                                                </thead>
                                                                
                                                                <tbody>
	                                                                
	<%
	sql="select appunti.*, admin.nominativo, appunti_letture.idappunto as lettura from appunti inner join admin on appunti.iduser=admin.iduser left join appunti_letture on appunti.id = appunti_letture.idappunto"
	if quali="pubblici" then
		sql=sql&" where pubblico=1 "
	elseif quali="privati" then
		sql=sql&" where appunti.iduser="&sessioniduser&" and pubblico=0 "
	end if
	sql=sql&" order by fatto, data desc"
	set rs=conn.execute(sql)
	do until rs.EOF
		
	%>                                                                
                                                                    <tr id="id_appunto<%=rs("id")%>">
                                                                        <td> <%=formatdatetime(rs("data"),2)%><% if rs("iduser")<>sessioniduser then response.write "<span style=""font-size: 90%;""> di <b>"&rs("nominativo") &"</b></span>"%><%if isnull(rs("lettura")) and rs("iduser")<>sessioniduser then response.write "&nbsp; <span class=""badge badge-danger"">Nuovo</span>"%>

                                                                        
                                                                        <br>
                                                                         </td>
                                                                        <td> <%=trim(rs("testo"))%> </td>
                                                                        <td class="task-config">
																				<div class="btn-group pull-right">
																					<button type="button" class="btn btn-default">Ho letto</button>
																					<button type="button" class="btn btn-default dropdown-toggle" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
																						<span class="caret"></span>
																						<span class="sr-only">Toggle Dropdown</span>
																					</button>
																					<ul class="dropdown-menu pull-right">
																						<li><a href="#">Fatto</a></li>
																						<li role="separator" class="divider"></li>
																						<li><a href="#">Modifica</a></li>
																					</ul>
																				</div>
                                                                        </td>

																	</tr>
    <%
	    rs.MoveNext
	loop
	    
	    
	    %>                                                              
                                                                    
                                                                    

                                                                </tbody>
                                                            </table>
															</div>
<%
end sub
	

	%>
