<!--#include virtual="/setup.asp" -->
<!--#include virtual="/pag_adm_ordini_inc.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<link rel="shortcut icon" type="image/ico" href="http://www.datatables.net/favicon.ico" />
		
		<title>DataTables example</title>
		<style type="text/css" title="currentStyle">
			@import "Jquery/js/DataTables/demo_page.css";
			@import "Jquery/js/DataTables/demo_table.css";
		</style>
<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
        <script type="text/javascript" src="Jquery/js/DataTables/jquery.dataTables.min.js"></script>
		<script type="text/javascript" charset="utf-8">
			$(document).ready(function() {
				$('#example').dataTable();
			} );
		</script>
	</head>
	<body id="dt_example">
		<div id="container">
			<div class="full_width big">
				DataTables zero configuration example
			</div>
			
			<h1>Preamble</h1>
			<p>DataTables has most features enabled by default, so all you need to do to use it with one of your own tables is to call the construction function (as shown below).</p>
			
			<h1>Live example</h1>
              <table id="example" class="display" cellspacing="0" width="100%">
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Position</th>
                    <th>Office</th>
                    <th>Age</th>
                    <th>Start date</th>
                    <th>Prova</th>
                  </tr>
                </thead>
                <tfoot>
                  <tr>
                    <th>Name</th>
                    <th>Position</th>
                    <th>Office</th>
                    <th>Age</th>
                    <th>Start date</th>
                    <th>Salary</th>
                  </tr>
                </tfoot>
                <tbody>
                  <tr>
                    <td>Tiger Nixon</td>
                    <td>System Architect</td>
                    <td>Edinburgh</td>
                    <td>61</td>
                    <td>2011/04/25</td>
                    <td>€ 320,800</td>
                  </tr>
                  <tr>
                    <td>Garrett Winters</td>
                    <td>Accountant</td>
                    <td>Tokyo</td>
                    <td>63</td>
                    <td>2011/07/25</td>
                    <td>€ 170,750</td>
                  </tr>
                  <tr>
                    <td>Ashton Cox</td>
                    <td>Junior Technical Author</td>
                    <td>San Francisco</td>
                    <td>66</td>
                    <td>2009/01/12</td>
                    <td>€ 86,000</td>
                  </tr>
                  <tr>
                    <td>Cedric Kelly</td>
                    <td>Senior Javascript Developer</td>
                    <td>Edinburgh</td>
                    <td>22</td>
                    <td>2012/03/29</td>
                    <td>€ 433,060</td>
                  </tr>
                  <tr>
                    <td>Airi Satou</td>
                    <td>Accountant</td>
                    <td>Tokyo</td>
                    <td>33</td>
                    <td>2008/11/28</td>
                    <td>€ 162,700</td>
                  </tr>
                  <tr>
                    <td>Brielle Williamson</td>
                    <td>Integration Specialist</td>
                    <td>New York</td>
                    <td>61</td>
                    <td>2012/12/02</td>
                    <td>€ 372,000</td>
                  </tr>
                  <tr>
                    <td>Herrod Chandler</td>
                    <td>Sales Assistant</td>
                    <td>San Francisco</td>
                    <td>59</td>
                    <td>2012/08/06</td>
                    <td>€ 137,500</td>
                  </tr>
                  <tr>
                    <td>Rhona Davidson</td>
                    <td>Integration Specialist</td>
                    <td>Tokyo</td>
                    <td>55</td>
                    <td>2010/10/14</td>
                    <td>€ 327,900</td>
                  </tr>
                  <tr>
                    <td>Colleen Hurst</td>
                    <td>Javascript Developer</td>
                    <td>San Francisco</td>
                    <td>39</td>
                    <td>2009/09/15</td>
                    <td>€ 205,500</td>
                  </tr>
                  <tr>
                    <td>Sonya Frost</td>
                    <td>Software Engineer</td>
                    <td>Edinburgh</td>
                    <td>23</td>
                    <td>2008/12/13</td>
                    <td>€ 103,600</td>
                  </tr>
                  <tr>
                    <td>Jena Gaines</td>
                    <td>Office Manager</td>
                    <td>London</td>
                    <td>30</td>
                    <td>2008/12/19</td>
                    <td>€ 90,560</td>
                  </tr>
                  <tr>
                    <td>Quinn Flynn</td>
                    <td>Support Lead</td>
                    <td>Edinburgh</td>
                    <td>22</td>
                    <td>2013/03/03</td>
                    <td>€ 342,000</td>
                  </tr>
                  <tr>
                    <td>Charde Marshall</td>
                    <td>Regional Director</td>
                    <td>San Francisco</td>
                    <td>36</td>
                    <td>2008/10/16</td>
                    <td>€ 470,600</td>
                  </tr>
                  <tr>
                    <td>Haley Kennedy</td>
                    <td>Senior Marketing Designer</td>
                    <td>London</td>
                    <td>43</td>
                    <td>2012/12/18</td>
                    <td>€ 313,500</td>
                  </tr>
                  <tr>
                    <td>Tatyana Fitzpatrick</td>
                    <td>Regional Director</td>
                    <td>London</td>
                    <td>19</td>
                    <td>2010/03/17</td>
                    <td>€ 385,750</td>
                  </tr>
                  <tr>
                    <td>Michael Silva</td>
                    <td>Marketing Designer</td>
                    <td>London</td>
                    <td>66</td>
                    <td>2012/11/27</td>
                    <td>€ 198,500</td>
                  </tr>
                  <tr>
                    <td>Paul Byrd</td>
                    <td>Chief Financial Officer (CFO)</td>
                    <td>New York</td>
                    <td>64</td>
                    <td>2010/06/09</td>
                    <td>€ 725,000</td>
                  </tr>
                  <tr>
                    <td>Gloria Little</td>
                    <td>Systems Administrator</td>
                    <td>New York</td>
                    <td>59</td>
                    <td>2009/04/10</td>
                    <td>€ 237,500</td>
                  </tr>
                  <tr>
                    <td>Bradley Greer</td>
                    <td>Software Engineer</td>
                    <td>London</td>
                    <td>41</td>
                    <td>2012/10/13</td>
                    <td>€ 132,000</td>
                  </tr>
                  <tr>
                    <td>Dai Rios</td>
                    <td>Personnel Lead</td>
                    <td>Edinburgh</td>
                    <td>35</td>
                    <td>2012/09/26</td>
                    <td>€ 217,500</td>
                  </tr>
                  <tr>
                    <td>Jenette Caldwell</td>
                    <td>Development Lead</td>
                    <td>New York</td>
                    <td>30</td>
                    <td>2011/09/03</td>
                    <td>€ 345,000</td>
                  </tr>
                  <tr>
                    <td>Yuri Berry</td>
                    <td>Chief Marketing Officer (CMO)</td>
                    <td>New York</td>
                    <td>40</td>
                    <td>2009/06/25</td>
                    <td>€ 675,000</td>
                  </tr>
                  <tr>
                    <td>Caesar Vance</td>
                    <td>Pre-Sales Support</td>
                    <td>New York</td>
                    <td>21</td>
                    <td>2011/12/12</td>
                    <td>€ 106,450</td>
                  </tr>
                  <tr>
                    <td>Doris Wilder</td>
                    <td>Sales Assistant</td>
                    <td>Sidney</td>
                    <td>23</td>
                    <td>2010/09/20</td>
                    <td>€ 85,600</td>
                  </tr>
                  <tr>
                    <td>Angelica Ramos</td>
                    <td>Chief Executive Officer (CEO)</td>
                    <td>London</td>
                    <td>47</td>
                    <td>2009/10/09</td>
                    <td>€ 1,200,000</td>
                  </tr>
                  <tr>
                    <td>Gavin Joyce</td>
                    <td>Developer</td>
                    <td>Edinburgh</td>
                    <td>42</td>
                    <td>2010/12/22</td>
                    <td>€ 92,575</td>
                  </tr>
                  <tr>
                    <td>Jennifer Chang</td>
                    <td>Regional Director</td>
                    <td>Singapore</td>
                    <td>28</td>
                    <td>2010/11/14</td>
                    <td>€ 357,650</td>
                  </tr>
                  <tr>
                    <td>Brenden Wagner</td>
                    <td>Software Engineer</td>
                    <td>San Francisco</td>
                    <td>28</td>
                    <td>2011/06/07</td>
                    <td>€ 206,850</td>
                  </tr>
                  <tr>
                    <td>Fiona Green</td>
                    <td>Chief Operating Officer (COO)</td>
                    <td>San Francisco</td>
                    <td>48</td>
                    <td>2010/03/11</td>
                    <td>€ 850,000</td>
                  </tr>
                  <tr>
                    <td>Shou Itou</td>
                    <td>Regional Marketing</td>
                    <td>Tokyo</td>
                    <td>20</td>
                    <td>2011/08/14</td>
                    <td>€ 163,000</td>
                  </tr>
                  <tr>
                    <td>Michelle House</td>
                    <td>Integration Specialist</td>
                    <td>Sidney</td>
                    <td>37</td>
                    <td>2011/06/02</td>
                    <td>€ 95,400</td>
                  </tr>
                  <tr>
                    <td>Suki Burks</td>
                    <td>Developer</td>
                    <td>London</td>
                    <td>53</td>
                    <td>2009/10/22</td>
                    <td>€ 114,500</td>
                  </tr>
                  <tr>
                    <td>Prescott Bartlett</td>
                    <td>Technical Author</td>
                    <td>London</td>
                    <td>27</td>
                    <td>2011/05/07</td>
                    <td>€ 145,000</td>
                  </tr>
                  <tr>
                    <td>Gavin Cortez</td>
                    <td>Team Leader</td>
                    <td>San Francisco</td>
                    <td>22</td>
                    <td>2008/10/26</td>
                    <td>€ 235,500</td>
                  </tr>
                  <tr>
                    <td>Martena Mccray</td>
                    <td>Post-Sales support</td>
                    <td>Edinburgh</td>
                    <td>46</td>
                    <td>2011/03/09</td>
                    <td>€ 324,050</td>
                  </tr>
                  <tr>
                    <td>Unity Butler</td>
                    <td>Marketing Designer</td>
                    <td>San Francisco</td>
                    <td>47</td>
                    <td>2009/12/09</td>
                    <td>€ 85,675</td>
                  </tr>
                  <tr>
                    <td>Howard Hatfield</td>
                    <td>Office Manager</td>
                    <td>San Francisco</td>
                    <td>51</td>
                    <td>2008/12/16</td>
                    <td>€ 164,500</td>
                  </tr>
                  <tr>
                    <td>Hope Fuentes</td>
                    <td>Secretary</td>
                    <td>San Francisco</td>
                    <td>41</td>
                    <td>2010/02/12</td>
                    <td>€ 109,850</td>
                  </tr>
                  <tr>
                    <td>Vivian Harrell</td>
                    <td>Financial Controller</td>
                    <td>San Francisco</td>
                    <td>62</td>
                    <td>2009/02/14</td>
                    <td>€ 452,500</td>
                  </tr>
                  <tr>
                    <td>Timothy Mooney</td>
                    <td>Office Manager</td>
                    <td>London</td>
                    <td>37</td>
                    <td>2008/12/11</td>
                    <td>€ 136,200</td>
                  </tr>
                  <tr>
                    <td>Jackson Bradshaw</td>
                    <td>Director</td>
                    <td>New York</td>
                    <td>65</td>
                    <td>2008/09/26</td>
                    <td>€ 645,750</td>
                  </tr>
                  <tr>
                    <td>Olivia Liang</td>
                    <td>Support Engineer</td>
                    <td>Singapore</td>
                    <td>64</td>
                    <td>2011/02/03</td>
                    <td>€ 234,500</td>
                  </tr>
                  <tr>
                    <td>Bruno Nash</td>
                    <td>Software Engineer</td>
                    <td>London</td>
                    <td>38</td>
                    <td>2011/05/03</td>
                    <td>€ 163,500</td>
                  </tr>
                  <tr>
                    <td>Sakura Yamamoto</td>
                    <td>Support Engineer</td>
                    <td>Tokyo</td>
                    <td>37</td>
                    <td>2009/08/19</td>
                    <td>€ 139,575</td>
                  </tr>
                  <tr>
                    <td>Thor Walton</td>
                    <td>Developer</td>
                    <td>New York</td>
                    <td>61</td>
                    <td>2013/08/11</td>
                    <td>€ 98,540</td>
                  </tr>
                  <tr>
                    <td>Finn Camacho</td>
                    <td>Support Engineer</td>
                    <td>San Francisco</td>
                    <td>47</td>
                    <td>2009/07/07</td>
                    <td>€ 87,500</td>
                  </tr>
                  <tr>
                    <td>Serge Baldwin</td>
                    <td>Data Coordinator</td>
                    <td>Singapore</td>
                    <td>64</td>
                    <td>2012/04/09</td>
                    <td>€ 138,575</td>
                  </tr>
                  <tr>
                    <td>Zenaida Frank</td>
                    <td>Software Engineer</td>
                    <td>New York</td>
                    <td>63</td>
                    <td>2010/01/04</td>
                    <td>€ 125,250</td>
                  </tr>
                  <tr>
                    <td>Zorita Serrano</td>
                    <td>Software Engineer</td>
                    <td>San Francisco</td>
                    <td>56</td>
                    <td>2012/06/01</td>
                    <td>€ 115,000</td>
                  </tr>
                  <tr>
                    <td>Jennifer Acosta</td>
                    <td>Junior Javascript Developer</td>
                    <td>Edinburgh</td>
                    <td>43</td>
                    <td>2013/02/01</td>
                    <td>€ 75,650</td>
                  </tr>
                  <tr>
                    <td>Cara Stevens</td>
                    <td>Sales Assistant</td>
                    <td>New York</td>
                    <td>46</td>
                    <td>2011/12/06</td>
                    <td>€ 145,600</td>
                  </tr>
                  <tr>
                    <td>Hermione Butler</td>
                    <td>Regional Director</td>
                    <td>London</td>
                    <td>47</td>
                    <td>2011/03/21</td>
                    <td>€ 356,250</td>
                  </tr>
                  <tr>
                    <td>Lael Greer</td>
                    <td>Systems Administrator</td>
                    <td>London</td>
                    <td>21</td>
                    <td>2009/02/27</td>
                    <td>€ 103,500</td>
                  </tr>
                  <tr>
                    <td>Jonas Alexander</td>
                    <td>Developer</td>
                    <td>San Francisco</td>
                    <td>30</td>
                    <td>2010/07/14</td>
                    <td>€  86,500</td>
                  </tr>
                  <tr>
                    <td>Shad Decker</td>
                    <td>Regional Director</td>
                    <td>Edinburgh</td>
                    <td>51</td>
                    <td>2008/11/13</td>
                    <td>€  183,000</td>
                  </tr>
                  <tr>
                    <td>Michael Bruce</td>
                    <td>Javascript Developer</td>
                    <td>Singapore</td>
                    <td>29</td>
                    <td>2011/06/27</td>
                    <td>€  183,000</td>
                  </tr>
                  <tr>
                    <td>Donna Snider</td>
                    <td>Customer Support</td>
                    <td>New York</td>
                    <td>27</td>
                    <td>2011/01/25</td>
                    <td>€  112,000</td>
                  </tr>
                </tbody>
              </table>
			<div class="spacer"></div>
			
			
			<h1>Initialisation code</h1>
			<pre class="brush: js;">€ (document).ready(function() {
	€ ('#example').dataTable();
} );</pre>
			<style type="text/css">
				@import "../examples_support/syntax/css/shCore.css";
			</style>
			<script type="text/javascript" language="javascript" src="../examples_support/syntax/js/shCore.js"></script>
			
			
			<h1>Other examples</h1>
			<div class="demo_links">
				<h2>Basic initialisation</h2>
				<ul>
					<li><a href="../basic_init/zero_config.html">Zero configuration</a></li>
					<li><a href="../basic_init/filter_only.html">Feature enablement</a></li>
					<li><a href="../basic_init/table_sorting.html">Sorting data</a></li>
					<li><a href="../basic_init/multi_col_sort.html">Multi-column sorting</a></li>
					<li><a href="../basic_init/multiple_tables.html">Multiple tables</a></li>
					<li><a href="../basic_init/hidden_columns.html">Hidden columns</a></li>
					<li><a href="../basic_init/complex_header.html">Complex headers - grouping with colspan</a></li>
					<li><a href="../basic_init/dom.html">DOM positioning</a></li>
					<li><a href="../basic_init/flexible_width.html">Flexible table width</a></li>
					<li><a href="../basic_init/state_save.html">State saving</a></li>
					<li><a href="../basic_init/alt_pagination.html">Alternative pagination styles</a></li>
					<li>Scrolling: <br>
						<a href="../basic_init/scroll_x.html">Horizontal</a> / 
						<a href="../basic_init/scroll_y.html">Vertical</a> / 
						<a href="../basic_init/scroll_xy.html">Both</a> / 
						<a href="../basic_init/scroll_y_theme.html">Themed</a> / 
						<a href="../basic_init/scroll_y_infinite.html">Infinite</a>
					</li>
					<li><a href="../basic_init/language.html">Change language information (internationalisation)</a></li>
					<li><a href="../basic_init/themes.html">ThemeRoller themes (Smoothness)</a></li>
				</ul>
				
				<h2>Advanced initialisation</h2>
				<ul>
					<li>Events: <br>
						<a href="../advanced_init/events_live.html">Live events</a> / 
						<a href="../advanced_init/events_pre_init.html">Pre-init</a> / 
						<a href="../advanced_init/events_post_init.html">Post-init</a>
					</li>
					<li><a href="../advanced_init/column_render.html">Column rendering</a></li>
					<li><a href="../advanced_init/html_sort.html">Sorting without HTML tags</a></li>
					<li><a href="../advanced_init/dom_multiple_elements.html">Multiple table controls (sDom)</a></li>
					<li><a href="../advanced_init/length_menu.html">Defining length menu options</a></li>
					<li><a href="../advanced_init/complex_header.html">Complex headers and hidden columns</a></li>
					<li><a href="../advanced_init/dom_toolbar.html">Custom toolbar (element) around table</a></li>
					<li><a href="../advanced_init/highlight.html">Row highlighting with CSS</a></li>
					<li><a href="../advanced_init/row_grouping.html">Row grouping</a></li>
					<li><a href="../advanced_init/row_callback.html">Row callback</a></li>
					<li><a href="../advanced_init/footer_callback.html">Footer callback</a></li>
					<li><a href="../advanced_init/sorting_control.html">Control sorting direction of columns</a></li>
					<li><a href="../advanced_init/language_file.html">Change language information from a file (internationalisation)</a></li>
					<li><a href="../advanced_init/defaults.html">Setting defaults</a></li>
					<li><a href="../advanced_init/localstorage.html">State saving with localStorage</a></li>
					<li><a href="../advanced_init/dt_events.html">Custom events</a></li>
				</ul>
				
				<h2>API</h2>
				<ul>
					<li><a href="../api/add_row.html">Dynamically add a new row</a></li>
					<li><a href="../api/multi_filter.html">Individual column filtering (using "input" elements)</a></li>
					<li><a href="../api/multi_filter_select.html">Individual column filtering (using "select" elements)</a></li>
					<li><a href="../api/highlight.html">Highlight rows and columns</a></li>
					<li><a href="../api/row_details.html">Show and hide details about a particular record</a></li>
					<li><a href="../api/select_row.html">User selectable rows (multiple rows)</a></li>
					<li><a href="../api/select_single_row.html">User selectable rows (single row) and delete rows</a></li>
					<li><a href="../api/editable.html">Editable rows (with jEditable)</a></li>
					<li><a href="../api/form.html">Submit form with elements in table</a></li>
					<li><a href="../api/counter_column.html">Index column (static number column)</a></li>
					<li><a href="../api/show_hide.html">Show and hide columns dynamically</a></li>
					<li><a href="../api/api_in_init.html">API function use in initialisation object (callback)</a></li>
					<li><a href="../api/tabs_and_scrolling.html">DataTables scrolling and tabs</a></li>
					<li><a href="../api/regex.html">Regular expression filtering</a></li>
				</ul>
			</div>
			
			<div class="demo_links">
				<h2>Data sources</h2>
				<ul>
					<li><a href="../data_sources/dom.html">DOM</a></li>
					<li><a href="../data_sources/js_array.html">Javascript array</a></li>
					<li><a href="../data_sources/ajax.html">Ajax source</a></li>
					<li><a href="../data_sources/server_side.html">Server side processing</a></li>
				</ul>
				
				<h2>Server-side processing</h2>
				<ul>
					<li><a href="../server_side/server_side.html">Obtain server-side data</a></li>
					<li><a href="../server_side/custom_vars.html">Add extra HTTP variables</a></li>
					<li><a href="../server_side/post.html">Use HTTP POST</a></li>
					<li><a href="../server_side/ids.html">Automatic addition of IDs and classes to rows</a></li>
					<li><a href="../server_side/object_data.html">Reading table data from objects</a></li>
					<li><a href="../server_side/row_details.html">Show and hide details about a particular record</a></li>
					<li><a href="../server_side/select_rows.html">User selectable rows (multiple rows)</a></li>
					<li><a href="../server_side/jsonp.html">JSONP for a cross domain data source</a></li>
					<li><a href="../server_side/editable.html">jEditable integration with DataTables</a></li>
					<li><a href="../server_side/defer_loading.html">Deferred loading of Ajax data</a></li>
					<li><a href="../server_side/pipeline.html">Pipelining data (reduce Ajax calls for paging)</a></li>
				</ul>
				
				<h2>Ajax data source</h2>
				<ul>
					<li><a href="../ajax/ajax.html">Ajax sourced data (array of arrays)</a></li>
					<li><a href="../ajax/objects.html">Ajax sourced data (array of objects)</a></li>
					<li><a href="../ajax/defer_render.html">Deferred DOM creation for extra speed</a></li>
					<li><a href="../ajax/null_data_source.html">Empty data source columns</a></li>
					<li><a href="../ajax/custom_data_property.html">Use a data source other than aaData (the default)</a></li>
					<li><a href="../ajax/objects_subarrays.html">Read column data from sub-arrays</a></li>
					<li><a href="../ajax/deep.html">Read column data from deeply nested properties</a></li>
				</ul>
				
				<h2>Plug-ins</h2>
				<ul>
					<li><a href="../plug-ins/plugin_api.html">Add custom API functions</a></li>
					<li><a href="../plug-ins/sorting_plugin.html">Sorting and automatic type detection</a></li>
					<li><a href="../plug-ins/sorting_sType.html">Sorting without automatic type detection</a></li>
					<li><a href="../plug-ins/paging_plugin.html">Custom pagination controls</a></li>
					<li><a href="../plug-ins/range_filtering.html">Range filtering / custom filtering</a></li>
					<li><a href="../plug-ins/dom_sort.html">Live DOM sorting</a></li>
					<li><a href="../plug-ins/html_sort.html">Automatic HTML type detection</a></li>
				</ul>
			</div>
			
			
			<div id="footer" class="clear" style="text-align:center;">
				<p>
					Please refer to the <a href="http://www.datatables.net/usage">DataTables documentation</a> for full information about its API properties and methods.<br>
					Additionally, there are a wide range of <a href="http://www.datatables.net/extras">extras</a> and <a href="http://www.datatables.net/plug-ins">plug-ins</a> which extend the capabilities of DataTables.
				</p>
				
				<span style="font-size:10px;">
					DataTables designed and created by <a href="http://www.sprymedia.co.uk">Allan Jardine</a> &copy; 2007-2011<br>
					DataTables is dual licensed under the <a href="http://www.datatables.net/license_gpl2">GPL v2 license</a> or a <a href="http://www.datatables.net/license_bsd">BSD (3-point) license</a>.
				</span>
			</div>
		</div>
	</body>
</html>