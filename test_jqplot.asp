<!DOCTYPE html>

<html>
<head>

    <title>Pie Charts and Options 2</title>

    <link class="include" rel="stylesheet" type="text/css" href="jquery/css/jquery.jqplot.min.css" />
    <!--[if lt IE 9]><script language="javascript" type="text/javascript" src="../excanvas.js"></script><![endif]-->
    <script class="include" type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
    
   
</head>
<body>
    <div class="colmask leftmenu">
      <div class="colleft">
        <div class="col1" id="example-content">

  
<!-- Example scripts go here -->


    <div id="chart7" style="margin-top:20px; margin-left:20px; width:460px; height:300px;"></div>



<script class="code" type="text/javascript">$(document).ready(function(){
  jQuery.jqplot.config.enablePlugins = true;
  plot7 = jQuery.jqplot('chart7', 
    [[['Verwerkende industrie', 9],['Retail', 8], ['Primaire producent', 7], 
    ['Out of home', 6],['Groothandel', 5], ['Grondstof', 4], ['Consument', 3], ['Bewerkende industrie', 2]]], 
    {
      title: ' ', 
      seriesDefaults: {shadow: true, renderer: jQuery.jqplot.PieRenderer, rendererOptions: { showDataLabels: true } }, 
      legend: { show:true }
    }
  );
});
</script>



<!-- End example scripts -->

<!-- Don't touch this! -->


    <script class="include" type="text/javascript" src="../jquery.jqplot.min.js"></script>
<!-- End Don't touch this! -->

<!-- Additional plugins go here -->

  <script class="include" type="text/javascript" src="../plugins/jqplot.pieRenderer.min.js"></script>

<!-- End additional plugins -->

        </div>
               </div>
    </div>

</body>


</html>
