<!--#include virtual="/setup.asp" -->
<%
tabella="ordini"
'add2log queryeform(),0
if session("idadmin") = "" then call login()
if request("idord")<>"" then idord=request("idord")
%>
		
        <form id="form_email">
		<input type="hidden" name="idord" value="<%=idord%>">
		<input type="hidden" name="tabella" value="<%=tabella%>">
		
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#E5E5E5" class="tabella1">
          <tr valign="middle" align="right" >
            <td colspan="3" align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">
              <textarea name="conferma_ordine" id="testo_mail"  rows="6" style="width:100%;" ></textarea></td>
          </tr>
          <tr valign="middle" align="right" >
            <td width="33%" align="left" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><script>
risposte = new Array(); 
<%
set rsm=conn.execute("select * from ordini_risposte_predefinite where tipo=1 order by nome_risposta ;")
do while not rsm.eof
testo_risposta=replace(rsm("testo_risposta"),"'","\'")
testo_risposta=replace(testo_risposta,"&","\&")
testo_risposta=replace(testo_risposta,vbCrLf,"<br>")
%>risposte [<%=rsm("idrisposta")%>]='<%=testo_risposta%>';
		<%
rsm.movenext
loop
rsm.close
%>
function Inseriscirisposta(){
		tinyMCE.editors['conferma_ordine'].setContent(risposte[document.form1.risposta.options[document.form1.risposta.selectedIndex].value]);
}
</script>
              <select name="risposta" id="risposta" onChange="Inseriscirisposta();">
                <option value="">Scegli una risposta</option>
                <%
				set rsm=conn.execute("select * from ordini_risposte_predefinite where tipo=1 order by nome_risposta ;")
				do while not rsm.eof
					response.write "<option value=""" & rsm("idrisposta") &""""&">" & rsm("nome_risposta")& "</option>"
					rsm.movenext
				loop
				rsm.close
				set rsm=nothing
				%>
              </select></td>
            <td width="33%" align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;"><input type="submit" name="conferma" id="invia_mail_btn" value="Invia email al cliente"></td>
            <td width="33%" align="center" style="border-bottom: 1px solid gray; background: #E5E5E5; padding-top: 3px; padding-bottom:3px;">&nbsp;</td>
          </tr>
        </table>
		</form>
		
<script>
$(function() {

/* Configurazione 2: no filemanager, no youtube */
tinyMCE.init({
	// file_browser_callback: 'openKCFinder',
	language : 'it',
	selector:"#testo_mail",
	statusbar: false, 
	menubar : false,

	toolbar_items_size: 'small',
	auto_focus : false,
	plugins: [ 
	"advlist autolink lists link image charmap print preview hr anchor pagebreak",
	"searchreplace wordcount visualblocks visualchars code fullscreen",
	"insertdatetime media nonbreaking save table contextmenu directionality",
	"paste textcolor textcolor"
	],
	toolbar1: " undo redo |  fontsizeselect | bold italic forecolor | alignleft aligncenter alignright alignjustify | bullist numlist outdent indent ",
	image_advtab: true ,
	
    font_size_style_values: "12px,13px,14px,16px,18px,20px"//,
	
	//external_filemanager_path:"/filemanager123/",
	//filemanager_title:"Responsive Filemanager" ,
	//external_plugins: { "filemanager" : "/filemanager123/plugin.min.js"}
	//file_browser_callback: "responsivefilemanager"
	//   open_manager_upload_path: 'uploads/'
});
	$("#invia_mail_btn").on("click", function (e) {
		alert("invia");
    });

});
</script>