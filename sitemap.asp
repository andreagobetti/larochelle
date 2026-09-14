<!--#include virtual="/setup.asp" -->
<%
Dim theDom, theRoot, theParent,theChild, theID, docInstruction

const NODE_ELEMENT = 1
const xmlns = "http://www.sitemaps.org/schemas/sitemap/0.9"

Set theDom = Server.CreateObject("Microsoft.XMLDOM")
Set theRoot = theDom.createElement("urlset")

Set theID = theDom.createAttribute("xmlns")
theID.Text = xmlNs
theRoot.setAttributeNode theID
theDom.appendChild theRoot

Set theParent = theDom.createNode(NODE_ELEMENT, "url", xmlns)     
Set theChild = theDom.createNode(NODE_ELEMENT, "loc", xmlns)
theChild.Text = "http://"&nomesito
theRoot.appendChild theParent
theParent.appendChild theChild

Set theChild = theDom.createNode(NODE_ELEMENT, "changefreq", xmlns)
theChild.Text = "weekly"
theRoot.appendChild theParent
theParent.appendChild theChild

'Elenco prodotti
sql="select prodotti.* from prodotti where visibilita=0"
set rs=conn.execute (sql)
do while not rs.eof
Set theParent = theDom.createNode(NODE_ELEMENT, "url", xmlns)     
Set theChild = theDom.createNode(NODE_ELEMENT, "loc", xmlns)
theChild.Text = "http://"&nomesito&"/product.asp?idpro="&rs("idpro")
theRoot.appendChild theParent
theParent.appendChild theChild

Set theChild = theDom.createNode(NODE_ELEMENT, "changefreq", xmlns)
theChild.Text = "weekly"
theRoot.appendChild theParent
theParent.appendChild theChild
rs.MoveNext
loop

'Elenco settori
sql="select settori.* from settori where nascondi=false"
set rs=conn.execute (sql)
do while not rs.eof
Set theParent = theDom.createNode(NODE_ELEMENT, "url", xmlns)     
Set theChild = theDom.createNode(NODE_ELEMENT, "loc", xmlns)
theChild.Text = "http://"&nomesito&"/category.asp?idsettore="&rs("idsettore")
theRoot.appendChild theParent
theParent.appendChild theChild

Set theChild = theDom.createNode(NODE_ELEMENT, "changefreq", xmlns)
theChild.Text = "weekly"
theRoot.appendChild theParent
theParent.appendChild theChild
rs.MoveNext
loop





Set docInstruction = theDom.createProcessingInstruction("xml","version='1.0' encoding='UTF-8'")
theDom.insertBefore docInstruction, theDom.childNodes(0)
theDom.Save Server.MapPath("/public/sitemap.xml")

%>