<%@LANGUAGE="VBSCRIPT" CodePage="936"%>
<%
'happyjh.com
Option Explicit
Response.Buffer = True
Server.scriptTimeout="20"
on error resume next
Dim strSourceFile,objXML,objRootsite,AllNodesNum
strSourceFile=Server.MapPath("xml/info.xml")
Set objXML=Server.CreateObject("Microsoft.XMLDOM")
objXML.load(strSourceFile)
Set objRootsite=objXML.documentElement.selectSingleNode("qqinfo")
Dim startX,startY
startX=objRootsite.childNodes.item(0).childNodes.item(2).text
startY=objRootsite.childNodes.item(0).childNodes.item(3).text
Set objRootsite=nothing
Set objXML=nothing
%>
var verticalpos="frombottom"
if (!document.layers)
document.write('</div>')
function JSFX_FloatTopDiv()
{
	var startX = <%=startX%>;
	startY = <%=startY%>;
	var ns = (navigator.appName.indexOf("Netscape") != -1);
	var d = document;
	function ml(id)
	{
		var el=d.getElementById?d.getElementById(id):d.all?d.all[id]:d.layers[id];
		if(d.layers)el.style=el;
		el.sP=function(x,y){this.style.left=x;this.style.top=y;};
		el.x = startX;
		if (verticalpos=="fromtop")
		el.y = startY;
		else{
		el.y = ns ? pageYOffset + innerHeight : document.body.scrollTop + document.body.clientHeight;
		el.y -= startY;
		}
		return el;
	}
	window.stayTopLeft=function()
	{
		if (verticalpos=="fromtop"){
		var pY = ns ? pageYOffset : document.body.scrollTop;
		ftlObj.y += (pY + startY - ftlObj.y)/8;
		}
		else{
		var pY = ns ? pageYOffset + innerHeight : document.body.scrollTop + document.body.clientHeight;
		ftlObj.y += (pY - startY - ftlObj.y)/8;
		}
		ftlObj.sP(ftlObj.x, ftlObj.y);
		setTimeout("stayTopLeft()", 10);
	}
	ftlObj = ml("divStayTopLeft");
	stayTopLeft();
}
JSFX_FloatTopDiv();