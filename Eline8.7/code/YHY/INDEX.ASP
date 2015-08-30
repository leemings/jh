<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
info=Session("info")
dim conn,rs
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<link rel="stylesheet" type="text/css" href="cc.css">
</head>

<body background="pic/bg.jpg">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="86%" id="AutoNumber1" height="134">
  <tr>
    <td width="26%" height="117"><img border="0" src="pic/a1a.gif"></td>
    <td width="74%" height="117">
    <img border="0" src="pic/jh1.gif" width="215" height="333" align="left"><br>
    　<p><img border="0" src="jy.gif"></p>
    <p>
    <span lang="zh-cn"><a href=xiaojie.asp >窈窕淑女,君子好逑.</a></span></p>
    <p><span lang="zh-cn"><a href=dengji.asp>孤苦伶仃,流落风尘.</a></span></p>
    <p><span lang="zh-cn"><a href=delgirl.asp>失足之恨,从良未晚</a></span><p>
    <span lang="zh-cn"><a href="sendphoto.asp">闭月羞花,沉鱼落雁</a></span><p>
    <span lang="zh-cn"><a href="money.asp">克扣剥削,命薄如纸</a></span></td>
  </tr>
  <tr>
    <td width="100%" height="19" colspan="2">
      <p align="center"><span lang="zh-cn">版权所有:『快乐江湖』&nbsp; 2003-2004</span>
    </td>
  </tr>
</table>


<SCRIPT language=JavaScript1.2>
var snowsrc="pic/1.gif"; var no = 18;
var ns4up = (document.layers) ? 1 : 0; var ie4up = (document.all) ? 1 : 0; var dx, xp, yp; var am, stx, sty;
var i, doc_width = 800, doc_height = 1200;
if (ns4up) {
doc_width = self.innerWidth;
doc_height = self.innerHeight;
} else if (ie4up) {
doc_width = document.body.clientWidth;
doc_height = document.body.clientHeight;
}
dx = new Array();
xp = new Array();
yp = new Array();
am = new Array();
stx = new Array();
sty = new Array();
for (i = 0; i < no; ++ i) {
dx[i] = 0;
xp[i] = Math.random()*(doc_width-50);
yp[i] = Math.random()*doc_height;
am[i] = Math.random()*20;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 0.7 + Math.random();
if (ns4up) {
if (i == 0) {
document.write("<layer name=\"dot"+ i +"\" left=\"15\" top=\"15\" visibility=\"show\"><img src='"+snowsrc+"' border=\"0\"></a></layer>");
} else {
document.write("<layer name=\"dot"+ i +"\" left=\"15\" top=\"15\" visibility=\"show\"><img src='"+snowsrc+"' border=\"0\"></layer>");
}
} else if (ie4up) {
if (i == 0) {
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: absolute; Z-INDEX: "+ i +"; VISIBILITY: visible; TOP: 15px; LEFT: 15px;\"><img src='"+snowsrc+"' border=\"0\"></a></div>");
} else {
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: absolute; Z-INDEX: "+ i +"; VISIBILITY: visible; TOP: 15px; LEFT: 15px;\"><img src='"+snowsrc+"' border=\"0\"></div>");
}
}
}
function snowNS() {
for (i = 0; i < no; ++ i) {
yp[i] += sty[i];
if (yp[i] > doc_height-50) {
xp[i] = Math.random()*(doc_width-am[i]-30);
yp[i] = 0;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 0.7 + Math.random();
doc_width = self.innerWidth;
doc_height = self.innerHeight;
}
dx[i] += stx[i];
document.layers["dot"+i].top = yp[i];
document.layers["dot"+i].left = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("snowNS()", 10);
}
function snowIE() {
for (i = 0; i < no; ++ i) {
yp[i] += sty[i];
if (yp[i] > doc_height-50) {
xp[i] = Math.random()*(doc_width-am[i]-30);
yp[i] = 0;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 0.7 + Math.random();
doc_width = document.body.clientWidth;
doc_height = document.body.clientHeight;
}
dx[i] += stx[i];
document.all["dot"+i].style.pixelTop = yp[i];
document.all["dot"+i].style.pixelLeft = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("snowIE()", 10);
}
if (ns4up) {
snowNS();
} else if (ie4up) {
snowIE();
}
         </SCRIPT>


</body>

</html>