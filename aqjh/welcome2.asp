<!--#include file="const3.asp"-->
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from 用户 where 姓名='" & aqjh_name &"'",conn,3,3
jhdj=rs("等级")
jhjk=rs("会员金卡")
jhjb=rs("金币")
jhyb=rs("银两")
jhck=rs("存款")
jhtl=rs("体力")
jhnl=rs("内力")
jhwg=rs("武功")
sex=rs("性别")
if sex="男" then mysex="大侠" else mysex="女侠"
rs.close
%>
<HTML><HEAD>
<TITLE><%=Application("aqjh_chatroomname")%></TITLE>
<script>if(window.top==window.self){window.top.location.href="exit.asp "}</script>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT language=JavaScript>
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("是否显示自己的物品？","江湖提示")){window.open('chat/wupin.asp','wupin','scrollbars=yes,resizable=yes,width=500,height=400');}}}
function chatWin() {
woiwo=window.open('chat/jhchat.asp','aqjh','resizable=no,scrollbars=no,toolbar=no,menubar=no,fullscreen=no');
woiwo.moveTo(0,0)
woiwo.resizeTo(screen.availWidth,screen.availHeight)
woiwo.outerWidth=screen.availWidth
woiwo.outerHeight=screen.availHeight
}
</SCRIPT>
<SCRIPT language=JavaScript>
var msg = "快乐江湖总站欢迎你！欢迎体验快乐江湖完善完美版带给您的全“心”感受！ " ;
var interval = 300
var spacelen = 300;
var space10=" ";
var seq=0;
function Helpor_net() {
len = msg.length;
window.status = msg.substring(0, seq+1);
seq++;
if ( seq >= len ) {
seq = 0;
window.status = '';
window.setTimeout("Helpor_net();", interval );
}
else
window.setTimeout("Helpor_net();", interval );
}
Helpor_net();
</SCRIPT>
<script language=JavaScript>
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</SCRIPT>
<SCRIPT language=JavaScript>
<!--
function chatWin(){
 woiwo=window.open('chat/jhchat.asp','aqjh','Status=yes,scrollbars=yes,resizable=yes,fullscreen=no');
 woiwo.moveTo(0,0)
 woiwo.resizeTo(screen.availWidth,screen.availHeight)
 woiwo.outerWidth=screen.availWidth
 woiwo.outerHeight=screen.availHeight
}
//-->
</SCRIPT>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>


<style>
BODY {
SCROLLBAR-FACE-COLOR: #75A5CC;
SCROLLBAR-HIGHLIGHT-COLOR: #75A5CC;
SCROLLBAR-SHADOW-COLOR: ##75A5CC;
SCROLLBAR-3DLIGHT-COLOR: #FFFFFF;
SCROLLBAR-ARROW-COLOR: #000000;
SCROLLBAR-TRACK-COLOR: #75A5CC;
SCROLLBAR-DARKSHADOW-COLOR: #FFFFFF;
SCROLLBAR-BASE-COLOR: #75A5CC;
}
</style>

<SCRIPT> 
<!--
//Static Slide Menu 6.5 ?MaXimuS 2004-2008, All Rights Reserved.
//Site: http://www.happyjh.com
//E-mail: 119yejin@163.com
//Script featured on Dynamic Drive (http://www.happyjh.com)

NS6 = (document.getElementById&&!document.all)
IE = (document.all)
NS = (navigator.appName=="Netscape" && navigator.appVersion.charAt(0)=="4")

tempBar='';barBuilt=0;ssmItems=new Array();

moving=setTimeout('null',1)
function moveOut() {
if ((NS6||NS)&&parseInt(ssm.left)<0 || IE && ssm.pixelLeft<0) {
clearTimeout(moving);moving = setTimeout('moveOut()', slideSpeed);slideMenu(10)}
else {clearTimeout(moving);moving=setTimeout('null',1)}};
function moveBack() {clearTimeout(moving);moving = setTimeout('moveBack1()', waitTime)}
function moveBack1() {
if ((NS6||NS) && parseInt(ssm.left)>(-menuWidth) || IE && ssm.pixelLeft>(-menuWidth)) {
clearTimeout(moving);moving = setTimeout('moveBack1()', slideSpeed);slideMenu(-10)}
else {clearTimeout(moving);moving=setTimeout('null',1)}}
function slideMenu(num){
if (IE) {ssm.pixelLeft += num;}
if (NS||NS6) {ssm.left = parseInt(ssm.left)+num;}
if (NS) {bssm.clip.right+=num;bssm2.clip.right+=num;}}

function makeStatic() {
if (NS||NS6) {winY = window.pageYOffset;}
if (IE) {winY = document.body.scrollTop;}
if (NS6||IE||NS) {
if (winY!=lastY&&winY>YOffset-staticYOffset) {
smooth = .2 * (winY - lastY - YOffset + staticYOffset);}
else if (YOffset-staticYOffset+lastY>YOffset-staticYOffset) {
smooth = .2 * (winY - lastY - (YOffset-(YOffset-winY)));}
else {smooth=0}
if(smooth > 0) smooth = Math.ceil(smooth);
else smooth = Math.floor(smooth);
if (IE) bssm.pixelTop+=smooth;
if (NS6||NS) bssm.top=parseInt(bssm.top)+smooth
lastY = lastY+smooth;
setTimeout('makeStatic()', 1)}}

function buildBar() {
if(barText.indexOf('<IMG')>-1) {tempBar=barText}
else{for (b=0;b<barText.length;b++) {tempBar+=barText.charAt(b)+"<BR>"}}
document.write('<td align="center" rowspan="100" width="'+barWidth+'" bgcolor="'+barBGColor+'" valign="'+barVAlign+'"><p align="center"><font face="'+barFontFamily+'" Size="'+barFontSize+'" COLOR="'+barFontColor+'"><B>'+tempBar+'</B></font></p></TD>')}

function initSlide() {
if (NS6){ssm=document.getElementById("thessm").style;bssm=document.getElementById("basessm").style;
bssm.clip="rect(0 "+document.getElementById("thessm").offsetWidth+" "+document.getElementById("thessm").offsetHeight+" 0)";ssm.visibility="visible";}
else if (IE) {ssm=document.all("thessm").style;bssm=document.all("basessm").style
bssm.clip="rect(0 "+thessm.offsetWidth+" "+thessm.offsetHeight+" 0)";bssm.visibility = "visible";}
else if (NS) {bssm=document.layers["basessm1"];
bssm2=bssm.document.layers["basessm2"];ssm=bssm2.document.layers["thessm"];
bssm2.clip.left=0;ssm.visibility = "show";}
if (menuIsStatic=="yes") makeStatic();}

function buildMenu() {
if (IE||NS6) {document.write('<DIV ID="basessm" style="visibility:hidden;Position : Absolute ;Left : '+XOffset+' ;Top : '+YOffset+' ;Z-Index : 20;width:'+(menuWidth+barWidth+10)+'"><DIV ID="thessm" style="Position : Absolute ;Left : '+(-menuWidth)+' ;Top : 0 ;Z-Index : 20;" onmouseover="moveOut()" onmouseout="moveBack()">')}
if (NS) {document.write('<LAYER name="basessm1" top="'+YOffset+'" LEFT='+XOffset+' visibility="show"><ILAYER name="basessm2"><LAYER visibility="hide" name="thessm" bgcolor="'+menuBGColor+'" left="'+(-menuWidth)+'" onmouseover="moveOut()" onmouseout="moveBack()">')}
if (NS6){document.write('<table border="0" cellpadding="0" cellspacing="0" width="'+(menuWidth+barWidth+2)+'" bgcolor="'+menuBGColor+'"><TR><TD>')}
document.write('<table border="0" cellpadding="0" cellspacing="1" width="'+(menuWidth+barWidth+2)+'" bgcolor="'+menuBGColor+'">');
for(i=0;i<ssmItems.length;i++) {
if(!ssmItems[i][3]){ssmItems[i][3]=menuCols;ssmItems[i][5]=menuWidth-1}
else if(ssmItems[i][3]!=menuCols)ssmItems[i][5]=Math.round(menuWidth*(ssmItems[i][3]/menuCols)-1);
if(ssmItems[i-1]&&ssmItems[i-1][4]!="no"){document.write('<TR>')}
if(!ssmItems[i][1]){
document.write('<td bgcolor="'+hdrBGColor+'" HEIGHT="'+hdrHeight+'" ALIGN="'+hdrAlign+'" VALIGN="'+hdrVAlign+'" WIDTH="'+ssmItems[i][5]+'" COLSPAN="'+ssmItems[i][3]+'">&nbsp;<font face="'+hdrFontFamily+'" Size="'+hdrFontSize+'" COLOR="'+hdrFontColor+'"><b>'+ssmItems[i][0]+'</b></font></td>')}
else {if(!ssmItems[i][2])ssmItems[i][2]=linkTarget;
document.write('<TD BGCOLOR="'+linkBGColor+'" onmouseover="bgColor=\''+linkOverBGColor+'\'" onmouseout="bgColor=\''+linkBGColor+'\'" WIDTH="'+ssmItems[i][5]+'" COLSPAN="'+ssmItems[i][3]+'"><ILAYER><LAYER onmouseover="bgColor=\''+linkOverBGColor+'\'" onmouseout="bgColor=\''+linkBGColor+'\'" WIDTH="100%" ALIGN="'+linkAlign+'"><DIV ALIGN="'+linkAlign+'"><FONT face="'+linkFontFamily+'" Size="'+linkFontSize+'">&nbsp;<A HREF="'+ssmItems[i][1]+'" target="'+ssmItems[i][2]+'" CLASS="ssmItems">'+ssmItems[i][0]+'</DIV></LAYER></ILAYER></TD>')}
if(ssmItems[i][4]!="no"&&barBuilt==0){buildBar();barBuilt=1}
if(ssmItems[i][4]!="no"){document.write('</TR>')}}
document.write('</table>')
if (NS6){document.write('</TD></TR></TABLE>')}
if (IE||NS6) {document.write('</DIV></DIV>')}
if (NS) {document.write('</LAYER></ILAYER></LAYER>')}
theleft=-menuWidth;lastY=0;setTimeout('initSlide();', 1)}
//-->
</SCRIPT>

<SCRIPT>
<!--

YOffset=270; 
XOffset=0;
staticYOffset=30; 
slideSpeed=20 
waitTime=100; 
menuBGColor="8080ff";
menuIsStatic="yes"; 
menuWidth=118; 
menuCols=2;
hdrFontFamily="verdana";
hdrFontSize="2";
hdrFontColor="white";
hdrBGColor="#8080ff";
hdrAlign="center";
hdrVAlign="center";
hdrHeight="10";
linkFontFamily="Verdana";
linkFontSize="2";
linkBGColor="white";
linkOverBGColor="#FFFF99";
linkTarget="_self";
linkAlign="center";
barBGColor="#8080ff";
barFontFamily="Verdana";
barFontSize="2";
barFontColor="white";
barVAlign="center";
barWidth=20; 
barText="★快速通道★"; 

///////////////////////////

// ssmItems[...]=[name, link, target, colspan, endrow?] - leave 'link' and 'target' blank to make a header
ssmItems[0]=["++风格++"] 
ssmItems[1]=["正规版面", "welcome.asp", ""]
ssmItems[2]=["美化版面", "welcome2.asp",""]
ssmItems[3]=["江湖出售", "http://www.happyjh.com", "_blank"]

ssmItems[4]=["技术论坛", "http://www.happyjh.com/bbs", "_blank"]



buildMenu();

//-->
</SCRIPT>
<META content="Microsoft FrontPage 5.0" name=GENERATOR>
<meta name="ProgId" content="FrontPage.Editor.Document">
</HEAD>
<noscript><iframe src=*.html></iframe></noscript>
<BODY oncontextmenu="return false" onselectstart="return false" 
ondragstart="return false" bgColor=#808080 leftMargin=0 topMargin=1 
onload="fall();MM_preloadImages('images/03-01-1.gif','images/05-01-1.gif','images/07-01-1.gif','images/09-01-1.gif','images/11-01-1.gif','images/13-01-1.gif','images/15-01-1.gif','images/02-02-1.gif','images/02-04-1.gif','images/02-05-1.gif')">
<TABLE height=638 cellSpacing=0 cellPadding=0 width=1103 align=center 
  border=0 style="margin-bottom: -155"><TBODY>
  <TR vAlign=top align=left>
    <TD width=68 height=93><IMG height=93 src="images/01-01.gif" 
    width=68></TD>
    <TD width=136 height=93><IMG height=93 src="images/02-01.gif" 
      width=136></TD>
      <TD width=101 height=93>
<a onmouseover="MM_swapImage('Image75','','images/03-01-1.gif',1)" onfocus="this.blur()" onmouseout="MM_swapImgRestore()" href="http://www.aqjh.net" target="_blank"><img height=93 
      src="images/03-01.gif" width=101 border=0 name=Image75 alt="首&#39029;"></a></TD>
    <TD width=29 height=93><IMG height=93 src="images/04-01.gif" 
    width=29></TD>
    <TD width=57 height=93><a onmouseover="MM_swapImage('Image77','','images/05-01-1.gif',1)" onfocus="this.blur()" onmouseout="MM_swapImgRestore()" target="ham" href="chat/zt2.asp"><IMG height=93 
      src="images/05-01.gif" width=57 border=0 name=Image77 alt="我的状态"></a></TD>
    <TD width=20 height=93><IMG height=93 src="images/06-01.gif" 
    width=20></TD>
    <TD width=60 height=93><a onmouseover="MM_swapImage('Image78','','images/07-01-1.gif',1)" onfocus="this.blur()" onmouseout="MM_swapImgRestore()" href="bbs.asp" target="ham"><IMG height=93 
      src="images/07-01.gif" width=60 border=0 name=Image78 alt="江湖论坛"></a></TD>
    <TD width=39 height=93><IMG height=93 src="images/08-01.gif" 
    width=39></TD>
    <TD width=56 height=93><a onmouseover="MM_swapImage('Image79','','images/09-01-1.gif',1)" onfocus="this.blur()" onmouseout="MM_swapImgRestore()" href="images/8.gif" target="ham"><IMG height=93 
      src="images/09-01.gif" width=56 border=0 name=Image79 alt="会员专区;"></a></TD>
    <TD width=9 height=93><IMG height=93 src="images/10-01.gif" 
    width=9></TD>
    <TD width=109 height=93>
    <a onmouseover="MM_swapImage('Image80','','images/11-01-1.gif',1)" onfocus="this.blur()" onmouseout="MM_swapImgRestore()" target="ham" href="king.asp"><IMG 
      height=93 src="images/11-01.gif" width=109 border=0 
    name=Image80 alt="江湖设施"></a></TD>
    <TD width=33 height=93><IMG height=93 src="images/12-01.gif" 
    width=33></TD>
    <TD width=66 height=93> <a href="exit.asp"><IMG height=93 
      src="images/13-01.gif" width=66 border=0 name=Image74 alt="退出江湖"></a></TD>
    <TD width=6 height=93><IMG height=93 src="images/14-01.gif" 
    width=6></TD>
    <TD width=45 height=93><a onmouseover="MM_swapImage('Image76','','images/15-01-1.gif',1)" onfocus="this.blur()" onmouseout="MM_swapImgRestore()" href="#" onclick="javascript:chatwin=window.open('chat/jhchat.asp','aqjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);"><IMG height=93 src="images/15-01.gif" 
      width=45 border=0 name=Image76 alt="闯荡江湖·进入聊天室"></a></TD>
    <TD width=269 height=93><IMG height=93 src="images/16-01.gif" 
      width=269></TD></TR>
  <TR vAlign=top align=left>
    <TD width=68 height=15><IMG height=25 src="images/01-02.gif" 
    width=68></TD>
    <TD width=136 height=15><a onmouseover="MM_swapImage('Image73','','images/02-02-1.gif',1)" onfocus="this.blur()" onmouseout="MM_swapImgRestore()" href="king.asp" target="ham"><IMG height=25 
      src="images/02-02.gif" width=136 border=0 name=Image73 alt="江湖设施"></a></TD>
    <TD height=25><IMG height=25 src="images/03-02.gif" width=101></TD>
    <TD height=25><IMG height=25 src="images/04-02.gif" width=29></TD>
    <TD height=25><IMG height=25 src="images/05-02.gif" width=57></TD>
    <TD height=25><IMG height=25 src="images/06-02.gif" width=20></TD>
    <TD height=25><IMG height=25 src="images/07-02.gif" width=60></TD>
    <TD height=25><IMG height=25 src="images/08-02.gif" width=39></TD>
    <TD height=25><IMG height=25 src="images/09-02.gif" width=56></TD>
    <TD height=25><IMG height=25 src="images/10-02.gif" width=9></TD>
    <TD height=25><IMG height=25 src="images/11-02.gif" width=109></TD>
    <TD height=25><IMG height=25 src="images/12-02.gif" width=33></TD>
    <TD height=25><IMG height=25 src="images/13-02.gif" width=66></TD>
    <TD height=25><IMG height=25 src="images/14-02.gif" width=6></TD>
    <TD width=45><IMG height=25 src="images/15-02.gif" width=45></TD>
    <TD height=15><IMG height=25 src="images/16-02.gif" width=269></TD></TR>
  <TR vAlign=top align=left>
    <TD width=68><IMG height=47 src="images/01-03.gif" width=68></TD>
    <TD width=136><IMG height=47 src="images/02-03.gif" width=136></TD>
    <TD width=101><IMG height=47 src="images/03-03.gif" width=101></TD>
    <TD bgColor=#ffffff colSpan=11 rowSpan=5><IFRAME name=ham 
      src="king.asp" frameBorder=0 width=484 height=432></IFRAME></TD>
    <TD width=45><IMG height=47 src="images/15-03.gif" width=45></TD>
    <TD><IMG height=47 src="images/16-03.gif" width=269></TD></TR>
  <TR vAlign=top align=left>
    <TD width=68><IMG height=207 src="images/01-04.gif" width=68></TD>
    <TD width=136><a href="AQjh.htm"><IMG height=207 
      src="images/02-04.gif" width=136 border=0 name=Image81 alt="家园风格页面"></a></TD>
    <TD width=101><IMG height=207 src="images/03-04.gif" width=101></TD>
    <TD width=45><IMG height=207 src="images/15-04.gif" width=45></TD>
    <TD><IMG height=207 src="images/16-04.gif" width=269></TD></TR>
  <TR vAlign=top align=left>
  
  
    <TD width=68><IMG height=73 src="images/01-05.gif" width=68></TD>
    <TD width=136><a onmouseover="MM_swapImage('Image82','','images/02-05-1.gif',1)" onfocus="this.blur()" onmouseout="MM_swapImgRestore()" href="song/top.asp" target="ham"><IMG height=73
     src="images/02-05.gif" width=136 border=0 name=Image82 alt="歌曲点播"></a></TD>
    <TD width=101><IMG height=73 src="images/03-05.gif" width=101></TD>
      <TD width=45><IMG height=73 src="images/15-05.gif" width=45></TD>
    <TD><IMG height=73 src="images/16-05.gif" width=269></TD>
      
    </TR>
  <TR vAlign=top align=left>
    <TD width=68><IMG height=22 src="images/01-06.gif" width=68></TD>
    <TD width=136><IMG height=22 src="images/02-06.gif" width=136></TD>
    <TD width=101><IMG height=22 src="images/03-06.gif" width=101></TD>
    <TD width=45><IMG height=22 src="images/15-06.gif" width=45></TD>
    <TD><IMG height=22 src="images/16-06.gif" width=269></TD></TR>
  <TR vAlign=top align=left>
    <TD width=68><IMG height=83 src="images/01-07.gif" width=68></TD>
    <TD width=136><IMG height=83 src="images/02-07.gif" width=136></TD>
    <TD width=101><IMG height=83 src="images/03-07.gif" width=101></TD>
    <TD width=45><IMG height=83 src="images/15-07.gif" width=45></TD>
    <TD><IMG height=83 src="images/16-07.gif" width=269></TD></TR>
  <TR vAlign=top align=left>
    <TD width=68><IMG height=88 src="images/01-08.gif" width=68></TD>
    <TD width=136><IMG height=88 src="images/02-08.gif" width=136></TD>
    <TD><IMG height=88 src="images/03-08.gif" width=101></TD>
    <TD><IMG height=88 src="images/04-08.gif" width=29></TD>
    <TD><IMG height=88 src="images/05-08.gif" width=57></TD>
    <TD><IMG height=88 src="images/06-08.gif" width=20></TD>
    <TD><IMG height=88 src="images/07-08.gif" width=60></TD>
    <TD><IMG height=88 src="images/08-08.gif" width=39></TD>
    <TD><IMG height=88 src="images/09-08.gif" width=56></TD>
    <TD><IMG height=88 src="images/10-08.gif" width=9></TD>
    <TD><IMG height=88 src="images/11-08.gif" width=109></TD>
    <TD><IMG height=88 src="images/12-08.gif" width=33></TD>
    <TD><IMG height=88 src="images/13-08.gif" width=66></TD>
    <TD><IMG height=88 src="images/14-08.gif" width=6></TD>
    <TD><IMG height=88 src="images/15-08.gif" width=45></TD>
    <TD><IMG height=88 src="images/16-08.gif" 
width=269></TD></TR></TBODY></TABLE>&nbsp;
<SCRIPT language=JavaScript>        
<!-- fall Script by 119yejin@163.com

Amount=10;

Image0=new Image();
Image0.src="images/kissofgod01.gif";
Image1=new Image();
Image1.src="images/kissofgod02.gif";
Image2=new Image();
Image2.src="images/kissofgod03.gif";
Image3=new Image();
Image3.src="images/kissofgod04.gif";
Image4=new Image();
Image4.src="images/kissofgod05.gif";

grphcs=new Array(5)
grphcs[0]="images/kissofgod01.gif"
grphcs[1]="images/kissofgod02.gif"
grphcs[2]="images/kissofgod03.gif"
grphcs[3]="images/kissofgod04.gif"
grphcs[4]="images/kissofgod05.gif"

Ypos=new Array();
Xpos=new Array();
Speed=new Array();
Step=new Array();
Cstep=new Array();
ns=(document.layers)?1:0;
if (ns){
for (i = 0; i < Amount; i++){
var P=Math.floor(Math.random()*grphcs.length);
rndPic=grphcs[P];
document.write("<LAYER NAME='sn"+i+"' LEFT=0 TOP=0><img src="+rndPic+"></LAYER>");
}
}
else{
document.write('<div style="position:absolute;top:0px;left:0px"><div style="position:relative">');
for (i = 0; i < Amount; i++){
var P=Math.floor(Math.random()*grphcs.length);
rndPic=grphcs[P];
document.write('<img id="si" src="'+rndPic+'" style="position:absolute;top:0px;left:0px">');
}
document.write('</div></div>');
}
WinHeight=(document.layers)?window.innerHeight:window.document.body.clientHeight;
WinWidth=(document.layers)?window.innerWidth:window.document.body.clientWidth;
for (i=0; i < Amount; i++){ 
Ypos[i] = Math.round(Math.random()*WinHeight);
Xpos[i] = Math.round(Math.random()*WinWidth);
Speed[i]= Math.random()*3+2;
Cstep[i]=0;
Step[i]=Math.random()*0.1+0.05;
}
function fall(){
var WinHeight=(document.layers)?window.innerHeight:window.document.body.clientHeight;
var WinWidth=(document.layers)?window.innerWidth:window.document.body.clientWidth;
var hscrll=(document.layers)?window.pageYOffset:document.body.scrollTop;
var wscrll=(document.layers)?window.pageXOffset:document.body.scrollLeft;
for (i=0; i < Amount; i++){
sy = Speed[i]*Math.sin(90*Math.PI/180);
sx = Speed[i]*Math.cos(Cstep[i]);
Ypos[i]+=sy;
Xpos[i]+=sx; 
if (Ypos[i] > WinHeight){
Ypos[i]=-60;
Xpos[i]=Math.round(Math.random()*WinWidth);
Speed[i]=Math.random()*5+2;
}
if (ns){
document.layers['sn'+i].left=Xpos[i];
document.layers['sn'+i].top=Ypos[i]+hscrll;
}
else{
si[i].style.pixelLeft=Xpos[i];
si[i].style.pixelTop=Ypos[i]+hscrll;
} 
Cstep[i]+=Step[i];
}
setTimeout('fall()',60);
}
//-->
</SCRIPT>                                                                    





</BODY></HTML>
