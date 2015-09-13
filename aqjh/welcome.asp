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
<html>
<head>
<title>『<%=Application("aqjh_chatroomname")%>』-版本≮快乐江湖love9.9≯图形版</title>
<LINK href="home2.files/css.css" rel=stylesheet>
<style>
BODY{
CURSOR: url('chat/4.cur');
SCROLLBAR-HIGHLIGHT-COLOR:buttonface;
SCROLLBAR-SHADOW-COLOR:buttonface;
SCROLLBAR-3DLIGHT-COLOR:buttonhighlight;
SCROLLBAR-TRACK-COLOR:eeeeee;
SCROLLBAR-DARKSHADOW-COLOR:buttonshadow;
margin-top: 0px;
margin-bottom: 0px;
}
A {COLOR: #B6B6B6; TEXT-DECORATION: none}
A:hover {COLOR:yellow;TEXT-DECORATION: none}
TD {COLOR: #B6B6B6;FONT-SIZE: 12px; LINE-HEIGHT: 15px}
</style>

<SCRIPT> 
<!--
//Static Slide Menu 6.5 ?MaXimuS 2004-2008, All Rights Reserved.
//Site: http://www.7758530.com
//E-mail: 119yejin@163.com
//Script featured on Dynamic Drive (http://www.7758530.com)

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
ssmItems[3]=["江湖出售", "http://www.7758530.com", "_blank"]

ssmItems[4]=["技术论坛", "http://www.7758530.com/bbs", "_blank"]



buildMenu();

//-->
</SCRIPT>

<SCRIPT language=JavaScript>
<!--
function MM_showHideLayers() { //v2.0
var i, visStr, args, theObj;
args = MM_showHideLayers.arguments;
for (i=0; i<(args.length-2); i+=3) { //with arg triples (objNS,objIE,visStr)
visStr   = args[i+2];
if (navigator.appName == 'Netscape' && document.layers != null) {
theObj = eval(args[i]);
if (theObj) theObj.visibility = visStr;
} else if (document.all != null) { //IE
if (visStr == 'show') visStr = 'visible'; //convert vals
if (visStr == 'hide') visStr = 'hidden';
theObj = eval(args[i+1]);
if (theObj) theObj.style.visibility = visStr;
} }
}
//-->
</SCRIPT>
<script>if(window.top==window.self){window.top.location.href="AQJH.HTM"}</script>
<meta http-equiv="Content-Type" content="text/html;">
</head><noscript><iframe src=*.html></iframe></noscript>
<body bgColor=731702 topmargin=0 oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false>
<table border="0" cellpadding="0" cellspacing="0" width="775" align=center>
  <tr>
   <td><img src="home2.files/spacer.gif" width="13" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="12" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="35" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="88" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="27" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="8" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="26" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="6" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="6" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="188" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="33" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="129" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="7" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="6" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="7" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="9" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="111" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="7" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="22" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="15" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="5" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="15" height="1" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="1" border="0" alt=""></td>
  </tr>

  <tr>
   <td rowspan="19" background="home2.files/home2_left_1.jpg"></td>
   <td colspan="9"><img name="home2_r1_c2" src="home2.files/home2_r1_c2.jpg" width="396" height="61" border="0" alt=""></td>
   <td rowspan="3" colspan="11"><img name="home2_r1_c11" src="home2.files/home2_r1_c11.jpg" width="351" height="147" border="0" alt=""></td>
   <td rowspan="19" background="home2.files/home2_right_1.jpg" width="15" height="926" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="61" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="9"><img name="home2_r2_c2" src="home2.files/home2_r2_c2.jpg" width="396" height="39" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="39" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="2" colspan="9"><img name="home2_r3_c2" src="home2.files/home2_r3_c2.jpg" width="396" height="61" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="47" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="11"><img name="home2_r4_c11" src="home2.files/home2_r4_c11.jpg" width="351" height="14" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="14" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="4" colspan="2"><img name="home2_r5_c2" src="home2.files/home2_r5_c2.jpg" width="47" height="84" border="0" alt=""></td>
   <td colspan="14" background="home2.files/home2_r5_c4.jpg" align=center><a href=bbs/ target=_blank><font color=yellow>江湖论坛</a> | <a href=# onClick="window.open('HCJS/bang/meeting.asp','aqjh_win','scrollbars=0,resizable=0,width=690,height=440')"><font color=yellow>武林大会</a> | <a href=# onClick="window.open('hcjs/pig/index.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=450')"><font color=yellow>养猪大赛</a> | <a href=# onClick="window.open('garden/index.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=450')"><font color=yellow>花园大赛</a> | <a href=jhshow/jl.asp target=_blank><font color=yellow>秀装大赛</a> | <a href=jhshow/ target=_blank><font color=yellow>虚拟形象</a> | <A href="TOP/TOP47.ASP" target=_blank><font color=yellow>在线排行</A> | <a href="#"  onClick="javascript:chatwin=window.open('top/top.htm','','scrollbars=no,resizable=no,width=760,height=400,top=0,left=0')"><font color=yellow>江湖排行</a> | <a href=# onClick="window.open('hcjs/sendphoto/PHOTO.ASP','aqjh_win','scrollbars=1,resizable=1,width=750,height=500')"><font color=yellow>网友照片</a> | <%if aqjh_grade>=10 then%><a href="aqjh_admin/login.asp" target="_blank"><font color=yellow>后台管理</a><%else%><a href=chat/hy.asp target=_blank><font color=yellow>购买会员</a><%end if%></td>
   <td rowspan="2" colspan="4"><img name="home2_r5_c18" src="home2.files/home2_r5_c18.jpg" width="49" height="33" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="21" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="3" colspan="2"><img name="home2_r6_c4" src="home2.files/home2_r6_c4.jpg" width="115" height="63" border="0" usemap="#m_home2_r6_c4" alt=""></td>
   <td colspan="12"><img name="home2_r6_c6" src="home2.files/home2_r6_c6.jpg" width="536" height="12" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="12" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="13"><img src=home2.files/home2_r7_c6.jpg usemap="#m_home2_r7_c6" border=0></td>
   <td colspan="3"><img name="home2_r7_c19" src="home2.files/home2_r7_c19.jpg" width="42" height="35" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="35" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="3" colspan="6"><img name="home2_r8_c6" src="home2.files/home2_r8_c6.jpg" width="267" height="188" border="0" usemap="#m_home2_r8_c6" alt=""></td>
   <td rowspan="6" colspan="3"><img name="home2_r8_c12" src="home2.files/home2_r8_c12.jpg" width="142" height="226" border="0" usemap="#m_home2_r8_c12" alt=""></td>
   <td rowspan="2" colspan="2"><img name="home2_r8_c15" src="home2.files/home2_r8_c15.jpg" width="16" height="26" border="0" alt=""></td>
   <td rowspan="2" colspan="3"><img name="home2_r8_c17" src="home2.files/home2_r8_c17.jpg" width="140" height="26" border="0" alt=""></td>
   <td rowspan="2" colspan="2"><img name="home2_r8_c20" src="home2.files/home2_r8_c20.jpg" width="20" height="26" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="16" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="3" colspan="4"><img name="home2_r9_c2" src="home2.files/home2_r9_c2.jpg" width="162" height="180" border="0" usemap="#m_home2_r9_c2" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="10" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="3" colspan="2" background="home2.files/home2_left_2.jpg"></td>
   <td rowspan="3" colspan="3" background="home2.files/home2_center_1.jpg" valign=top align=center>
<table width=130><tr><td valign=top>
欢迎<font color=yellow><%=aqjh_name%></font><%=mysex%>光临<BR><BR>
等级:<%=jhdj%><BR>
管理:<%=aqjh_grade%><BR>
现金:<%=jhjk%><BR>
金币:<%=jhjb%><BR>
银子:<%=jhyb%><BR>
存款:<%=jhck%><BR>
体力:<%=jhtl%><BR>
内力:<%=jhnl%><BR>
武功:<%=jhwg%>
</td></tr></table>
</td>
   <td rowspan="3" colspan="2" background="home2.files/home2_right_2.jpg"></td>
   <td><img src="home2.files/spacer.gif" width="1" height="162" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="4"><img name="home2_r11_c6" src="home2.files/home2_r11_c6.jpg" width="46" height="8" border="0" alt=""></td>
   <td rowspan="6" colspan="2"><img name="home2_r11_c10" src="home2.files/home2_r11_c10.jpg" width="221" height="264" border="0" usemap="#m_home2_r11_c10" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="8" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="4" colspan="3"><img name="home2_r12_c2" src="home2.files/home2_r12_c2.jpg" width="135" height="220" border="0" usemap="#m_home2_r12_c2" alt=""></td>
   <td rowspan="5" colspan="5"><img name="home2_r12_c5" src="home2.files/home2_r12_c5.jpg" width="73" height="256" border="0" usemap="#m_home2_r12_c5" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="17" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="2" colspan="7"><img name="home2_r13_c15" src="home2.files/home2_r13_c15.jpg" width="176" height="31" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="13" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="3" colspan="3"><img name="home2_r14_c12" src="home2.files/home2_r14_c12.jpg" width="142" height="226" border="0" usemap="#m_home2_r14_c12" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="18" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="2" colspan="2" background="home2.files/home2_left_2.jpg"></td>
   <td rowspan="2" colspan="3" background="home2.files/home2_center_1.jpg" align=center><script src=top.asp></script></td>
   <td rowspan="2" colspan="2" background="home2.files/home2_right_2.jpg"></td>
   <td><img src="home2.files/spacer.gif" width="1" height="172" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="3"><img name="home2_r16_c2" src="home2.files/home2_r16_c2.jpg" width="135" height="36" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="36" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="2"><img name="home2_r17_c2" src="home2.files/home2_r17_c2.jpg" width="12" height="37" border="0" alt=""></td>
   <td rowspan="2" colspan="4"><img name="home2_r17_c3" src="home2.files/home2_r17_c3.jpg" width="158" height="37" border="0" alt=""></td>
   <td rowspan="2"><img name="home2_r17_c7" src="home2.files/home2_r17_c7.jpg" width="26" height="37" border="0" alt=""></td>
   <td rowspan="2" colspan="6"><img name="home2_r17_c8" src="home2.files/home2_r17_c8.jpg" width="369" height="37" border="0" alt=""></td>
   <td rowspan="3"><img name="home2_r17_c14" src="home2.files/home2_r17_c14.jpg" width="6" height="245" border="0" alt=""></td>
   <td colspan="7"><img name="home2_r17_c15" src="home2.files/home2_r17_c15.jpg" width="176" height="12" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="12" border="0" alt=""></td>
  </tr>
  <tr>
   <td rowspan="2" colspan="6" bgcolor=#A90005 align=center><script language=javascript src="jhshow/disp.asp?name=<%=session("aqjh_name")%>&id=0&sex=<%=sex%>"></script></td>
   <td rowspan="2"><img name="home2_r18_c21" src="home2.files/home2_r18_c21.jpg" width="5" height="233" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="25" border="0" alt=""></td>
  </tr>
  <tr>
   <td background="home2.files/home2_left_3.jpg"></td>
   <td colspan="4" background="home2.files/home2_center_2.jpg" align=center valign=top><table width=90%><tr><td><script language=javascript src="diary/toplist.asp?tlen=8&n=7"></script></td></tr></table></td>
   <td background="home2.files/home2_left_4.jpg"></td>
   <td colspan="6" background="home2.files/home2_center_3.jpg" align=center valign=top><table width=100% border=0><tr><td width=10></td><td>
<span id=yuzi>Loading...</span>
<script src=bbs/new.asp?id=yuzi&TopicCount=11&Count=14&showtime=1&icon=·></script></td></tr></table></td>
   <td><img src="home2.files/spacer.gif" width="1" height="208" border="0" alt=""></td>
  </tr>
  <tr>
   <td colspan="3"><img name="home2_r20_c1" src="home2.files/home2_r20_c1.jpg" width="60" height="43" border="0" alt=""></td>
   <td colspan="2"><img name="home2_r20_c4" src="home2.files/home2_r20_c4.jpg" width="115" height="43" border="0" alt=""></td>
   <td colspan="3"><img name="home2_r20_c6" src="home2.files/home2_r20_c6.jpg" width="40" height="43" border="0" alt=""></td>
   <td colspan="4"><img name="home2_r20_c9" src="home2.files/home2_r20_c9.jpg" width="356" height="43" border="0" alt=""></td>
   <td colspan="3"><img name="home2_r20_c13" src="home2.files/home2_r20_c13.jpg" width="20" height="43" border="0" alt=""></td>
   <td colspan="2"><img name="home2_r20_c16" src="home2.files/home2_r20_c16.jpg" width="120" height="43" border="0" alt=""></td>
   <td colspan="5"><img name="home2_r20_c18" src="home2.files/home2_r20_c18.jpg" width="64" height="43" border="0" alt=""></td>
   <td><img src="home2.files/spacer.gif" width="1" height="43" border="0" alt=""></td>
  </tr>
</table>
</body>
</html>
<!--#include file="map.asp"-->
<DIV id=overDiv style="Z-INDEX: 1; POSITION: absolute"></DIV>
<SCRIPT language=JavaScript src="home2.files/overlib.js"></SCRIPT>
<map name="m_home2_r6_c4">
<area shape="rect" coords="9,12,109,47" href="#" onclick="javascript:chatwin=window.open('chat/jhchat.asp','aqjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);">
</map>
<map name="m_home2_r8_c6">
<area shape="rect" coords="218,132,259,168" onmouseover="drs('<b>∷仁和药铺∷</b><br>江湖神医，悬壶济<br>世，只要你还有一口<br>气，俺保证救活你'); return true;" onclick="window.open('hcjs/jhjs/yaopu.asp','aqjh_win','left=0,scrollbars=yes,resizable=yes,resizable=0,width=700,height=400')" 
      onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="154,156,210,188" onmouseover="drs('<b>∷武器商店∷</b><br>闯荡江湖没有武器可<br>是危险大大的'); return true;" onclick="window.open('hcjs/jhjs/wuqi.asp','aqjh_win','left=0,scrollbars=yes,resizable=yes,width=550,height=450')" 
      onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="78,153,127,188" href=# onClick="window.open('hcjs/jhjs/dan.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')" onmouseover="drs('<b>∷江湖当铺∷</b><br>没钱了，把你值钱的东西拿来当吧！'); return true;" onmouseout="nd(); return true;">
<area shape="rect" coords="0,157,46,188" href=# onClick="window.open('jhjd/jd.asp','aqjh_win','scrollbars=0,resizable=0,width=740,height=450')"  onmouseover="drs('<b>∷快乐酒店∷</b><br>来此喝酒可以增加属<br>性和状态，不过要看<br>看你腰包鼓不鼓了'); return true;" onmouseout="nd(); return true;">
<area shape="rect" coords="169,97,213,134" href=# onClick="window.open('jhmp/wuguan.asp','aqjh_win','scrollbars=0,resizable=0,width=700,height=400')"  onmouseover="drs('<b>∷江湖密室∷</b><br>有师傅的可以来密室里习武！'); return true;" onmouseout="nd(); return true;">
<area shape="rect" coords="186,52,234,86" onmouseover="drs('<b>∷江湖帮派∷</b><br>闯荡江湖找个好老大<br>照顾准没错'); return true;" onclick="window.open('jhmp/index.asp','aqjh_win','left=0,scrollbars=yes,resizable=yes,width=760,height=450')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="46,96,94,128" href=# onClick="window.open('hcjs/pub/pub.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')"  onmouseover="drs('<b>∷龙门客栈∷</b><br>跑了一天，一定累了吧，来这里歇歇脚！'); return true;" onmouseout="nd(); return true;">
<area shape="rect" coords="84,61,131,93" href=# onClick="window.open('hcjs/yhy/','aqjh_win','scrollbars=1,resizable=0,width=680,height=470')" onmouseover="drs('<b>∷江湖青楼∷</b><br>这里为你全套服务，来吧，说你呢？帅哥！'); return true;" onmouseout="nd(); return true;">
<area shape="rect" coords="0,56,59,81" onmouseover="drs('<b>∷富贵金赌坊∷</b><br>想当年俺是穷光蛋，<br>出来是腰缠万贯，来<br>看看你的运气如何？'); return true;" onClick="window.open('bet/betindex.asp','aqjh_win','left=0,scrollbars=0,resizable=0,width=700,height=400')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="8,16,74,49" onmouseover="drs('<b>∷江湖伐木∷</b><br>从树林采集木材'); return true;" 
      onclick="window.open('work/famu/famumain.asp','aqjh_win','left=0,scrollbars=yes,resizable=yes,width=600,height=370')" 
      onmouseout="nd(); return true;" href="#">
</map>
<map name="m_home2_r8_c12">
<area shape="rect" coords="13,164,80,206" onmouseover="drs('<b>∷会员卡片屋∷</b><br>只有会员才能购买的<br>好东东，你有会员费<br>吗？');return true;" onclick="window.open('hcjs/card/card.asp','aqjh_win','left=0,scrollbars=1,resizable=0,width=700,height=400')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="26,102,93,153" onmouseover="drs('<b>∷炼制丹药∷</b><br>比那太上老君的炼丹<br>炉还厉害，我炼！'); return true;" onclick="window.open('twt/peiyao/peiyao.asp','aqjh_win','left=0,scrollbars=yes,resizable=yes,width=780,height=460')" 
      onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="82,16,136,51" onmouseover="drs('<b>∷遗迹藏宝∷</b><br>据说是魔鬼浪子藏了个<br>美女在这里，还有宝物<br>嘿嘿......'); return true;" onclick="window.open('twt/wabao/wabao.asp','aqjh_win','left=0,scrollbars=yes,resizable=yes,width=450,height=340')" 
 onmouseout="nd(); return true;" href="#">
</map>
<map name="m_home2_r9_c2">
<area shape="rect" coords="72,149,120,180" onmouseover="drs('<b>∷江湖股票∷</b><br>买卖自由想发财就<br>财就来试试运气赚了<br>钱可别忘了我！!'); return true;" 
 onClick="window.open('gupiao/stock.asp','aqjh_win','left=0,scrollbars=1,resizable=0,width=700,height=400')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="34,100,82,133" onmouseover="drs('<b>∷官府天牢∷</b><br>少做点坏事哦......<br>不然......嘿嘿...'); return true;" onclick="window.open('yamen/laofan.asp','aqjh_win','left=0,scrollbars=no,resizable=0,width=700,height=400')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="72,20,120,57" onmouseover="drs('<b>∷江湖官府∷</b><br>江湖管理投诉的地方'); return true;" 
      onclick="window.open('yamen/admin.asp','aqjh_win','scrollbars=0,resizable=0,width=740,height=450')" 
      onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="12,0,58,34" onmouseover="drs('<b>∷江湖采矿∷</b><br>从山里采集矿石'); return true;" 
      onclick="window.open('work/mine/minemain.asp','aqjh_win','left=0,scrollbars=yes,resizable=1,width=500,height=370')" 
      onmouseout="nd(); return true;" href="#">
</map>
<map name="m_home2_r11_c10">
<area shape="rect" coords="100,196,180,256" href="#" alt="" onmouseover="drs('<b>∷养猪∷</b><br>放牛太麻烦，还是来养猪吧!'); return true;" onclick="window.open('hcjs/pig/zhu.asp','aqjh_win','left=0,scrollbars=yes,resizable=yes,width=670,height=400')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="177,31,221,83" onmouseover="drs('<b>∷极地采冰∷</b><br>采集冰块然后融化成冰水，很珍贵的'); return true;" onclick="window.open('work/ice/icemain.asp','aqjh_win','scrollbars=1,resizable=1,width=670,height=430')" onmouseout="nd(); return true;" href="#" >
<area shape="rect" coords="86,32,129,73" onmouseover="drs('<b>∷江湖花店∷</b><br>很漂亮的鲜花，应有尽有...'); return true;" onclick="window.open('hcjs/jhjs/hua.asp','aqjh_win','left=0,scrollbars=1,resizable=1,width=670,height=460')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="0,48,68,87" onmouseover="drs('<b>∷可爱宠物店∷</b><br>各类宠物应有尽有，<br>大侠要不要搞只回去<br>玩玩？还可以帮助你<br>攻击敌人！'); return true;" onclick="window.open('hcjs/jhjs/cw.asp','aqjh_win','left=0,scrollbars=yes,resizable=yes,width=550,height=400')" 
      onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="15,0,63,25" href=# onClick="window.open('yamen/yiyuan.asp','aqjh_win','scrollbars=1,resizable=0,width=700,height=400')" onmouseover="drs('<b>∷人民医院∷</b><br>快乐江湖第一人民医院...！'); return true;" onmouseout="nd(); return true;">
</map>
<map name="m_home2_r12_c2">
<area shape="rect" coords="47,171,122,208" onmouseover="drs('<b>∷转世投胎∷</b><br>这里是转世投胎轮回盘，通过它你可以到另一个世界...'); return true;" onclick="window.open('HCJS/zstt/zs.asp','aqjh_win','left=0,scrollbars=0,resizable=0,width=690,height=440')" 
      onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="62,79,135,116" onmouseover="drs('<b>∷秘密花园∷</b><br>在这里可以种花，<br>照顾花呀，这里的花<br>可是买不到的!'); return true;" onclick="window.open('garden/hua.htm','aqjh_win','left=0,scrollbars=yes,resizable=yes,width=670,height=400')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="24,30,72,68" onmouseover="drs('<b>∷领取工资∷</b><br>发钱喽，不知是否有<br>奖金？'); return true;" onclick="window.open('jhmp/money.asp','aqjh_win','left=0,scrollbars=yes,resizable=no,width=430,height=340')" onmouseout="nd(); return true;" href="#">
</map>
<map name="m_home2_r12_c5">
<area shape="rect" coords="0,0,55,39" onmouseover="drs('<b>∷救助孤儿∷</b><br>这可是积德行善的好<br>事，会增加道德值哦'); return true;" onclick="window.open('jhmp/gry.asp','aqjh_win','left=0,scrollbars=no,resizable=no,width=350,height=230')" 
 onmouseout="nd(); return true;" href="#">
</map>
<map name="m_home2_r14_c12">
<area shape="rect" coords="19,80,77,158" onmouseover="drs('<b>∷江湖祈愿∷</b><br>很灵的哦......<br>不信.....您就试试...'); return true;" onclick="window.open('hcjs/wish/wish.asp','aqjh_win','left=0,scrollbars=no,resizable=0,width=700,height=400')" onmouseout="nd(); return true;" href="#">
<area shape="rect" coords="69,9,119,48" onmouseover="drs('<b>∷职业转换处∷</b><br>想换换职业来这儿吧'); return true;" onclick="window.open('work/changework.asp','aqjh_win','left=0,scrollbars=0,resizable=0,width=480,height=380')" onmouseout="nd(); return true;" href="#">
</map>
<map name="m_home2_r7_c6">
<area shape="rect" coords="471,9,529,27" href=exit.asp target=_top>
<area shape="rect" coords="396,10,456,26" href="#" onmousedown="MM_showHideLayers('document.layers[\'Layer6\']','document.all[\'Layer6\']','show')">
<area shape="rect" coords="317,10,380,28" href="#" onmousedown="MM_showHideLayers('document.layers[\'Layer5\']','document.all[\'Layer5\']','show')">
<area shape="rect" coords="243,9,305,28" href="#" onmousedown="MM_showHideLayers('document.layers[\'Layer4\']','document.all[\'Layer4\']','show')">
<area shape="rect" coords="167,9,228,29" href="#" onmousedown="MM_showHideLayers('document.layers[\'Layer3\']','document.all[\'Layer3\']','show')">
<area shape="rect" coords="90,8,154,29" href="#" onmousedown="MM_showHideLayers('document.layers[\'Layer2\']','document.all[\'Layer2\']','show')">
<area shape="rect" coords="8,6,81,28" href="#" onmousedown="MM_showHideLayers('document.layers[\'Layer1\']','document.all[\'Layer1\']','show')">
</map>