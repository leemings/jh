<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<title>快乐江湖-风雪舞剑♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="File-List" href="fengxue.files/filelist.xml">

<link rel="stylesheet" href="../style4.css" type="text/css">
<script language="JavaScript">

//Configure below to change URL path to the snow image
var snowsrc="look.gif"
// Configure below to change number of snow to render
var no = 12;

var ns4up = (document.layers) ? 1 : 0;  // browser sniffer
var ie4up = (document.all) ? 1 : 0;

var dx, xp, yp;    // coordinate and position variables
var am, stx, sty;  // amplitude and step variables
var i, doc_width = 800, doc_height = 600;

if (ns4up) {
doc_width = self.innerWidth;
doc_height = self.innerHeight;
} else if (ie4up) {
}

dx = new Array();
xp = new Array();
yp = new Array();
am = new Array();
stx = new Array();
sty = new Array();

for (i = 0; i < no; ++ i) {
dx[i] = 0;                        // set coordinate variables
xp[i] = Math.random()*(doc_width-50);  // set position variables
yp[i] = Math.random()*doc_height;
am[i] = Math.random()*20;         // set amplitude variables
stx[i] = 0.02 + Math.random()/10; // set step variables
sty[i] = 0.7 + Math.random();     // set step variables
if (ns4up) {                      // set layers
if (i == 0) {
document.write("<layer name=\"dot"+ i +"\" left=\"35\" top=\"35\" visibility=\"show\"><img src='"+snowsrc+"' border=\"0\"></a></layer>");
} else {
document.write("<layer name=\"dot"+ i +"\" left=\"35\" top=\"35\" visibility=\"show\"><img src='"+snowsrc+"' border=\"0\"></layer>");
}
} else if (ie4up) {
if (i == 0) {
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: absolute; Z-INDEX: "+ i +"; VISIBILITY: visible; TOP: 15px; LEFT: 15px;\"><a href=\"#\"><img src='"+snowsrc+"' border=\"0\"></a></div>");
} else {
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: absolute; Z-INDEX: "+ i +"; VISIBILITY: visible; TOP: 15px; LEFT: 15px;\"><img src='"+snowsrc+"' border=\"0\"></div>");
}
}
}

function snowNS() {  // Netscape main animation function
for (i = 0; i < no; ++ i) {  // iterate for every dot
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

function snowIE() {  // IE main animation function
for (i = 0; i < no; ++ i) {  // iterate for every dot
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
setTimeout("snowIE()", 20);
}

if (ns4up) {
snowNS();
} else if (ie4up) {
snowIE();
}
</script>
<script language="JavaScript">
<!--
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->

function MM_findObj(n, d) { //v4.0
var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script>
<!--[if !mso]>
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]-->
<style fprolloverstyle>A:hover {color: #FF0000; font-weight: bold}
</style>
<!--[if gte mso 9]>
<xml><o:shapedefaults v:ext="edit" spidmax="1027"/>
</xml><![endif]-->
<style type="text/css">
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</head>
<body bgcolor="#000000" text="#000000" leftmargin="0" topmargin="0" style=overflow-x:hidden;overflow-y:hidden oncontextmenu=self.event.returnValue=false background="8.JPG" bgproperties="fixed">
<script language="JavaScript">
<!--
function MM_showHideLayers() { //v3.0
var i,p,v,obj,args=MM_showHideLayers.arguments;
for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
obj.visibility=v; }
}

function MM_displayStatusMsg(msgStr) { //v1.0
status=msgStr;
document.MM_returnValue = true;
}
//-->
</script>
<div id="Layer1" style="position: absolute; left: -10; top: 84; width: 754; height: 422; z-index: 1">
<form method=POST action='huayuan.asp' name=xiaohuayuan>
<input name=h value=1 type=hidden>
<div id="lt" style="position:absolute; left:119px; top:97px; width:72px; height:153px; z-index:1; visibility: hidden">
<table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休息 type="submit" name=sh>
</div>
</td>
</tr>
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=挥剑 type="submit" name=sh>
</div>
</td>
</tr>
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=斩杀 type="submit" name=sh>
</div>
</td>
</tr>
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破坏 type="submit" name=sh>
</div>
</td>
</tr>
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=查找 type="submit" name=sh>
</div>
</td>
</tr>
</table>
</div>
<div id="ssx" style="position:absolute; left:400px; top:245px; width:72px; height:153px; z-index:1; visibility: hidden">
</div>
<div id="xx" style="position:absolute; left:343px; top:113px; width:72px; height:153px; z-index:1; visibility: hidden">
</div>
<div id="js" style="position:absolute; left:659px; top:210px; width:72px; height:153px; z-index:1; visibility: hidden">
</div>
<div id="tud" style="position:absolute; left:246px; top:17px; width:72px; height:153px; z-index:1; visibility: hidden">
<table width="100" border="1" cellspacing="0" cellpadding="0" bordercolor="#999999">
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=休息 type="submit" name=sh22>
</div>
</td>
</tr>
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=挥剑 type="submit" name=sh22>
</div>
</td>
</tr>
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=斩杀 type="submit" name=sh22>
</div>
</td>
</tr>
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=破坏 type="submit" name=sh22>
</div>
</td>
</tr>
<tr bgcolor="#FFCC33">
<td height="30">
<div align="center" class="unnamed1">
<input onMouseOut="this.style.backgroundColor='#FFCC33'" onMouseOver="this.style.backgroundColor='#FFF7EE'" style="BACKGROUND-COLOR: #FFCC33; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=查找 type="submit" name=sh22>
</div>
</td>
</tr>
</table>
</div>
<div id="du" style="position:absolute; left:124px; top:284px; width:72px; height:153px; z-index:1; visibility: hidden">
</div>
<table width="752" border="0" cellspacing="0" cellpadding="0" align="center" height="420">
<tr>
<td width="46%">
<table width="681" border="0" cellspacing="0" cellpadding="0" align="center" height="356">
<tr>
<td width="681">
<div align="left" class="unnamed1"><div align="center">
<input onMouseOut="this.style.backgroundColor='#FFF7EE'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: #FFF7EE; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=动作 type="button" name=gb onClick="MM_showHideLayers('lt','','hide','ssx','','hide','xx','','hide','js','','hide','hysc','','hide','qcd','','hide','du','','hide','tud','','show','han','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">
</div></div>
</td>
</tr>
<tr>
<td height="13" width="681">&nbsp; </td>
</tr>
<tr>
<td width="681">

</td>
</tr>
<tr>
<td height="51" width="681">
&nbsp;&nbsp; <input onMouseOut="this.style.backgroundColor='#B8B4B0'" onMouseOver="this.style.backgroundColor='#FFCC33'" style="BACKGROUND-COLOR: ##B8B4B0; BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid; COLOR: #986830" value=动作 type="button" name=lt1 onClick="MM_showHideLayers('lt','','show','ssx','','hide','xx','','hide','js','','hide','tud','','hide','du','','hide');MM_displayStatusMsg('ffff');return document.MM_returnValue">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <!--[if gte vml 1]><v:shapetype id="_x0000_t136"
 coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e">
 <v:formulas>
  <v:f eqn="sum #0 0 10800"/>
  <v:f eqn="prod #0 2 1"/>
  <v:f eqn="sum 21600 0 @1"/>
  <v:f eqn="sum 0 0 @2"/>
  <v:f eqn="sum 21600 0 @3"/>
  <v:f eqn="if @0 @3 0"/>
  <v:f eqn="if @0 21600 @1"/>
  <v:f eqn="if @0 0 @2"/>
  <v:f eqn="if @0 @4 21600"/>
  <v:f eqn="mid @5 @6"/>
  <v:f eqn="mid @8 @5"/>
  <v:f eqn="mid @7 @8"/>
  <v:f eqn="mid @6 @7"/>
  <v:f eqn="sum @6 0 @5"/>
 </v:formulas>
 <v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"
  o:connectangles="270,180,90,0"/>
 <v:textpath on="t" fitshape="t"/>
 <v:handles>
  <v:h position="#0,bottomRight" xrange="6629,14971"/>
 </v:handles>
 <o:lock v:ext="edit" text="t" shapetype="t"/>
</v:shapetype><v:shape id="_x0000_s1031" type="#_x0000_t136" style='width:77.25pt;
 height:75pt' fillcolor="yellow" stroked="f">
 <v:fill color2="#f93" angle="-135" focusposition=".5,.5" focussize="" focus="100%"
  type="gradientRadial">
  <o:fill v:ext="view" type="gradientCenter"/>
 </v:fill>
 <v:shadow on="t" color="silver" opacity="52429f"/>
 <v:textpath style='font-family:"华文行楷";font-size:20pt;v-text-kern:t' trim="t"
  fitpath="t" string="唯钢刀月&#13;&#10;我柔光下&#13;&#10;独问见舞&#13;&#10;尊世影剑"/>
</v:shape><![endif]--><![if !vml]><img border=0 width=106 height=103
src="fengxue.files/image001.gif" alt="唯钢刀月&#13;&#10;我柔光下&#13;&#10;独问见舞&#13;&#10;尊世影剑"
v:shapes="_x0000_s1031"><![endif]></td>
</tr>
<tr>
<td width="681">
</td>
</tr>
<tr>
              <td height="43" width="681">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; </td>
</tr>
<tr>
<td width="681">&nbsp; </td>
</tr>
<tr>
<td width="681">&nbsp; </td>
</tr>
<tr>
<td width="681">
<div align="right"> </div>
</td>
</tr>
<tr>
<td width="681">&nbsp; </td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</div>
<table width="763" border="0" cellspacing="0" cellpadding="0" align="center" height="295">
  <tr> 
    <td width="35" height="28"></td>
    <td background="../../main/kuangbg.gif" width="451" height="38"> 
            <%
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
ip=Request.ServerVariables("REMOTE_ADDR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
sip=split(ip,".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 性别,门派 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
if rs("性别")="男" then sex="大侠"
if rs("性别")="女" then sex="女侠"
mp=rs("门派")
'查最后ip的地区
rs.close
rs.open "select * FROM z WHERE a<="& num &" and b>="&num,conn
if rs.eof or rs.bof then
	country="亚洲"
	city="未知"
else
	country=rs("c")
	city=rs("d")
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
      欢迎<%=sex%><font color=ffffff><font color=red><%=session("sjjh_name")%></font><font color="#000000">来到这里表演<%=mp%>独门剑法</font></font></td>
    <td colspan="2" height="38" rowspan="2"> </td>
  </tr>
  <tr> 
    <td background="../../main/zuobg.gif" width="35"></td>
    <td width="451" height="332"> 
      <table border="0" cellpadding="0" cellspacing="0" width="450" align="center">
        <!-- fwtable fwsrc="Untitled" fwbase="005.gif" -->
        <tr> 
          <!-- Shim row, height 1. -->
          <td><img src="shim.gif" width="450" height="1" border="0"></td>
          <td><img src="shim.gif" width="1" height="1" border="0"></td>
        </tr>
        <tr valign="top"> 
          <!-- row 1 -->
          <td>
          <img name="n005_r1_c1" src="005_r1_c1.gif" width="512" height="56" border="0"></td>
          <td><img src="shim.gif" width="1" height="56" border="0"></td>
        </tr>
        <tr valign="top"> 
          <!-- row 2 -->
          <td>
          <img name="n005_r2_c1" src="005_r2_c1.gif" width="512" height="56" border="0"></td>
          <td><img src="shim.gif" width="1" height="56" border="0"></td>
        </tr>
        <tr valign="top"> 
          <!-- row 3 -->
          <td>
          <img name="n005_r3_c1" src="005_r3_c1.gif" width="512" height="56" border="0"></td>
          <td><img src="shim.gif" width="1" height="56" border="0"></td>
        </tr>
        <tr valign="top"> 
          <!-- row 4 -->
          <td>
          <img name="n005_r4_c1" src="005_r4_c1.gif" width="512" height="56" border="0"></td>
          <td><img src="shim.gif" width="1" height="56" border="0"></td>
        </tr>
        <tr valign="top"> 
          <!-- row 5 -->
          <td>
          <img name="n005_r5_c1" src="005_r5_c1.gif" width="511" height="56" border="0"></td>
          <td><img src="shim.gif" width="1" height="56" border="0"></td>
        </tr>
        <tr valign="top"> 
          <!-- row 6 -->
          <td>
          <img name="n005_r6_c1" src="005_r6_c1.gif" width="511" height="56" border="0"></td>
          <td><img src="shim.gif" width="1" height="56" border="0"></td>
        </tr>
        <tr valign="top"> 
          <!-- row 7 -->
          <td>
          <img name="n005_r7_c1" src="005_r7_c1.gif" width="511" height="56" border="0"></td>
          <td><img src="shim.gif" width="1" height="56" border="0"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="35" height="5"></td>
    <td width="451" height="5" align="center" valign="bottom"> </td>
    <td width="41" height="5"></td>
    <td width="236" height="5">&nbsp; </td>
  </tr>
</table>
<p align="center"> </p>
</body>
</html>