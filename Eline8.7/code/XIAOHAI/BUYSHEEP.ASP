
<%
set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
function ccdate(sdate)
temp=cdate(sdate)
if len(month(temp))=1 then
m="0"&month(temp)
else
m=month(temp)
end if

if len(day(temp))=1 then
d="0"&day(temp)
else
d=day(temp)
end if
ccdate=year(temp)&"-"&m&"-"&d
end function

function cctime (stime)
if len(hour(stime))=1 then
h="0"&hour(stime)
else
h=hour(stime)
end if
if len(minute(stime))=1 then
m="0"&minute(stime)
else
m=minute(stime)
end if
if len(second(stime))=1 then
s="0"&second(stime)
else
s=second(stime)
end if
cctime=h&":"&m&":"&s
end function
'=====================================================

if session("sjjh_name")="" then
response.redirect"warning.htm"
else
sheepname=request.form("sheepname")
if sheepname="" or len(sheepname)>16 then
response.redirect"warning.htm"
else

set conn=server.createobject("adodb.connection") 
conn.open Application("sjjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
rs.open"select * from 用户 where 姓名='"&session("sjjh_name")&"'",conn,1,1
if rs("银两")<100000 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"您没有足够的钱！",0,"FLUSH"
history.back
</script>
<%
else
money=rs("银两")-100000
conn.execute("update 用户 set 银两='"&money&"'  where 姓名='"&session("sjjh_name")&"'")
rs.close%>
<!--#include file="data1.asp"-->

<%rs.open"select * from sheep where user='"&session("sjjh_name")&"'",conn,1,1
if not rs.bof then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"一个人只能领养一个小孩",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select * from sheep",conn,1,1
if rs.recordcount>10000 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"剑侠里面的孩子已经超过了10000个，你不能再领养",0,"FLUSH"
history.back
</script>
<!--#include file="data1.asp"-->
<%
else
rs.close
rs.open"select * from rules",conn,1,1
cleaninit=rs("cleaninit")
sheephappyinit=rs("happyinit")
sheephealthinit=rs("healthinit")
milkinit=rs("milkinit")
lifeinit=rs("lifeinit")
hungryinit=rs("hungryinit")
sheepdate=date
feeddate=DATEADD("d",-1,DATE)
workload=0
rs.close
conn.execute"insert into sheep(sheephappy,sheephealth,user,sheepname,sheepdate,life,milk,hungry,workload,clean,feeddate,feedsheepday,logintoday) values('"&sheephappyinit&"','"&sheephealthinit&"','"&session("sjjh_name")&"','"&sheepname&"','"&sheepdate&"','"&lifeinit&"','"&milkinit&"','"&hungryinit&"','"&workload&"','"&cleaninit&"','"&feeddate&"','0','"&feeddate&"')"
rs.open"select * from sheep",conn,1,1
rs.movelast
sheepid=rs("id") '孩子编号
conn.execute"update sheep set user='"&session("sjjh_name")&"' where sheepname='"&sheepname&"'"
%>

<html>

<head>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
body {font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt; color:'ffff77';
	scrollbar-face-color:'222222'; 
	scrollbar-shadow-color:'222222'; 
	scrollbar-highlight-color:'EFBA56'; 
	scrollbar-3dlight-color:'222222'; 
	scrollbar-darkshadow-color: 'AB7510'; 
	scrollbar-track-color:'FBEED7'; 
	scrollbar-arrow-color:'AB7510' ;

	}

td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:F7DEAC}
option	{font-size: 9pt; background-color:F7DEAC}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'ffff77'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff4d4d;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: '291C03'; border: 1 solid '291C03'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-RIGHT: 1px; PADDING-LEFT: 1px; 
	BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'EBA82C'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-RIGHT-WIDTH: 1px; 
	BORDER-RIGHT-COLOR:'EBA82C'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: '333333'; 
	border: 333333px solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'F3CB81'; color: '83590C'; 
	border: 1px solid AB7510;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'EBA82C'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'EBA82C'}
</style>
<style>.skin0 {
	BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; VISIBILITY: hidden; BORDER-LEFT: 1px solid; WIDTH: 80px; CURSOR: default; LINE-HEIGHT: 20px; BORDER-BOTTOM: 1px solid; FONT-FAMILY: Verdana; POSITION: absolute; BACKGROUND-COLOR: white
}
.menuitems {
	PADDING-RIGHT: 10px; PADDING-LEFT: 15px
}
                </style>



<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>领养孤儿</title>
</head>

<body bgcolor="#000000">

<div align="center">
  <center>
  <table border="0" width="95%" cellspacing="0" cellpadding="0" height="183" background="../picc/bg.jpg">
    <tr>
      <td width="100%" colspan="3" background="picc/k2.gif" height="18">　</td>
    </tr>
    <tr>
      <td width="4%" rowspan="4" background="picc/k3.gif" height="147">　</td>
      <td width="88%" align="center" height="26"><img border="0" src="61.gif" width="111" height="22"></td>
      <td width="4%" background="picc/k3.gif" height="147" rowspan="4">　</td>
    </tr>
    <tr>
      <td width="88%" height="70" align="center"><img border="0" src="picc/kg8.jpg" width="520" height="309"></td>
    </tr>
    <tr>
      <td width="88%" height="103" valign="top">
        <table border="0" width="100%">

          <tr>
            <td width="100%" align="center">□-您当前的位置<font color="#ffcc00">-</font><font color="#ffff00">-</font><a
        href="indexsheep.asp">孤儿院</a><font color="#ffcc00">-</font><a
        href="feedsheep.asp">托儿所</a></td>                    
          </tr>

          <tr>
            <td width="100%" align="center">恭喜！您已经成功领养了<%=sheepname%></td>
          </tr>

          <tr>
            <td width="100%" align="center">          
快回</font><a
        href="feedsheep.asp"><font color="#FFFFFF">托儿所</font></a>开始照顾<%=sheepname%>吧！</font></td>
          </tr>
         <td><hr>

          <tr>
            <td width="100%" align="center"></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td width="88%" height="18"></td>
    </tr>
    <tr>
      <td width="100%" colspan="3" background="picc/k2.gif" height="18">　</td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
 <% 
rs.close 
conn.close 
end if 
end if
end if 
end if 
end if 
%>
<script language="JavaScript">                                                
<!--                                                 
var runAs=1;                                                
N = 40;                                                 
Y = new Array();                                                
X = new Array();                                                
S = new Array();                                                
A = new Array();                                                
B = new Array();                                                
M = new Array();                                                
V = (document.layers)?1:0;                                                
iH=(document.layers)?window.innerHeight:window.document.body.clientHeight;                                                
iW=(document.layers)?window.innerWidth:window.document.body.clientWidth;                                                
for (i=0; i < N; i++){                                                 
Y[i]=Math.round(Math.random()*iH);                                                
X[i]=Math.round(Math.random()*iW);                                                
S[i]=Math.round(Math.random()*5+2);                                                
A[i]=0;                                                
B[i]=Math.random()*0.1+0.1;                                                
M[i]=Math.round(Math.random()*1+1);                                                
}                                                
if (V){                                                
for (i = 0; i < N; i++)                                                
{document.write("<LAYER NAME='sn"+i+"' LEFT=0 TOP=0 BGCOLOR='#FFFFF0' CLIP='0,0,"+M[i]+","+M[i]+"'></LAYER>")}                                                
}                                                
else{                                                
document.write('<div style="position:absolute;top:0px;left:0px">');                                                
document.write('<div style="position:relative">');                                                
for (i = 0; i < N; i++)                                                
{document.write('<div id="si" style="position:absolute;top:0;left:0;width:'+M[i]+';height:'+M[i]+';background:#fffff0;font-size:'+M[i]+'"></div>')}                                                
document.write('</div></div>');                                                
}                                                
function snow(){                                                
var H=(document.layers)?window.innerHeight:window.document.body.clientHeight;                                                
var W=(document.layers)?window.innerWidth:window.document.body.clientWidth;                                                
var T=(document.layers)?window.pageYOffset:document.body.scrollTop;                                                
var L=(document.layers)?window.pageXOffset:document.body.scrollLeft;                                                
for (i=0; i < N; i++){                                                
sy=S[i]*Math.sin(90*Math.PI/180);                                                
sx=S[i]*Math.cos(A[i]);                                                
Y[i]+=sy;                                                
X[i]+=sx;                                                 
if (Y[i] > H){                                                
Y[i]=-10;                                                
X[i]=Math.round(Math.random()*W);                                                
M[i]=Math.round(Math.random()*1+1);                                                
S[i]=Math.round(Math.random()*5+2);                                                
}                                                
if (V){document.layers['sn'+i].left=X[i];document.layers['sn'+i].top=Y[i]+T}                                                
else{si[i].style.pixelLeft=X[i];si[i].style.pixelTop=Y[i]+T}                                                 
A[i]+=B[i];                                                
}                                                
if(runAs==1){setTimeout('snow()',5);}                                                
}                                                
//-->                                                
</script>                                                
                                                
                                   
<script language=javascript>                                                                                                                     
     function Click(){                                                                                                                    
     alert('快乐江湖欢迎您！');                                                                                                                    
     window.event.returnValue=false;                                                                                                                    
     }                                                                                                                    
     document.oncontextmenu=Click;                                                                                                                    
     </script> 





