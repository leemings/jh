<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="const.asp"-->
<%
Response.Expires=0
Response.charset="gb2312"
%>
<HTML><HEAD>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<META content="MSHTML 6.00.2800.1126" name=GENERATOR>
<style>
body{
CURSOR: url('56.ani');
scrollbar-face-color:#ffffff;
scrollbar-shadow-color:#999999;
scrollbar-highlight-color:#999999;
scrollbar-3dlight-color:#ffffff;
scrollbar-darkshadow-color:#ffffff;
scrollbar-track-color:#ffffff;
scrollbar-arrow-color:#999999;
}
</style>
</HEAD>
<BODY style="FONT-SIZE: 10pt; FONT-FAMILY: 宋体" text=#000000 bgColor=#bfbfbf onload="dspltip()" onunload="endf2()">
<DIV style="LEFT: 10px; WIDTH: 58px; POSITION: absolute; TOP: 10px; HEIGHT:200px; BACKGROUND-COLOR: #808080; TEXT-ALIGN: center">
<IMG style="POSITION: relative; TOP: 18px" height=31 src="light.gif" 
width=23 border=0><BR>

<SPAN id=cntr style="FONT-WEIGHT: bold; FONT-SIZE: 10pt; COLOR: #fffff0; FONT-FAMILY: Arial; POSITION: relative; TOP: 140px"></SPAN></DIV>
<DIV style="BORDER-RIGHT: #808080 1px solid; BORDER-TOP: #808080 1px solid; LEFT: 68px; BORDER-LEFT: #808080 1px solid; WIDTH: 346px; BORDER-BOTTOM: #808080 1px solid; POSITION: absolute; TOP: 10px; HEIGHT: 200px; BACKGROUND-COLOR: #ffffff">
<BR>&nbsp;&nbsp;
<FONT style="FONT-SIZE: 10pt" face=宋体 color=#0000ff>
<B>每日一贴 － 快乐江湖</B></FONT><br><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 你知道吗? 
<HR style="POSITION: relative; TOP: 10px" SIZE=1>
<SPAN id=mtxt style="PADDING-RIGHT: 10px; PADDING-LEFT: 10px; FONT-SIZE: 9pt; PADDING-BOTTOM: 10px; OVERFLOW: auto; WIDTH: 100%; PADDING-TOP: 15px; FONT-FAMILY: Arial; POSITION: relative; TOP: 1px; HEIGHT: 118px"></SPAN></DIV>
<DIV style="LEFT: 80px; WIDTH: 100%; POSITION: absolute; TOP: 209px">
<INPUT id=dspl1 type=checkbox CHECKED>&nbsp;<SPAN onclick=chgstatus(1)>登陆江湖显示本页</SPAN>
<INPUT id=dspl2 type=checkbox CHECKED>&nbsp;<SPAN onclick=chgstatus(2)>登陆江湖播放音乐</SPAN>
</DIV>
<DIV style="LEFT: 200px; WIDTH: 100%; POSITION: absolute; TOP: 232px">
<INPUT style="BORDER-RIGHT: #ffffff 2px outset; BORDER-TOP: #ffffff 2px outset; FONT-SIZE: 9pt; LEFT: 0px; BORDER-LEFT: #ffffff 2px outset; WIDTH: 70px; COLOR: #000000; BORDER-BOTTOM: #ffffff 2px outset; FONT-FAMILY: Arial; POSITION: relative; BACKGROUND-COLOR: #bfbfbf" onclick=ntip() type=button value=下一贴>&nbsp;&nbsp;<INPUT style="BORDER-RIGHT: #ffffff 2px outset; BORDER-TOP: #ffffff 2px outset; FONT-SIZE: 9pt; LEFT: 0px; BORDER-LEFT: #ffffff 2px outset; WIDTH: 70px; COLOR: #000000; BORDER-BOTTOM: #ffffff 2px outset; FONT-FAMILY: Arial; POSITION: relative; BACKGROUND-COLOR: #bfbfbf" onclick=endf() type=button value=关闭程序> 
</DIV>
<SCRIPT language=JavaScript>
todmsg=new Array();

todmsg[0]="<%=opday0%>"
todmsg[1]="<%=opday1%>"
todmsg[2]="<%=opday2%>"
todmsg[3]="<%=opday3%>"
todmsg[4]="<%=opday4%>"
todmsg[5]="<%=opday5%>"
todmsg[6]="<%=opday6%>"
todmsg[7]="<%=opday7%>"
todmsg[8]="<%=opday8%>"

</SCRIPT>

<SCRIPT language=JavaScript>
cnt=0;
todmsg.sort(rndm)

function rndm(a,b){
return (Math.random()-Math.random())
}

function chgstatus(){
if (dspl1.checked)
dspl1.checked=false;
else
dspl1.checked=true;
}
function chgstatus(){
if (dspl2.checked)
dspl2.checked=false;
else
dspl2.checked=true;
}

function dspltip(){
mtxt.innerHTML=todmsg[0]
cntr.innerHTML="1/"+todmsg.length
}

function ntip(){
if (cnt<todmsg.length-1)
cnt++;
else
cnt=0;
mtxt.innerHTML=todmsg[cnt];
cntr.innerHTML=(cnt+1)+"/"+todmsg.length
}
function endf(){
endf2();
window.close();
}
function endf2(){
var Arrset = "1";
if (!dspl1.checked) Arrset="0";
if (!dspl2.checked) 
Arrset+="0";
else
Arrset+="1";
window.returnValue=Arrset
}
//alert(window.dialogArguments)
</SCRIPT>
</BODY></HTML>
