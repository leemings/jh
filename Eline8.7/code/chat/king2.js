// =============================================================
tBack=new Array(6);
tOut=new Array(6);
function menuOut(whichMenu){
var curMenu=eval("menu"+whichMenu);
	curMenu.runtimeStyle.zIndex=6;
	clearTimeout(tBack[whichMenu]);
	moveOut(whichMenu);
}
function menuBack(whichMenu){
var curMenu=eval("menu"+whichMenu);
	curMenu.runtimeStyle.zIndex=curMenu.style.zIndex;
	clearTimeout(tOut[whichMenu]);
	moveBack(whichMenu);
}
function moveOut(curNum){
var	curMenu=eval("menu"+curNum);
	if(curMenu.style.posLeft>0) {
		curMenu.style.posLeft-=2;
		tOut[curNum]=setTimeout("moveOut('"+curNum+"')",1);
		}
}
function moveBack(curNum){
var	curMenu=eval("menu"+curNum);
	if(curMenu.style.posLeft<30) {
		curMenu.style.posLeft+=2;
		tBack[curNum]=setTimeout("moveBack('"+curNum+"')",1);
		}

}
function swapClass(){
var o=event.srcElement;
	if(o.className=="td1") o.className="td2"
		else if(o.className=="td2") o.className="td1";
}
document.onmouseover=swapClass;
document.onmouseout=swapClass;
// =============================================================
document.write("<STYLE type=text/css>")
document.write(".alpha{FILTER: Alpha(Opacity=80);}")
document.write(".td1{FONT-SIZE: 12px; FONT-FAMILY: verdana;}")
document.write(".td2{background-color: #FF66FF; height: 20px; width:33%; border: #FFFFFF; border-style: dotted; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}")
document.write(".maskl{OVERFLOW: hidden;}")
document.write(".cardtitle{BORDER-RIGHT: #222222 1px solid; BORDER-TOP: #222222 1px solid; FONT-SIZE: 12px; BORDER-LEFT: #222222 0px solid; CURSOR: default; TEXT-INDENT: 4pt; BORDER-BOTTOM: #222222 0px solid; FONT-FAMILY: verdana;}")
document.write(".dfont {FONT-SIZE: 8pt; FONT-FAMILY: Webdings;color:#222222;}")
document.write(".cardbottom{BORDER-RIGHT: #222222 1px solid; BORDER-TOP: #222222 0px solid; FILTER: Alpha(Opacity=90); BORDER-LEFT: #222222 1px solid; BORDER-BOTTOM: #222222 1px solid; BACKGROUND-COLOR: #0068FF;}")
document.write("</STYLE>")

document.write("<DIV class=maskl id=menuPos style=\"Z-INDEX: 10; LEFT: -5px; WIDTH: 132px; POSITION: relative; TOP: 5px; HEIGHT: 160px;\">")
document.write("<DIV id=menu1 onmouseover=menuOut(1) style=\"Z-INDEX: 1; LEFT: 30px; WIDTH: 130px; POSITION: absolute; TOP: 0px; HEIGHT: 20px;\" onmouseout=menuBack(1)>")
document.write("<DIV class=cardbottom id=Layer1 style=\"Z-INDEX: 1; LEFT: 0px; WIDTH: 100px; POSITION: absolute; TOP: 17px; HEIGHT: 90px;\">")
document.write("<TABLE height=100% cellSpacing=0 cellPadding=0 width=100% align=center border=0>")
document.write("<TBODY><TR>")
document.write("<TD height=2></TD></TR>")
document.write("<TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")

document.write("<TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")

document.write("<TR>")
document.write("<TD height=4></TD></TR>")
document.write("</TBODY></TABLE>")
document.write("</DIV>")
document.write("<TABLE cellSpacing=0 cellPadding=0 width=100 border=0>")
document.write("<TBODY>")

document.write("<TR>")
document.write("<TD width=14 height=18><IMG height=21 src=stan.gif width=14></TD>")
document.write("<TD class=cardtitle width=90 bgColor=#0068FF height=14><div class=p5>&nbsp; <font color=#00ff00>公共菜单  <font class=dfont>6</font></font></div></TD></TR></TBODY></TABLE>")
document.write("</DIV>")
document.write("<DIV id=menu2 onmouseover=menuOut(2) style=\"Z-INDEX: 1; LEFT: 30px; WIDTH: 130px; POSITION: absolute; TOP: 20px; HEIGHT: 20px\" onmouseout=menuBack(2)>")
document.write("<DIV class=cardbottom id=Layer2 style=\"Z-INDEX: 1; LEFT: 0px; WIDTH: 100px; POSITION: absolute; TOP: 17px; HEIGHT: 90px\">")
document.write("<TABLE height=100% cellSpacing=0 cellPadding=0 width=100% align=center border=0>")
document.write("<TBODY><TR>")
document.write("<TD height=2></TD></TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")

document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TD height=4></TD></TR>")
document.write("</TBODY></TABLE></DIV>")
document.write("<TABLE cellSpacing=0 cellPadding=0 width=100 border=0>")

document.write("<TBODY>")
document.write("<TR>")
document.write("<TD width=14 height=18><IMG height=21 src=STAN.gif width=14></TD>")
document.write("<TD class=cardtitle width=90 bgColor=#0068FF height=14><div class=p5>&nbsp; <font color=#00ff00>打斗菜单  <font class=dfont>6</font></font></div></TD></TR></TBODY></TABLE>")
document.write("</DIV>")
document.write("<DIV id=menu3 onmouseover=menuOut(3) style=\"Z-INDEX: 1; LEFT: 30px; WIDTH: 130px; POSITION: absolute; TOP: 40px; HEIGHT: 20px\" onmouseout=menuBack(3)>")
document.write("<DIV class=cardbottom id=Layer3 style=\"Z-INDEX: 1; LEFT: 0px; WIDTH: 100px; POSITION: absolute; TOP: 17px; HEIGHT: 90px\">")
document.write("<TABLE height=100% cellSpacing=0 cellPadding=0 width=100% align=center border=0>")
document.write("<TBODY><TR>")
document.write("<TD height=2></TD></TR>")

document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD>")
document.write("<TD class=td1 width=33% align=middle onClick=\"document.location = 'http://www.camellia.com.tw/bbs/list.asp?boardid=1';\">菜单</TD></TR>")
document.write("<TR>")
document.write("<TD height=4></TD></TR>")

document.write("</TBODY></TABLE></DIV>")
document.write("<TABLE cellSpacing=0 cellPadding=0 width=100 border=0>")
document.write("<TBODY>")
document.write("<TR>")
document.write("<TD width=14 height=18><IMG height=21 src=stan.gif width=14></TD>")
document.write("<TD class=cardtitle width=90 bgColor=#0068FF height=14><div class=p5>&nbsp; <font color=#00ff00>个人菜单  <font class=dfont>6</font></font></div></TD></TR>")
document.write("</TBODY></TABLE>")
document.write("</DIV>")

document.write("</DIV>")