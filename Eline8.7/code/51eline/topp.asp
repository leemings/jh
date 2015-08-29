<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0032)http://mforest.17173.com/top.htm -->
<HTML><HEAD><TITLE>≮E线江湖≯- www.51eline.com</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
%>
<STYLE type=text/css>TD {
	FONT-SIZE: 12px; LINE-HEIGHT: 18px
}
A {
	COLOR: #009300; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff0000; TEXT-DECORATION: underline
}
</STYLE>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>

<script language="JavaScript">
<!--
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("是否显示自己的物品？","江湖提示")){window.open('chat/wupin.asp','wupin','scrollbars=yes,resizable=yes,width=500,height=400');}}}
function chatWin() {
woiwo=window.open('chat/jhchat.asp','sjjh','resizable=no,scrollbars=no,toolbar=no,menubar=no,fullscreen=no');
woiwo.moveTo(0,0)
woiwo.resizeTo(screen.availWidth,screen.availHeight)
woiwo.outerWidth=screen.availWidth
woiwo.outerHeight=screen.availHeight
}
<!--

//-->//-->
</script>
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
</HEAD>
<BODY bgColor=#00a200 leftMargin=0 topMargin=0>
<DIV align=center>
<TABLE class=font cellSpacing=0 cellPadding=0 width=100 border=0>
  <TBODY>
  <TR>
    <TD><IMG height=44 src="51eline.com/top.files/index_01.jpg" width=388></TD>
    <TD><IMG height=44 src="51eline.com/top.files/index_02.jpg" 
width=387></TD></TR></TBODY></TABLE>
<TABLE class=font cellSpacing=0 cellPadding=0 width=293 border=0>
  <TBODY>
  <TR>
    <TD width=50><IMG height=148 src="51eline.com/top.files/index_03.jpg" width=50></TD>
    <TD width=429>
      <TABLE height=148 cellSpacing=0 cellPadding=0 width=429 
      background=51eline.com/top.files/index_04.jpg border=0>
        <TBODY>
        <TR>
          <TD>
            <OBJECT 
            codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0 
            height=148 width=429 
            classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000>
                    <PARAM NAME="movie" VALUE="51eline.com/top.files/banner.swf"><PARAM NAME="quality" VALUE="high"><PARAM NAME="wmode" VALUE="transparent">
                                                                            
            <embed src="51eline.com/top.files/banner.swf" quality=high 
            pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" 
            type="application/x-shockwave-flash" width="429" height="148" 
            wmode="transparent">                 </embed> 
      </OBJECT></TD></TR></TBODY></TABLE></TD>
    <TD width=246><IMG height=148 src="51eline.com/top.files/index_05.jpg" width=246></TD>
    <TD width=8><IMG height=148 src="51eline.com/top.files/index_06.jpg" 
  width=50></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=457 border=0>
  <TBODY>
  <TR>
    <TD width=50><IMG height=96 src="51eline.com/top.files/index_07.jpg" width=50></TD>
    <TD width=675>
      <TABLE class=font height=96 cellSpacing=0 cellPadding=0 width=675 
      background=51eline.com/top.files/index_08.jpg border=0>
        <TBODY>
        <TR>
                <TD><FONT color=#009300>| <A 
            onmouseover="MM_swapImage('Image136','','jpg/index_logo_r2_c6_1.jpg',1)" 
            title=进入江湖聊天室闯闯 onclick=chatWin() onmouseout=MM_swapImgRestore() 
            href="neweline.asp#" target=_self>杀入江湖</A> | <A href="bbs/z_visual.asp" target=_blank>我的形象</A> 
                  | <a target="_blank" href="work/changework.asp">选择职业</A> | <a href="#" onClick="window.open('jhmp/money.asp','money','scrollbars=no,resizable=no,width=430,height=340')">领取工职</A> | <a href="#" title="非掌门武功排行" onClick="window.open('top/top.htm','','scrollbars=no,resizable=yes,width=780,height=400')">江湖排行</a> | <a href="#" onClick="window.open('jl/jlmain.asp','','scrollbars=yes,resizable=yes,width=740,height=450')">江湖大事</A> |<BR>
                  | <a href="#" onClick="window.open('hcjs/jhzb/index.asp','wupin','scrollbars=yes,resizable=yes,width=550,height=300')" title="查看一下自己有什么物品">我的物品</A> 
                  | <a href="#" onClick="window.open('hcjs/jhzb/myhua.asp','myhua','scrollbars=no,resizable=no,width=430,height=340')">我的鲜花</A> 
                  | <a href="JHMP/MP5.ASP" target="_blank">我的徒弟</A> | <a href="zcmp/mmbh.asp" target="_blank">密码保护</A> 
                  | <a href="yamen/modify.asp" target="_blank">修改密码</a> | <A href="yamen/regmodi.asp" 
            target=_blank>修改档案</A> |<BR>
                  | <a href="#" onClick="window.open('diary/diary.asp','','scrollbars=no,resizable=no,width=734,height=400')">心情日记</A> 
                  | <A 
            href="zh/flzh.ASP" 
            target=_blank>法力转换</A> | <A 
            href="zh/qgzh.ASP" 
            target=_blank>轻功转换</A> | 
			<a href="#"onClick="window.open('jhmp/maidou.asp','dou','scrollbars=yes,resizable=yes,width=320,height=340')" title="卖豆子！">卖威力豆</A> | 
			<a href="#"onClick="window.open('hunyin/killer.asp','','scrollbars=no,resizable=no,width=700,height=400')">聘请杀手</A> | 
			<a href="exit.asp" title=安全离开江湖 style="text-decoration: none">离开江湖</A> |</FONT></TD>
              </TR></TBODY></TABLE></TD>
    <TD width=10><IMG height=96 src="51eline.com/top.files/index_09.jpg" 
  width=50></TD></TR></TBODY></TABLE></DIV></BODY></HTML>
