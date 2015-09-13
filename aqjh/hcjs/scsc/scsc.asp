<%@ LANGUAGE=VBScript codepage ="936" %><%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<HTML><HEAD><TITLE>生产市场</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 leftMargin=0 topMargin=0>
<DIV align=center><FONT color=#ff0000 face=华文新魏 size=5><B>欢迎</FONT><FONT 
color=#00ff00 face=华文新魏 size=5><font color="blue"><%=aqjh_name%></font></FONT><FONT color=#ff0000 face=华文新魏 
size=5>光临生产市场<IMG border=0 src="31.gif"></FONT></B></DIV>
<DIV align=center>　</DIV>
<DIV align=center><FONT color=#ffffff><A 
href="sc/zh.asp"><FONT color=#ff9900 
size=2>转换市场</FONT></A><FONT color=#ff9900 size=2>&nbsp; </FONT><A  
href="sc/mysc.asp"><FONT color=#ff9900  
size=2>自由市场</FONT></A><FONT color=#ff9900 size=2>&nbsp; </FONT><A  
href="dzwq/index.asp"><FONT color=#00ffff  
size=2>锻造武器</FONT></A><FONT color=#00ffff size=2>&nbsp; </FONT><A  
href="/peiyao.asp"><FONT color=#00ffff  
size=2>配制药品</FONT></A><FONT color=#00ffff size=2>&nbsp; </FONT><A  
href="sc/bzsc.asp"><FONT color=#00ffff  
size=2>包装市场</FONT></A></FONT></DIV> 
<DIV align=center> 
<TABLE border=1 borderColorDark=#ffffff borderColorLight=#000000 cellSpacing=1  
width=600> 
  <TBODY> 
  <TR> 
    <TD align=middle width=105><FONT color=#c0c0c0  
      style="FONT-SIZE: 9pt">物品名称</FONT></TD> 
    <TD align=middle width=172><FONT color=#c0c0c0><SPAN  
      style="FONT-SIZE: 9pt">生产类型</SPAN></FONT></TD> 
    <TD align=middle width=106><FONT color=#c0c0c0><SPAN  
      style="FONT-SIZE: 9pt">生产物品</SPAN></FONT></TD> 
    <TD align=middle width=180><FONT color=#c0c0c0  
      style="FONT-SIZE: 9pt">物品说明</FONT></TD> 
    <TD align=middle width=70><FONT color=#c0c0c0  
      style="FONT-SIZE: 9pt">开始</FONT><FONT color=#c0c0c0><SPAN  
      style="FONT-SIZE: 9pt">生产</SPAN></FONT></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="mt.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>木头===&gt;&gt;木炭</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="mutan.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>5木头</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="ks.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>矿石===&gt;&gt;黑铁矿</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="htk.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>8矿石</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="ks.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>矿石===&gt;&gt;银矿</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="yk.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>12矿石</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="ks.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>矿石===&gt;&gt;金矿</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="jk.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>16矿石</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="http://www.lingfei.com/lfjh/sc/wzw/index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="ks.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>矿石===&gt;&gt;红宝石</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="hbs.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>4矿石</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="ks.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>矿石===&gt;&gt;绿宝石</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="lvbs.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>8矿石</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="ks.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>矿石===&gt;&gt;蓝宝石</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="lbs.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>6矿石</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="ks.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>矿石===&gt;&gt;水晶石</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="sjs.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>20矿石</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR> 
  <TR> 
    <TD align=middle width=11><FONT color=#ffffff size=2><IMG border=0  
      src="lhr.gif"></FONT></TD> 
    <TD align=middle width=172><FONT color=#ffffff  
      size=2>虎肉===&gt;&gt;老虎肉</FONT></TD> 
    <TD align=middle width=106><FONT color=#ffffff size=2><IMG border=0  
      src="lhr.gif"></FONT></TD> 
    <TD align=middle width=180> 
      <P align=left><FONT color=#9bc7f2 size=2>需要材料:</FONT><FONT color=#ffff00  
      size=2>2虎肉</FONT></P></TD> 
    <TD align=middle width=70><A  
      href="index.asp"><FONT color=#ff0000  
      size=2>生产</FONT></A></TD></TR></TBODY></TABLE></DIV> 
<P align=center><font color="#c0c0c0" size="2">快乐江湖独家首创</font></P></BODY></HTML>
