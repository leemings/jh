<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<HTML><HEAD><TITLE><%=Application("aqjh_chatroomname")%>--交通工具设置</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type="text/css">
<!-- 
a { text-decoration: none} 
--> 
A:link {COLOR: #0000ff;FONT-FAMILY: 宋体; }
A:visited {COLOR: #b8860b; FONT-FAMILY: 宋体; }
A:active {FONT-FAMILY: 宋体; }
A:hover {FONT-FAMILY: 宋体;COLOR: #FF0000; }
BODY {FONT-FAMILY: 宋体; FONT-SIZE: 10pt;COLOR: #ffffff;}
TABLE {FONT-FAMILY: 宋体; FONT-SIZE:10pt;COLOR : #ffffff;}
</style>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY bgColor=#85C2E0 text=#ffffff topMargin=5>
<DIV align=center>
<p align="center">[ <font color="#FF3300"><%=Application("aqjh_chatroomname")%>--我的交通工具</font> ]</p>
<font color="#FF00FF"><a href="vehicle.asp" target="_self"  title="去购买各式各样的交通工具">点击这儿返回[交通工具专卖店]页面</a></font><br><br>
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM vh WHERE b='"&aqjh_name&"'",conn
do while not rs.eof and not rs.bof
%>
  <td width=25% align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';"> 
  <img src="images/<%=rs("h")%>"><br><%=rs("a")%>；　<font color="#ff0000">使用状况：<br></font><%if rs("c") then%>进入：[使用]<%else%> 进入：[不用]<%end if%> <%if rs("e") then%>　退出：[使用] <%else%>　退出：[不用]<%end if%></td>
  <td valign="middle"><table><tr><td colspan=2>最后日期：<%=rs("l")%>；　耐久度：<%=rs("i")%>；　状态：<%if rs("j") then%> 损坏 <%else%> 正常<%end if%>；　维修次数：<%=rs("k")%>　　<a href="rpvh.asp?id=<%=rs("id")%>" target="_self" title="维修需要支付一定的费用">[进行维修]</a> <a href="rrvp.asp?id=<%=rs("id")%>" target="_self" title="维修需要支付一定的费用">[出售旧车]</a></td></tr><tr>
  
  <td colspan=2>
  <form  method="post" action="myvhok.asp">
  <div>座驾昵称：<input type="text" name="vhname" value="<%=rs("g")%>" maxlength=50 size=18>
  <select name="jrtc" >
  <%if not rs("c") and rs("e")  then%>
  <option value=true >进入聊天室时</option><option value=false selected >退出聊天室时</option></select>
  <%else%>
  <option value=true  selected >进入聊天室时</option><option value=false>退出聊天室时</option></select>
  <%end if%>
  <select name="jr"><option value=false>不用</option><option value=true selected>使用</option></select>
  <tr>
  <td >
  <%if not rs("c") and rs("e") then%>
  公告：<input type="text" name="jrtg" value="<%=rs("f")%>" maxlength=255 size=60>
  <%else%>
  公告：<input type="text" name="jrtg" value="<%=rs("d")%>" maxlength=255 size=60>
  <%end if%>
  </td>
  <td>
  <input type="hidden" name="id" value="<%=rs("id")%>">
  <input type="submit" name="submit" value="确定"></td>
  </tr>
  </div>
  </form></td></tr></table>
  </td>
  </tr><tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing

%>
</tr>
</table>
</div><div align="left">
<br> <font color="#0000ff">
提　　示：进入时公告可以使用##代替自己的名字，$$代替座驾昵称。<br>
如进入时：##开着刚买不久的$$，来到了<%=Application("aqjh_chatroomname")%>，见到各位大虾，拱手曰：“众位大虾，小生有礼了”<br>
或退出时：##泪眼模糊对江湖里的各位大虾道：“我有事先走一步了。”话音刚落，钻进自家的$$，只见屁股冒烟，立即消失无影踪。<br>
<br>
注　　意：进入及退出公告等不可以包含半角字符及空格等，请使用全角标点符号；每种座驾只能有一个昵称。<br>
　　　　　进入或退出聊天室时最多只有一种交通工具处于使用状态，设置其中一个为使用状态会清掉其它的为非使用状态
<br>最后日期：当维修次数为0时是购买日期，否则是最后一次维修的日期</font><br><br>
<font color="#ff0000">
警    告：进入公告要是写得太无聊或无耻的话监禁ID或删ID没商量，请大家文明用语！</font> 
<br><br><DIV align=center>
<font color="#FF00FF"><a href="#" onClick="window.open('vehicle.asp',target='_self')"  title="去购买各式各样的交通工具">点击这儿返回[交通工具专卖店]页面</a></font><br><br><FONT color=#0000ff>&copy; 版权所有 2005-2006 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>快乐江湖网</FONT></A></div>
</body>
</html>