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
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
%>
<HTML><HEAD><TITLE><%=Application("aqjh_chatroomname")%>交通工具专卖店</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK href="../../css.css" type=text/css rel=stylesheet>
<style type="text/css">
<!-- 
a { text-decoration: none} 
--> 
A:link {COLOR: #ff00ff;FONT-FAMILY: 宋体; }
A:visited {COLOR: #ff00ff; FONT-FAMILY: 宋体; }
A:active {FONT-FAMILY: 宋体; }
A:hover {FONT-FAMILY: 宋体;COLOR: #FF0000; }
BODY {FONT-FAMILY: 宋体; FONT-SIZE: 10pt;COLOR: #ffffff;}
TABLE {FONT-FAMILY: 宋体; FONT-SIZE:10pt;COLOR : #ffffff;}
</style>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR></HEAD>
<BODY bgColor=#85C2E0 text=#ffffff topMargin=5>
<DIV align=center>
<p align="center">[ <font color="#FF3300"><%=Application("aqjh_chatroomname")%>交通工具专卖店</font> ]</p>
<font color="#FF00FF"><a href="myvh.asp" target="_self" title="装备交通工具及定制进入退出聊天室公告">点击这儿进入[装备座驾]页面</a></font><br><br>
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM b WHERE b='交通'",conn
hh=-1
do while not rs.eof and not rs.bof 
hh=hh+1
ha=hh mod 4
if ha<>0 then
%>
  <td width=25% align="center" valign="bottom" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';"> 
  <img src="images/<%=rs("i")%>"><br><br>[<%=rs("a")%>]　　耐久度[<%=rs("j")%>]<br>银两:<%=rs("h")%>两，金币:<%=rs("m")%>个<br>条件:[<%=rs("l")%>]　　<a href="buyVH.asp?wpname=<%=rs("a")%>"><font color="#0000FF">购买</font></a> 
  </td>
<% else %>
 </tr><tr>
  <td width=25% align="center" valign="bottom" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';"> 
  <img src="images/<%=rs("i")%>"><br><br>[<%=rs("a")%>]　　耐久度[<%=rs("j")%>]<br>银两:<%=rs("h")%>两，金币:<%=rs("m")%>个<br>条件:[<%=rs("l")%>]　　<a href="buyVH.asp?wpname=<%=rs("a")%>"><font color="#0000FF">购买</font></a> 
  </td>
<%end if
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</tr>
</table>
</div>
<div align="center">
<br><font color="#FF00FF"><a href="myvh.asp" target="_self"  title="装备交通工具及定制进入退出聊天室公告">点击这儿进入[装备座驾]页面</a></font><br><br>
<font color="#ff00ff"><FONT color=#0000ff>&copy; 版权所有 2004-2005 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>快乐江湖网</FONT></A></font></div>
</div></body></html>