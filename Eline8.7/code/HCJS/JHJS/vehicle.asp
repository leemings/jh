<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<HTML><HEAD><TITLE><%=Application("sjjh_chatroomname")%>交通工具专卖店</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
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
<p align="center">[ <font color="#FF3300" size=4><%=Application("sjjh_chatroomname")%>交通工具专卖店</font> ]</p>
<font color="#FF00FF" size=4><a href="myvh.asp" target="_self" title="装备交通工具及定制进入退出聊天室公告">点击这儿进入[装备座驾]页面</a></font><br><br>
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center">
<tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
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
<br><font color="#FF00FF" size=4><a href="myvh.asp" target="_self"  title="装备交通工具及定制进入退出聊天室公告">点击这儿进入[装备座驾]页面</a></font><br><br>
<font color="#ff00ff" size=4>E线江湖</font></div>

</div>
</body>
</html>
