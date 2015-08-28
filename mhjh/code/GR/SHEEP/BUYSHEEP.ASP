
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
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

if session("yx8_mhjh_username")="" then
response.redirect"warning.htm"
else
sheepname=request.form("sheepname")
if sheepname="" or len(sheepname)>16 then
response.redirect"warning.htm"
else

set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
rs.open"select * from 用户 where 姓名='"&session("yx8_mhjh_username")&"'",conn,1,1
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
conn.execute("update 用户 set 银两='"&money&"'  where 姓名='"&session("yx8_mhjh_username")&"'")
rs.close%>
<!--#include file="data1.asp"-->

<%rs.open"select * from sheep where user='"&session("yx8_mhjh_username")&"'",conn,1,1
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
msgbox"剑侠里面的孩子已经超过了10000只，你不能再领养",0,"FLUSH"
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
conn.execute"insert into sheep(sheephappy,sheephealth,user,sheepname,sheepdate,life,milk,hungry,workload,clean,feeddate,feedsheepday,logintoday) values('"&sheephappyinit&"','"&sheephealthinit&"','"&session("yx8_mhjh_username")&"','"&sheepname&"','"&sheepdate&"','"&lifeinit&"','"&milkinit&"','"&hungryinit&"','"&workload&"','"&cleaninit&"','"&feeddate&"','0','"&feeddate&"')"
rs.open"select * from sheep",conn,1,1
rs.movelast
sheepid=rs("id") '小羊编号
conn.execute"update sheep set user='"&session("yx8_mhjh_username")&"' where sheepname='"&sheepname&"'"
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>剑侠情缘</title>
<link rel="StyleSheet" href="../../style.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0" background='../../chatroom/bg1.gif'>

<div align="center">
</div>
<div align="center">
  <table border="0" width="470" bordercolor="#FFFFFF" cellspacing="1"
  cellpadding="0" height="20">
    <tr>
      <td class="p6"><font color="#000000">□-您当前的位置--<a
        href="indexsheep.asp"></font><font color="#0000FF">孤儿院</font><font color="#000000">></font></a><font color="#000000">-</font><a
        href="feedsheep.asp"><font color="#0000FF">托儿所</font></a></td>
    </tr>
  </table>
</div>
<div align="center">
  <center>
  <table border="0" width="470" cellspacing="1" cellpadding="0" height="50">
    <tr>
      <td class="p2" width="100%">
        <p align="center"><font color="#000000">恭喜！您已经成功领养了<%=sheepname%></font></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <table border="0" width="470" cellspacing="1" cellpadding="0">
    <tr>
      <td class="p3" width="100%">
        <p align="center"><font color="#000000">快回</font><a
        href="feedsheep.asp"><font color="#0000FF">托儿所</font></a><font color="#000000">开始照顾<%=sheepname%>吧！</font></td>
    </tr>
  </table>
</div>
<div align="center">
  <center></center>
</div>

<p align="right">　</p>

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















