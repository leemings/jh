<%
username=session("aqjh_name")
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rst=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 where 姓名='" &username& "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
response.write "你不是江湖中人或者连接超时"
conn.close
response.end
else
hy=rs("会员")
hydj=rs("会员等级")
tl=rs("体力")
nl=rs("内力")
gj=rs("攻击")
fy=rs("防御")
if username="" then
%>
<script language=vbscript>
MsgBox "对不起，你还没有登录"
location.href ="../../exit.asp"
</script>
<%
else
if Application("aqjh_hy")=true then
%>
<script language=vbscript>
MsgBox "错误！会员功能尚未开放！"
location.href ="javascript:history.back()"
</script>
<%
else
if hy=false and hydj=0 then
%>
<script language=vbscript>
MsgBox "错误！你不是会员，请回吧！"
location.href ="javascript:history.back()"
</script>
<%
else
sj=DateDiff("n",rs("操作时间"),now())
if sj<10 then
ss=10-sj
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('你刚刚才来过怎么又来啦？请等："&ss&"分后再来吧！');location.href = 'javascript:history.back()';}</script>"
Response.End 
end if
jiu=request.form("jiu")
Jname=session("aqjh_name")
select case jiu
case "lao"
tili=40
yin=20000
case "wu"
tili=60
yin=30000
case "du"
tili=80
yin=40000
case "mao"
tili=100
yin=50000
end select
%>
<!--#include file="data.asp"-->
<%
sql="select * from 用户 where 姓名='" & Jname & "' and 状态<>'死亡' and 状态<>'入狱' and 状态<>'逮捕'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess=Jname & "，你不能操作！"
else
if rs("银两")<yin then
mess=Jname & "，你的银两不够！"
else
daode=rs("道德")
jiuliang=int((daode)/100)
if jiuliang>1 then
sql="update 用户 set 银两=银两-" & yin & ",道德=道德+"& tili &" where 姓名='" & Jname & "'"
conn.execute sql
mess=Jname & "从相国寺出来后，觉得自己得益非浅，同时也悟出不少道理!"
else 
sql="update 用户 set 银两=银两-" & yin & ",道德=道德+"& tili &",登录=now(),状态='眠' where 姓名='" & Jname & "'"
conn.execute sql
mess=Jname & "，由于你和寺内助持谈的投机，于是留你在寺内吃了饭再走，请在1小时后使用该帐号！"
response.cookies("Jname")=""
response.cookies("Jpai")=""
conn.close
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.asp target=_top"
</script>
<%
end if
end if
end if
%>
<head>
<link rel="stylesheet" href="../../css.css" type="text/css">
</head>
<body bgcolor="#000000">
<table border="1" bgcolor="#00A200" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#005b00">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="index.asp">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
<%
end if
end if
end if
end if
%>