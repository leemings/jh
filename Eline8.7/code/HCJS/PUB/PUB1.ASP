<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"


time1=request.form("time")+0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if time1=0 then
mess="请选择正确的时间！"
elseif  time1<0 then
mess="请选择正确的时间！"
elseif  time1>24 then
mess="请选择正确的时间！"
else
Jname=sjjh_name
Jpass=request.cookies("pass")
yin=abs((time1)*10)
shi=0.0416*time1
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'" & "  and 状态<>'死' and 状态<>'狱'",conn
if rs.eof or rs.bof then
mess=Jname & "，你不能操作！"
else
if rs("银两")<yin then
mess=Jname & "，你的银两不够！"
else
conn.execute "update 用户 set 银两=银两-" & yin & ",体力=体力+"& yin &",内力=内力+"&yin&",登录=now()+" & shi & ",状态='眠' where 姓名='"&sjjh_name&"'"
mess=Jname & "，你已经成功的到客栈休息了，请在" &time1& "小时后使用该帐号！"
session.Abandon
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.htm"
</script>
<%
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("sjjh_chatroomname")%>客栈</title></head>

<body bgcolor="#000000" background="../../jhimg/bk_hc3w.gif" text="#000000" link="#0000FF">
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="pub.asp">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>