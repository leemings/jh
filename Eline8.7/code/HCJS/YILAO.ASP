<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<head>
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("sjjh_chatroomname")%>蝴蝶谷♀wWw.51eline.com♀</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>

<body background="../images/bg.gif" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 体力 from 用户 where 姓名='"&sjjh_name&"'" ,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "对不起，不能进行操作！你还不是江湖中人，滚。"
window.close()
</script>
<%
else
sm=rs("体力")
conn.close
set conn=nothing
set rs=nothing

%>
<font color="#000000">欢迎<%=sjjh_name%>来蝴蝶谷治疗（医疗事故不负责任） </font> </div>
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table>
<form method=POST action='yilao2.asp' onsubmit="this.ok.disabled=true">
<tr>
<td>胡医仙说：你的生命力为<%=sm%>。
<%
if sm>=1000 then response.write "不需要治疗！":p=1
if sm<1000 and sm>0 then response.write "建议多吃些补品！":p=2
if sm<0 and sm>-500 then response.write "你身体有恙，建议治疗！":p=3
if sm<-500 then response.write "你伤得不轻，如不治疗将危及生命！":p=4
%>
<tr>
<td>治疗方法：<input name="yilao" readonly size="8"
value="<%
if p=1 then response.write "无"
if p=2 then response.write "吃些补品"
if p=3 then response.write "一般治疗"
if p=4 then response.write "抢救生命"
%>"> <input name="ok" type="submit" value="确定"> <input type="button"
value="返回" onclick="location.href='../jh.asp'"></td>
</tr>
<tr>
<td colspan="2" style="font-size:9pt">
<hr>
全部的治疗的目标都是将生命调到1000。<br>
1、滋补主要是针对生命值在0--1000间的，每两银子加生命值1.5<br>
2、一般治疗是针对生命值在-500--0间的，每两银子加生命值1.0<br>
3、全力抢救是针对生命值在-500以下的，每两银子加生命值0.5。</td>
</tr>
</form>
</table>
</td>
</tr>
</table>
<%end if%>
<div align="center"><%=Application("sjjh_chatroomname")%>版权所有</div>
</body>
</html>
