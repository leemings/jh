<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('刷钱是吗？喜欢是吗？点啊，刷啊！！');i=i+1;}top.location.href='../../exit.asp'}</script>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
name=aqjh_name
yilao=request.form("yilao")
if yilao<>"无" then
'校验用户
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 姓名='"&aqjh_name&"'" &" and 体力<1000",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"分钟]再操作！');}</script>"
	Response.End
end if	

If Rs.Bof OR Rs.Eof Then
mess="你不能进行治疗"
else
sm=rs("体力")
yl=rs("银两")
select case yilao
case "吃些补品"
bl=1.5
cd=1000-sm
case "一般治疗"
bl=1
cd=1000-sm
case "抢救生命"
bl=0.5
cd=1000-sm
end select
fy=int(cd/bl)
sm=1000
if yl<fy then
fy=yl
sm=yl*bl
end if
conn.execute "update 用户 set 体力=体力+" & sm & ",银两=银两-" & fy & ",操作时间=now() where 姓名='"&aqjh_name&"'"
mess="经过胡医仙的治疗，你花了" & fy & "两银子增加生命到" & sm & "点"
end if
conn.close
set conn=nothing
else
mess="你不用治疗"
end if
%>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000">

<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="YILAO.ASP">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>
