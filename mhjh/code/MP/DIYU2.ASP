
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=session("yx8_mhjh_username")
yilao=request.form("yilao")
if yilao<>"无" then
set conn=server.createobject("adodb.connection")
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 资质 FROM 用户 WHERE 姓名='" & name & "' "
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
mess="你不能进行修炼，还不够资格哦！"
else
select case yilao
case "一级炼火"
cd=10000
case "二级炼火"
cd=20000
case "三味真火"
cd=100000
end select
if cd>rs("资质") then
mess="你的资质这么差劲，还想炼火哦？！先去加点资质吧！懒虫！！！"
else
sql="update 用户 set 杀人=杀人-1,资质=资质-" & cd & " where 姓名='" & name & "'"
conn.execute sql
mess="经过地狱的无情之火煅烧，不断反省自己以前的过失，哎~~恩怨宜解不宜结！杀人不过头点地，何必要人命？？？你耗费了" & cd & "点资质，终于领悟到了玩江湖的真缔！罪恶减轻了一个点数！"
end if
end if
conn.close
set conn=nothing
else
mess="你这个通缉犯，再不来炼狱就要进入无穷的痛苦轮回了！"
end if
%>

<head>
<link rel="stylesheet" href="../style.css">
<title>地狱炼火</title>
</head>

<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false>
<table width="332" border="0" cellspacing="0" cellpadding="0" align="CENTER">
<tr>
<td height="223" width="603">
<table width="494" border="0" cellspacing="0" cellpadding="0" height="100%">
<tr>
<td width="7" ></td>
<td  width="499" align="center"><br>
<div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="374" align="center" bordercolorlight="#cccc99" bordercolordark="#FFFFFF">
<tr>
<td width="364">
<p align=center><%=mess%></p><br>
<p align=center><a href="diyu.asp">返回</a></p>
</td>
</tr>
</table>
</div>
</td>
<td width="5" ></td>
</tr>
</table>
</td>
</tr>
</table>
</body>

