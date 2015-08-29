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


jiu=request.form("jiu")
Jname=sjjh_name
select case jiu
case "lao"
tili=50
yin=250
case "wu"
tili=60
yin=300
case "du"
tili=70
yin=350
case "mao"
tili=80
yin=400
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'" & " and 状态<>'死' and 状态<>'狱' and 状态<>'练'",conn
if rs.eof or rs.bof then
mess=Jname & "，你不能操作！"
else
if rs("银两")<yin then
mess=Jname & "，你的银两不够！"
else
wugong=rs("武功")
jiuliang=int((wugong)/100)
if jiuliang>1 then
conn.execute "update 用户 set 银两=银两-" & yin & ",体力=体力+"& tili &" where 姓名='"&sjjh_name&"'"
mess=Jname & "好酒量!你喝完一杯,不停的赞道:好酒,好酒!"
else	
shi=0.0416*1
conn.execute "update 用户 set 银两=银两-" & yin & ",体力=体力+"& tili &",登录=now()+" & shi & ",状态='眠' where 姓名='"&sjjh_name&"'"
mess=Jname & "，由于您酒量不够,趴在桌上就睡着了,几个伙计将您抬到客栈去休息了，请在1小时后使用该帐号！"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.asp"
</script>
<%
end if
end if
end if
%>

<head>
<style>td           { font-size: 14 }
</style>
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
<p align="center" style="font-size:14;color:red"><img border="0" src="../../chat/img/pic5.gif"><%=mess%></p>
<p align="center"><a href="pub.asp">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>

