<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
my=sjjh_name
id=request("id")
'校验用户
sql="SELECT * FROM 用户 WHERE 姓名='" & my & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
mess=my&"，你不能操作！"
else
if rs("门派")="官府" and rs("身份")>"8" then
UID=rs("ID")
sql="select * from 用户 where (状态='花柳病' or 状态='梅毒' or 状态='尿道炎' or 状态='僵尸王') and 门派<>'官府' and id=" & id
set rs=conn.execute(sql)
if not rs.eof and not rs.bof then
sql="update 用户 set 状态='正常',登录=now() where id=" & id
conn.execute sql
mess=my&"，你已经帮他把病医好了！"
pos=application("chatpos")
chats=application("chats")
sj=formatdatetime(time(),3)
sj="<font color=#FF00FF size=1>(" & sj & ")</font>"
says="<font color=red><b>" & rs("姓名") & "被神医妙手回春治好了</b>(神医ID="&UID&")</font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
else
mess="没有这个犯人或是此人是官府的人"
end if
else
mess=mye&"，你无此权力！！！"
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=Application("sjjh_chatroomname")%></title></head>

<body background="../jhimg/bk_hc3w.gif">
<div align="center">操作成功^_^<br>
</div>
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="yiyuan.asp">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>