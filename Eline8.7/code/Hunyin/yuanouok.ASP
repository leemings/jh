<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
name=trim(request("name"))
my=trim(request("my"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from h where a='" & name & "'and b='" & my &"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "没有你要离婚的人啊！"
	response.end
end if
fgyin=rs("d")
rs.close
rs.open "select * from 用户 where 姓名='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "没有你要离婚的人啊！"
	response.end
end if
conn.execute "update 用户 set 配偶='无' where 姓名='" & name & "'"
conn.execute "update 用户 set 配偶='无',银两=银两+"&fgyin&" where 姓名='" & my & "'"
conn.execute "delete * from h where c='离婚' and (a='" & name & "' or b='" & name & "')"
conn.execute "delete * from t where b='" & name & "' or c='" & name & "'"
message="[离婚]" & my & "与" & name & "离婚成功！" & my & "为此得到"&fgyin&"两的生活补贴，让我们为他们重获自由鼓掌！"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
says="<font color=red>【消息】</font><b><font color=#000000>'"& message &"'</font></b>"			'聊天数据
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
%>
<link rel="stylesheet" href="../style1.css">
<body bgcolor="#000000" background="../jhimg/bk_hc3w.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=message%></p>
<p align=center><a href=yuanou.asp>返回</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>
