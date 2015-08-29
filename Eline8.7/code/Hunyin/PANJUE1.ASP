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
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select 配偶,等级,银两,道德 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
peiou=rs("配偶")
if peiou="无" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('【"&sjjh_name&"】你还没有讨结婚呢,逃什么婚啊！');</script>"
		Response.End
end if
if rs("等级")<15 then
Response.Write "<script language=JavaScript>{alert('江湖小辈也逃婚呀,练到15级等级再说吧！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("银两")<100000 then
Response.Write "<script language=JavaScript>{alert('你的银两不够10万呀，怎么也跑来逃婚啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("道德")<20000 then
Response.Write "<script language=JavaScript>{alert('你的道德还不够2万就学人逃婚？你想做通缉犯呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.Execute ("update 用户 set 配偶='无',道德=道德-20000,银两=银两-100000 where 姓名='" & sjjh_name &"'")
rs.close
rs.open "select 配偶 FROM 用户 WHERE 姓名='"&peiou&"'",conn
if not(Rs.Bof OR Rs.Eof) then
conn.Execute ("update 用户 set 配偶='无'  where 姓名='"&peiou&"'")
end if
message="[逃婚]" & sjjh_name & "拿出了10万银两与" & peiou & "离婚失败，结果跑出法院负了2万道德终于逃婚成功了！"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【消息】</font><b><font color=blank>'"& message &"'</font></b>"		'聊天数据
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
<body bgcolor="#000000" background="../jhimg/bk_hc3w.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=message%></p>
<p align=center><a href=taohun.asp>返回</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>
