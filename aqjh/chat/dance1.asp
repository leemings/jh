<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
name=LCase(trim(Request.QueryString("name")))
yn=LCase(trim(Request.QueryString("yn")))
to1=LCase(trim(Request.QueryString("to1")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 姓名 FROM 用户 WHERE 姓名='"&name&"'",conn
momo=rs("姓名")
if yn=1 then
	if aqjh_name=to1 then
		if Application("aqjh_online")<>aqjh_name then
			Response.Write "<script language=JavaScript>{alert('提示：网络超时……');}</script>"
			Response.End 
		end if
		Response.Write "<script language=JavaScript>{alert('我答应你的要求了.跟你语聊！');}</script>"
		conn.execute "update 用户 set 道德=道德+500,魅力=魅力-100 where 姓名='" & aqjh_name &"'"
		conn.execute "update 用户 set 道德=道德-100,魅力=魅力+500 where 姓名='"&name&"'"
		dance="恭喜：["& name &"]邀请{"& aqjh_name &"}语聊成功，大家来看他们语聊！！<br><a href=liao.asp target=_blank><img src=img/dt5.gif width=149 height=85></a><br>"& name &"魅力上升500道德减100"& aqjh_name &"道德加500魅力减100"
	else
		randomize()
		r = Int(7 * Rnd)+1
	Select Case r
		Case 1
			dance=aqjh_name&"说：支持："&aqjh_name&"我支持["&momo&"与"&to1&"]语聊，他们语聊很风趣呵………"
		case 2 
			dance=aqjh_name&"说：支持："&aqjh_name&"我赞成["&momo&"与"&to1&"]语聊,最好还可以跳街舞哈."
		case 3
			dance=aqjh_name&"说：支持："&aqjh_name&"我赞成["&momo&"与"&to1&"]语聊,最好还可以跳肚皮舞哈."
		case 4
			dance=aqjh_name&"说：支持："&aqjh_name&"我赞成["&momo&"与"&to1&"]语聊,最好还可以一边跳一边脱这样才过隐哈."
		case 5
			dance=aqjh_name&"说：支持："&aqjh_name&"我赞成["&momo&"与"&to1&"]语聊,最好还可以跳街舞哈."
		case else
			dance=aqjh_name&"说：支持："&aqjh_name&"我赞成["&momo&"与"&to1&"]语聊,最好还可以跳街舞哈."
	end select
	if momo=aqjh_name then
			dance=aqjh_name&"说：支持："&aqjh_name&"我支持["&momo&"与"&to1&"]语聊，真是不要脸，哪里有自己支持自己语聊的！"
	end if
	end if
else
	if aqjh_name=to1 then
			if Application("aqjh_online")<>aqjh_name then
				Response.Write "<script language=JavaScript>{alert('提示：网络超时……');}</script>"
				Response.End 
			end if
		Response.Write "<script language=JavaScript>{alert('我才不干呢语聊想累死我阿！');}</script>"
		conn.execute "update 用户 set 魅力=魅力-100 where 姓名='"&name&"'"
		randomize()
		r = Int(6 * Rnd)+1
		Select Case r
			Case 1
				dance="【拒绝语聊】"&aqjh_name&"对"&momo&"说：你可以去找别人吗语聊很累也!["& momo &"]邀请{"& aqjh_name &"}语聊不成，魅力下降100点!"
			case 2
				dance="【拒绝语聊】"&aqjh_name&"对"&momo&"说：你可以去找别人吗语聊不是我的专长也.上床才是我的强项!["& momo &"]邀请{"& aqjh_name &"}语聊不成，魅力下降100点!"
			case 3
				dance="【拒绝语聊】"&aqjh_name&"对"&momo&"说：你可以去找别人吗语聊我不在行.邀我睡觉我可很愿意!["& momo &"]邀请{"& aqjh_name &"}语聊不成，魅力下降100点!"
			case 4
				dance="【拒绝语聊】"&aqjh_name&"对"&momo&"说：你可以去找别人吗语聊很累也!["& momo &"]邀请{"& aqjh_name &"}语聊不成，魅力下降100点!"
			case else
				dance="【拒绝语聊】"&aqjh_name&"对"&momo&"说：你可以去找别人吗语聊很累也!["& momo &"]邀请{"& aqjh_name &"}语聊不成，魅力下降100点!"
			end select
				
	else
		randomize()
		r = Int(6 * Rnd)+1
	Select Case r
		Case 1
			dance=aqjh_name&"说：反对："&aqjh_name&"我反对["&momo&"与"&to1&"]语聊，他们语聊根本就不好看丑死了！"
		case 2
			dance=aqjh_name&"说：反对："&aqjh_name&"我反对["&momo&"与"&to1&"]语聊，不如跟我跳.我跳街舞最厉害了呵~~~"
		case 3
			dance=aqjh_name&"说：反对："&aqjh_name&"我反对["&momo&"与"&to1&"]语聊，不如跟我跳.我跳肚皮舞最厉害了呵~~~"
		case 4
			dance=aqjh_name&"说：反对："&aqjh_name&"我反对["&momo&"与"&to1&"]语聊，不如跟我跳.我跳街舞最厉害了呵~~~"
		case 5
			dance=aqjh_name&"说：反对："&aqjh_name&"我反对["&momo&"与"&to1&"]语聊，不如跟我跳.我跳街舞最厉害了呵~~~"
		case else
			dance=aqjh_name&"说：反对："&aqjh_name&"我反对["&momo&"与"&to1&"]语聊，不如跟我跳.我跳街舞最厉害了呵~~~"
	end select
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【邀请语聊消息】</b></font>"&dance	
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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