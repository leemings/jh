<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE="JavaScript">if(window.name!="d"){var i=1;while(i<=50){window.alert("刷钱是吗？喜欢是吗？点啊，刷啊！！");
i=i+1;}top.location.href="../exit.asp"}</script>
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
yn=LCase(trim(request.querystring("yn")))
ns=application("aqjh_sjqr")
if ns="" then
	Response.Write "<script language=JavaScript>{alert('没有人提出过情人请求。');}</script>"
	Response.End 
end if
qrdata=split(ns,"|")
name=qrdata(0)
to1=qrdata(1)
tcsj=qrdata(2)
erase qrdata
if yn<>0 and yn<>1 then
	Response.Write "<script language=JavaScript>{alert('滚!刷不成金卡了。');}</script>"
	Response.End 
end if
nowsj=DateDiff("s",tcsj,now())
if nowsj>=30 then
	Response.Write "<script language=JavaScript>{alert('错误：网络超时，此情人请求超过30秒，已失效。');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 姓名,会员金卡 FROM 用户 WHERE 姓名='"&name&"'",conn,2,2
if rs("会员金卡")<2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方会员金卡不足2元!');}</script>"
	Response.End 
end if
if yn=1 then
	if aqjh_name=to1 then
		Application.Lock
		Application("aqjh_sjqr")=""
		Application.unLock
		Response.Write "<script language=JavaScript>{alert('我答应你的要求了！');}</script>"
		conn.execute "update 用户 set 会员金卡=会员金卡-2,情人='"& aqjh_name &"',结婚次数=结婚次数+1 where 姓名='" & name &"'"
		conn.execute "update 用户 set 会员金卡=会员金卡+1,情人='"& name &"',结婚次数=结婚次数+1 where 姓名='" & aqjh_name & "'"
		hunyin="恭喜：["& aqjh_name &"]与{"& name &"}喜结良缘，大家表示祝贺！！<img src='img/004.gif'><br><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>喜喜</font></marquee>"
		Application("aqjh_sjqr")=""
	else
		randomize()
		r = Int(7 * Rnd)+1
	Select Case r
		Case 1
			hunyin=aqjh_name&"说：支持："&aqjh_name&"我支持["&name&"与"&to1&"]作情人，一定行的………"
		case 2 
			hunyin=aqjh_name&"说：支持："&aqjh_name&"我支持["&name&"与"&to1&"]作情人,你是疯子他是傻,二半吊子少半啦,没有你就不疯来,少了他也就不在傻,你在哪来他在哪,你要死了就少俩."
		case 3
			hunyin=aqjh_name&"说：支持："&aqjh_name&"我支持["&name&"与"&to1&"]作情人,秋风清，秋月明。落叶聚还散，寒鸦栖复惊。相思相见知何日，此时此夜难为情。这种男人已经难找了，你就嫁给他吧!"
		case 4
			hunyin=aqjh_name&"说：支持："&aqjh_name&"我支持["&name&"与"&to1&"]作情人,我太佩服你了，而且敢这么说的一定是勇敢的人。我平时太腼腆，一见到异性就脸红，真没办法。"
		case 5
			hunyin=aqjh_name&"说：支持："&aqjh_name&"我支持["&name&"与"&to1&"]作情人,等到一切都看透，希望她能陪你看细水常流！如果她不答应你,那可能是她想和我看。"
		case else
			hunyin=aqjh_name&"说：支持："&aqjh_name&"我支持["&name&"与"&to1&"]作情人，天荒地老，希望你们情不变…"
	end select
	if name=aqjh_name then
			hunyin=aqjh_name&"说：支持："&aqjh_name&"我支持["&name&"与"&to1&"]作情人，真是不要脸，哪里有自己说自己好的！"
	end if
	end if
else
	if aqjh_name=to1 then
		Application.Lock
		Application("aqjh_sjqr")=""
		Application.unLock
		Response.Write "<script language=JavaScript>{alert('我才不干呢！');}</script>"
		conn.execute "update 用户 set 魅力=魅力-100 where 姓名='&name&'"
		randomize()
		r = Int(6 * Rnd)+1
		Select Case r
			Case 1
				hunyin="【拒婚】"&aqjh_name&"对"&name&"说：你可以找到比我更好的人……我不想再欺骗你和自己，希望你原谅我吧！我会永远祝福你的!["& name &"]向{"& aqjh_name &"}求婚不成，魅力下将100点!"
			case 2
				hunyin="【拒婚】"&aqjh_name&"对"&name&"说：我家里的狗不太喜欢你，你家里的猫不太喜欢我，所以不要再说什么，以后你慢慢就明白的!"
			case 3
				hunyin="【拒婚】"&aqjh_name&"对"&name&"说：不知何时，快乐悄悄褪色，失去已无法挽回.....做朋友好吗？"
			case 4
				hunyin="【拒婚】"&aqjh_name&"对"&name&"说：曾经的沧海已经退潮，我的心里已经不再有激情，请你还是找别的女孩吧。"
			case else
				hunyin="【拒婚】"&aqjh_name&"对"&name&"说：你看看我的样，你也配，我就是嫁给猪无能，也不会嫁给你这个鳖三…"
			end select
				
	else
		randomize()
		r = Int(6 * Rnd)+1
	Select Case r
		Case 1
			hunyin=aqjh_name&"说：反对："&aqjh_name&"我反对["&name&"与"&to1&"]作情人，他们根本不合适………还不如我！"
		case 2
			hunyin=aqjh_name&"说：反对："&aqjh_name&"我反对["&name&"与"&to1&"]作情人，别嫁给他，你看我比他英俊多了，我有车子，房子，票子，现在就缺妻子了。然后ZZZ撇撇嘴想：呵呵~~~先骗骗你再说~~~"
		case 3
			hunyin=aqjh_name&"说：反对："&aqjh_name&"我反对["&name&"与"&to1&"]作情人，他是色狼！前几天我还看到他从丽春院出来……不行！！"
		case 4
			hunyin=aqjh_name&"说：反对："&aqjh_name&"我反对["&name&"与"&to1&"]作情人，他是色狼！前几天我还看到他从丽春院出来……不行！！"
		case 5
			hunyin=aqjh_name&"说：反对："&aqjh_name&"我反对["&name&"与"&to1&"]作情人，他虚情假意 恐怕只会拍马屁 !"
		case else
			hunyin=aqjh_name&"说：反对："&aqjh_name&"我反对["&name&"与"&to1&"]作情人，嫁给他前认为美丽的事，嫁给他后便完全变为丑恶；山盟海誓的真情，也变成了虚情假意的欺骗，所以请你千万不要嫁给他!"
	end select
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>【江湖消息】</b></font>"&hunyin	
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
