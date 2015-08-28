<%
Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
username=session("yx8_mhjh_username")
usercorp=session("yx8_mhjh_usercorp")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if username="" then Response.Redirect "../error.asp?id=016"
if username<>Application("yx8_mhjh_admin") or usercorp<>"官府" then
	Response.Write "<script language=JavaScript>{alert('提示：这一些只有站长才可以操作!');parent.history.go(-1);}</script>"
	Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chatroom")=0 then 
Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n 按确定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
        response.write "禁止站外提交！"
        response.end
end if
cz=LCase(trim(Request.QueryString("cz")))
value=int(abs(clng(Request.QueryString("value"))))
randomize timer
s=int(rnd*value)+1
yn=0
select case cz
case "银两"
	yn=1
	fafang="<bgsound src=../mid/Faqian.wav loop=1><font color=green>【发钱】</font><font color=blue>"& username &"</font><B><font color=#ff0000>考虑到灾民过多,决定开仓放粮,向聊天室的每个人发了<font color=blue>+"& s &"</font>两<img src=../image/251.gif>解决灾情！</font></b>"
case "体力"
	yn=1
	fafang="<bgsound src=../mid/Faqian.wav loop=1><font color=green>【发体力】</font>大侠<font color=blue>"& username &"</font><B><font color=#ff0000>最近发福了，路过福缘客栈，请所有大侠吃饭,大侠们体力增加：<font color=blue>+"& s &"</font>点!！</font></b>"
case "积分"
	yn=1
	fafang="<bgsound src=../mid/Faqian.wav loop=1><font color=green>【发积分】</font><font color=blue>"& username &"</font><B><font color=#ff0000>泡点是很难的，站长为了鼓励各位玩家泡点，每个人发放积分<font color=blue>+"& s &"</font>点，大家还要继续努力~~</font></b>"
case "内力"
	yn=1
	fafang="<bgsound src=../mid/Faqian.wav loop=1><font color=green>【发内力】</font>大侠<font color=blue>"& username &"</font><B><font color=#ff0000>路过江湖演武场，不忍看到江湖兄弟们辛苦修练，便传授一招任我行的吸星大法，谁知却被各位大虾每人吸走内力:<font color=blue>+"& s &"</font>，唉！好心没有好报~~</font></b>"
case "衰哥"
	yn=0
	Application.Lock
	Application("jxqy_sg")=s+100
	Application.UnLock
	abc="<a href=boy.asp?tl="&Application("jxqy_sg")&" target=optfrm><img src=../image/p42.gif border=0></a>"
	fafang="<font color=green>【消息】</font><font color=red>一个衰哥听说这个江湖的美女特别多，跑到江湖想泡美女，大家准备好棒子打呀。打跑了官府有奖。。衰哥体力："&s+100&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</a></marquee>"
case "元宝"
	yn=0
	Application.Lock
	Application("jxqy_yb")=s+100
	Application.UnLock
	abc="<a href=yb.asp?tl="&Application("jxqy_yb")&" target=optfrm><img src=../image/251.GIF border=0></a>"
	fafang="<font color=green>【消息】</font><font color=red>今天是什么日子，突然间从天上掉下来一个大元宝，价值："&s+100&"两!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
case "精力"
	yn=1
	fafang="<bgsound src=../mid/Faqian.wav loop=1><font color=green>【精力】</font>大侠<font color=blue>"& username &"</font><B><font color=#ff0000>最近发福了，路过福缘客栈，请所有大侠吃饭,大侠们精力增加：<font color=blue>+"& s &"</font>点!！</font></b>"
case "攻击"
	yn=1
	fafang="<bgsound src=../mid/Faqian.wav loop=1><font color=green>【攻击】</font>大侠<font color=blue>"& username &"</font><B><font color=#ff0000>路过江湖演武场，不忍看到江湖兄弟们辛苦修练，便传授一招任我行的吸星大法，谁知却被各位大虾每人吸走攻击力:<font color=blue>+"& s &"</font>，唉！好心没有好报~~</font></b>"
case "防御"
	yn=1
	fafang="<bgsound src=../mid/Faqian.wav loop=1><font color=green>【防御】</font>大侠<font color=blue>"& username &"</font><B><font color=#ff0000>路经此地，给聊天室中每人发放防御力：<font color=blue>+"& s &"</font>点~~~</font></b>"
end select
if yn=1 then
    onlinelist=Application("yx8_mhjh_onlinelist")
    set conn=server.createobject("adodb.connection") 
    conn.open Application("yx8_mhjh_connstr")
    for i=1 to ubound(onlinelist) step 6
    conn.Execute "update 用户 set "&cz&"="&cz&"+" & s & " where 姓名='"&onlinelist(i)&"'"
    next
    erase onlinelist
	conn.close
	set conn=nothing
end if
dim newtalkarr(600) 
talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)=""&fafang&""
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
Response.Write "<script language=JavaScript>{alert('[操作完成]');parent.history.go(-1);}</script>"	
Response.End 
%>