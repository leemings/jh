<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
username=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if username="" then Response.Redirect "../error.asp?id=440"
if Application("jxqy_sg")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你动作或者反应太慢，机会已经错过了！');</script>"
	Response.Write "<script Language=Javascript>location.href = 'option.asp';</script>"
	Response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("jxqy_sg"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	Response.Write "<script Language=Javascript>location.href = 'option.asp';</script>"
	Response.end
end if
Application.Lock
Application("jxqy_sg")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("yx8_mhjh_connstr")
conn.execute "update 用户 set 内力=内力-"&tempjs&" where 姓名='" & username &"'"
rs.open "select 内力,头像 FROM 用户 WHERE 姓名='" & username &"'",conn
if rs("内力")<-100 then
    conn.execute "update 用户 set 状态='死亡',被杀=被杀+1 where 姓名='"&username&"'"
conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&username&"','吝啬','内力不足累死')"
	kl="哈哈，<img border=0 src="&rs("头像")&">["&username&"]贪财不要命，在前往神秘岛屿途中，连"&tempjs*2&"的钱都不肯花，弄得精疲力尽，内力空空，活活的给累死了………尸体在荒野上花作鬼魂，["&username&"]这个名字却永远留在后人纪念的[英烈堂]！"
	session.Abandon
else
conn.execute "update 用户 set 银两=银两+"& tempjs*30 &",精力=精力+10000 where 姓名='" & username & "'"
kl="青山依旧在，几度夕阳红。<img border=0 src="&rs("头像")&">["&username&"]梦中神游蓬莱仙境，巧遇传说中的魔幻之神派送生命之果，由于["&username&"]功力深厚，小胜魔幻之神一招，赢得了一颗红色的果实，无奈，在内力不足时被人抢走，那人丢给他白银<img src=../image/251.gif>"&tempjs*30&"两！在疗伤时，意外提高精力10000点！"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing

says="<font color=red><b>【梦想成真】</b></font>"&kl	
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
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script Language=Javascript>location.href = "option.asp";</script>
