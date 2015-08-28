<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
username=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if username="" then Response.Redirect "../error.asp?id=440"
if Application("jxqy_sg")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子来晚了，衰哥被别人打跑了！');</script>"
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
conn.execute "update 用户 set 体力=体力-"&tempjs&" where 姓名='" & username &"'"
rs.open "select 体力,头像 FROM 用户 WHERE 姓名='" & username &"'",conn
if rs("体力")<-100 then
    conn.execute "update 用户 set 状态='死亡',被杀=被杀+1 where 姓名='"&username&"'"
	kl="哈哈，<img border=0 src="&rs("头像")&">["&username&"]拿着棒子来打衰哥，谁知道体力不如人家，["&username&"]挨了衰哥一招晴空霹雳，死翘翘了………"
	session.Abandon
else
conn.execute "update 用户 set 银两=银两+"& tempjs*3 &" where 姓名='" & username & "'"
kl="哈哈，<img border=0 src="&rs("头像")&">["&username&"]拿着棒子打跑了来江湖泡美女的衰哥，得到了官府奖励<img src=../image/251.gif>"&tempjs*3&"两！在打败衰哥后，"&username&"得到了[<b><font color=red>衰哥的一条毛</font></b>]……"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing

says="<font color=red><b>【江湖消息】</b></font>"&kl	
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
