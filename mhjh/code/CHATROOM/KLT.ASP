<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
st=Session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_klt")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不ssss是没事找事呀，哪里有鬼？！');</script>"
	Response.Write "<script Language=Javascript>location.href = 'option.asp';</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("yx8_mhjh_klt"))))
'if tempjs<>tl then
	'Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	'Response.Write "<script Language=Javascript>location.href = 'option.asp';</script>"
	'response.end
'end if
conn.execute "update 用户 set 体力=体力-"&Application("yx8_mhjh_klt")&" where 姓名='"&st&"'"
rs.open "select 体力,门派 FROM 用户 WHERE 姓名='"&st&"'",conn
if rs("门派")="官府" then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不ssss是没事找事呀，官府的阳气这么重，鬼早就闪了！');</script>"
	Response.Write "<script Language=Javascript>location.href = 'option.asp';</script>"
	Application("yx8_mhjh_klt")=0
	Response.end
end if
if rs("体力")<Application("yx8_mhjh_klt") then
	conn.execute "update 用户 set 状态='死亡',被杀=被杀+1 where 姓名='"&st&"'"
	kl="<img src='data/kl.gif'>太不幸了，["&st&"]为了保护大家的安全，舍身取义，被僵尸咬的七零八落，死翘翘了………"
	Response.Write "<script Language=Javascript>alert('提示：你英勇可嘉，但是体力太弱，给僵尸活生生咬死了！');</script>"
	session.Abandon
	response.write "<script language=javascript>parent.top.location.href='chaterror.asp?id=009';</script>"
else
	conn.execute "update 用户 set 银两=银两+"& Application("yx8_mhjh_klt") &" where 姓名='" & st & "'"
	kl="英雄，真是英雄，["&st&"]与僵尸大打出手……僵尸终于倒下了，而他自己也伤的不轻，官府奖励<img src=data/251.gif>"&Application("yx8_mhjh_klt")&"两"
	Application.Lock
	Application("yx8_mhjh_klt")=0
	Application.UnLock
	rs.close
set rs=nothing
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
		newtalkarr(599)="<font color=red>【僵尸】"&kl
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script Language=Javascript>location.href = "option.asp";</script>
