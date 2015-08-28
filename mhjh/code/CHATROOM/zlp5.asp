<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
st=Session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
if Application("yx8_mhjh_klt5")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你动作或者反应太慢，这种战利品已经被人拿走了！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("yx8_mhjh_klt5"))))
'if tempjs<>jst3 then
	'Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');</script>"
	'Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	'response.end
'end if
conn.execute "update 用户 set 体力=体力-"&Application("yx8_mhjh_klt5")&" where 姓名='"&st&"'"
rs.open "select 体力,门派 FROM 用户 WHERE 姓名='"&st&"'",conn
if rs("门派")="十八地狱" then
	Response.Write "<script Language=Javascript>alert('提示：你小子是不是没事找事呀，鬼魂还想得便宜啊！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>"
	Application("yx8_mhjh_klt5")=0
	Response.end
end if
if rs("体力")<0 then
	conn.execute "update 用户 set 状态='死亡',被杀=被杀+1 where 姓名='"&st&"'"
conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&username&"','吝啬','体力不足累死')"
	kl="<img src='data/kl.gif'>太不幸了，<font color=0000FF>"&st&"</font>抢夺战利品时遇到强敌,由于体力不足而死,有够惨的!"
	Response.Write "<script Language=Javascript>alert('提示：你贪财不要命，体力太弱，丢了小命，去复活吧！没能耐别乱闯！');</script>"
	session.Abandon
	response.write "<script language=javascript>parent.top.location.href='chaterror.asp?id=009';</script>"
else
	conn.execute "update 用户 set 风铃=风铃+1 where 姓名='" & st & "'"
	kl="<font color=0000FF>"&st&"</font>动作迅速，成功抢得武术家掉下的一个风铃！"
	Application.Lock
	Application("yx8_mhjh_klt5")=0
	Application.UnLock
	rs.close
set rs=nothing
conn.close
set conn=nothing
end if
says="<font color=red><b>【抢战利品】</b></font>"&kl	
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
		newtalkarr(598)="008000" 
		newtalkarr(599)=""&says&""
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script Language=Javascript>alert('提示：抢夺战利品成功,注意看聊天室内的江湖通报！');</script>"
	Response.Write "<script Language=Javascript>location.href ='javascript:window.close()';</script>
