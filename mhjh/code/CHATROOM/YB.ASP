<%
Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
username=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if username="" then Response.Redirect "../error.asp?id=016"
if Application("jxqy_yb")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你小子来晚了，元宝没有了...');location.href='option.asp';</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
temp=int(abs(clng(Application("jxqy_yb"))))
if temp<>tl then
	Response.Write "<script Language=Javascript>alert('提示：滚，你想搞什么，按回车键不行了!');location.href='option.asp';</script>"
	response.end
end if
Application.Lock
Application("jxqy_yb")=0
Application.UnLock
name=username
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("yx8_mhjh_connstr")
conn.execute "update 用户 set 银两=银两+"& temp &" where 姓名='" & username &"'"
yb="<font color=red><b>【江湖消息】</b></font>今天真是走运，出门都有钱砸到我的脑袋上，呵。。。因此：["&name&"]得到:<img src=../image/251.gif>"&temp&"两!"
conn.close
set conn=nothing
		
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
		newtalkarr(599)=""&yb&""
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
%>
<script Language=Javascript>location.href = "option.asp";</script>