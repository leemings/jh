<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then
Response.Redirect "../error.asp?id=016"
else
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 配偶,洞房花烛,银两,道德 from 用户 where  姓名='"&username&"'",conn
if rst.EOF or rst.BOF then Response.Redirect "error.asp?id=019"
mate=rst("配偶")
dong=rst("洞房花烛")
yl=rst("银两")
dd=rst("道德")
rst.Close
set rst=nothing
if yl>=1000000 or dd>=2000 then
if mate<>"无" then 
conn.Execute "update 用户 set 洞房花烛='无',配偶='无',银两=银两-1000000,道德=道德-2000 where 姓名='"&username&"'"
conn.Execute "update 用户 set 洞房花烛='无' where 姓名='"&dong&"'"
conn.Execute "update 用户 set 配偶='无' where 姓名='"&mate&"'"
else
Response.write "你还没有配偶，赶快找一个吧！"
Response.end
end if
else
Response.write "你的钱不够100万或者道德不够2000，不能逃离苦海了！"
Response.end
end if
conn.close
set conn=nothing
Response.Write "<head><title>"&Application("yx8_mhjh_systemname")&"</title><link rel='stylesheet' href='../style4.css'></head><body background='../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false><p align=center><br>逃婚成功!你脱离情海了！</p></body>"
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
		newtalkarr(599)="<font color=red>【江湖逃婚】</font><font color=blue>"&username&"对自己的配偶实在失望，最终决定花费100万银子逃离了婚姻的苦海！道德下降2000</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>