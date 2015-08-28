<%
username=session("yx8_mhjh_username")
if username="" then
Response.Write "<script language=javascript>"
Response.Write "parent.game1.location.href='/p';"
Response.Write "</script>"
Response.End
end if
Set cn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
datestr=Application("yx8_mhjh_connstr")
cn.open datestr
cn.Execute ("update 用户 set 精力=精力+500 where 姓名='"&username&"'")
	kl="<font color=0000FF>"&username&"</font>在娱乐中心玩五子棋,凭着过人的智慧,赢了500精力！"
cn.Close
set cn=nothing
says="<font color=red><b>【娱乐赌博】</b></font>"&kl	
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
		newtalkarr(594)=name
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
cn.Execute ("update game set win=win+1 where username='"&username&"'")
cn.Close
set cn=nothing
Response.Write "<script language=javascript>"
Response.Write "parent.fb.location.href='button.asp';"
Response.Write "parent.t.location.href='start.asp';"
Response.Write "</script>"
Response.End
%>