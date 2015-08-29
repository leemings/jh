<!--#include file="../config.asp"-->
<%
chatroomsn=session("yx8_mhjh_userchatroomsn")
username=session("yx8_mhjh_username")
chatroomname=Application("yx8_mhjh_systemname"&chatroomsn)
onlinelist=Application("yx8_mhjh_onlinelist")
aberrantname=Application("yx8_mhjh_aberrantname")
chatroomnum=Application("yx8_mhjh_chatroomnum")
maxnosaytime=Application("yx8_mhjh_nosaytime")
paycent=Application("yx8_mhjh_paycent")
lastsavetime=session("yx8_mhjh_userlastsavetime")
if lastsavetime="" or isnull(lastsavetime) then lastsavetime=now()
'if instr(aberrantname,username)<>0 then
'Response.Write "<script language=javascript>alert('状态异常中，无法退出！如果强行退出后果自负！！');this.history.back();</script>"
'Response.End
'end if
onlinelistubd=ubound(onlinelist)
nowtime=now()
expvalue=datediff("n",lastsavetime,nowtime)
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
sqlstr="update 用户 set 攻击=攻击+"&expvalue&",防御=防御+"&expvalue&",体力=体力+"&expvalue&",内力=内力+"&expvalue&",银两=银两+"&expvalue*paycent&",精力=精力+"&expvalue&",积分=积分+"&expvalue&",保护=False where 姓名='"&username&"'"
conn.Execute sqlstr
conn.Close
set conn=nothing
session("yx8_mhjh_userlastsavetime")=nowtime
msg=replace(yx8_mhjh_guestleft,"@#",Application("yx8_mhjh_systemname"&chatroomsn))
j=1
dim newonlinelist()
dim newonlinename()
newonlinenum=0
for i=1 to chatroomnum
redim preserve newonlinename(i)
newonlinename(i)=";"
next
for i=1 to onlinelistubd step 6
if onlinelist(i)<>username and datediff("s",onlinelist(i+4),nowtime)<=maxnosaytime then
newonlinenum=newonlinenum+1
newonlinename(onlinelist(i+5))=newonlinename(onlinelist(i+5))&onlinelist(i)&";"
redim preserve newonlinelist(j),newonlinelist(j+1),newonlinelist(j+2),newonlinelist(j+3),newonlinelist(j+4),newonlinelist(j+5)
newonlinelist(j)=onlinelist(i)
newonlinelist(j+1)=onlinelist(i+1)
newonlinelist(j+2)=onlinelist(i+2)
newonlinelist(j+3)=onlinelist(i+3)
newonlinelist(j+4)=onlinelist(i+4)
newonlinelist(j+5)=onlinelist(i+5)
j=j+6
elseif 	datediff("s",onlinelist(i+4),nowtime)>maxnosaytime and onlinelist(i)<>username then
Application.Lock
Application("yx8_mhjh_allonline")=replace(Application("yx8_mhjh_allonline"),";"&onlinelist(i)&";",";")
Application.UnLock
end if
next
Application.Lock
if newonlinenum=0 then
dim index(0)
Application("yx8_mhjh_onlinelist")=index
else
Application("yx8_mhjh_onlinelist")=newonlinelist
end if
Application("yx8_mhjh_allonlinenum")=newonlinenum
for i=1 to chatroomnum
Application("yx8_mhjh_onlinename"&i)=newonlinename(i)
next
Application.UnLock
erase newonlinelist
erase newonlinename
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=Application("yx8_mhjh_talkpoint")
dim newtalkarr(600)
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
if username=Application("yx8_mhjh_admin") then 
    if Application("yx8_mhjh_yinshen")=True then
      newtalkarr(599)=""
    else
      newtalkarr(599)="<font color=FF0000>【公告】</font>"&msg&"<font class=timsty>"&nowtime&"</font>"
    end if
else
    newtalkarr(599)="<font color=FF0000>【公告】</font>"&msg&"<font class=timsty>"&nowtime&"</font>"
end if
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
Application("yx8_mhjh_allonlinenum")=clng(Application("yx8_mhjh_onlinenum1")*(rnd()+2))
session("yx8_mhjh_inchat")=""
Response.Write "<html><body><script language=javascript>parent.window.close();</script></body></html>"
%>

