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
if instr(aberrantname,username)<>0 then
Response.Write "<script language=javascript>alert('状态异常中，无法退出！如果强行退出后果自负！！');this.history.back();</script>"
Response.End
end if
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

Application("yx8_mhjh_allonlinenum")=clng(Application("yx8_mhjh_onlinenum1")*(rnd()+2))
session("yx8_mhjh_inchat")=""
if Instr(Application("yx8_mhjh_zlname"),";"&username&";")<>0 then
zllist=Application("yx8_mhjh_zlname")
Application.Lock
Application("yx8_mhjh_zlname")=replace(zllist,";"&username&";",";")
Application.UnLock
end if
Response.Write "<html><body><script language=javascript>parent.window.close();</script></body></html>"
%>

