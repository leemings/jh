<%@ LANGUAGE=VBScript%>
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
response.write "你提交的路径有误，禁止从站点外部提交"
response.end
end if
chatroomsn=Session("yx8_mhjh_userchatroomsn")
useronlinename=Application("yx8_mhjh_onlinename"&chatroomsn)
nickname=session("yx8_mhjh_username")
if nickname="" or session("yx8_mhjh_inchat")<>"in" then Response.Redirect "../../error.asp?id=016"
id=Request.Form ("id")
if id="" then Response.end
towho=Request.Form ("towho")
if towho="" or towho=nickname or towho="大家" or Instr(useronlinename,";"&towho&";")=0 or towho="魔导之神"  then 
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您好像送错了对象！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
Response.End 
end if
%><!--#include file="data2.asp"--><%  
rs.Open ("select * from pet where username='"&nickname&"' and rest=0 and id="&id),connb
if rs.BOF or rs.EOF then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您没有此种宠物，你是不是想作弊！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
rs.Close
set rs=nothing
connb.Close
set connb=nothing
Response.End 
else
animalname=rs("animalname")
connb.Execute ("update myanimal set username='"&towho&"' where username='"&nickname&"' and rest=0 and id="&id)
msg="<font color=FF0000>【消息】</font>["&nickname&"]将自己心爱的宠物["&animalname&"]送给了朋友["&towho&"]！但愿这份爱心能换取["&towho&"]纯洁的友谊！"
call chat(says)
rs.Close
set rs=nothing
connb.Close
set connb=nothing
sub chat(says)
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
		newtalkarr(599)=""&msg&""
        newtalkarr(600)=chatroomsn
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr	
end sub

Response.Write "<script language=JavaScript>{alert('宠物赠送成功！！');location.href='myanimal.asp';}</script>"
end if
%>

