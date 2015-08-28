<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
response.write "<br><br><center><table border=1 cellpadding=20 bordercolor=black width=450>"
response.write "<tr><td style='font:9pt Verdana'>"
response.write "你提交的路径有误，禁止从站点外部提交数据"
response.write "</td></tr></table></center>"
response.end
end if
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
chatroomsn=session("yx8_mhjh_userchatroomsn")
animalname=Request.Form ("animalname")
newname=Request.Form ("newname")
if newname="" or animalname="" then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您没有此种神龙，或您给您的神龙取名为空！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
Response.End 
end if
function chuser(u)
dim filter,xx,newnameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
if chuser(newname) then Response.Redirect "../../error.asp?id=524"
for each element in Request.Form
elevalue=Request.Form(element)
if instr(elevalue,"'")<>0 or instr(elevalue,chr(34))<>0 or instr(elevalue,"\")<>0 or instr(elevalue,";") or instr(elevalue,"or") or instr(elevalue,"&")  then Response.Redirect "../error.asp?id=056"
next
for i=1 to len(newname)
newnamechr=mid(newname,i,1)
if asc(newnamechr)>0 then Response.Redirect "../../error.asp?id=003"
next
namelen=0
for i=1 to len(newname)
zh=mid(newname,i,1)
zhasc=asc(zh)
if zhasc<0 then
namelen=namelen+2
else
namelen=namelen+1
if CStr(server.URLEncode(zh))<>CStr(zh) then Response.Redirect "../../error.asp?id=510"
end if
next
if namelen>8 then 
Response.Write "<script language=JavaScript>{alert('  对不起！\n  名字太长，最多4个汉字，请返回更正！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
Response.end
end if

namelen=0
if InStr(LCase(newname),"fuck")<>0 or InStr(LCase(newname),"sex")<>0 or InStr(newname,"奸")<>0 or InStr(newname,"淫")<>0 or InStr(newname,"娼")<>0 or InStr(newname,"嫖")<>0 or InStr(newname,"性")<>0 and InStr(newname,"交")<>0 or InStr(newname,"妓")<>0 or InStr(newname,"色")<>0 and InStr(newname,"黄")<>0 or InStr(newname,"色")<>0 and InStr(newname,"情")<>0 or InStr(newname,"日")<>0 and InStr(newname,"妈")<>0 or InStr(newname,"日")<>0 and InStr(newname,"妹")<>0 or InStr(newname,"日")<>0 and InStr(newname,"姐")<>0 or InStr(newname,"日")<>0 and InStr(newname,"娘")<>0 or InStr(newname,"日")<>0 and InStr(newname,"奶")<>0 or InStr(newname,"乳")<>0 or InStr(newname,"阴")<>0 or InStr(newname,"操")<>0 then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  名字含有不雅字眼，请返回更正！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
Response.End 
end if

%><!--#include file="data.asp"--><% 
rs.Open ("select * from myanimal where animalname='"&animalname&"' and username='"&username&"' and rest=0"),connb
if rs.BOF or rs.EOF then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您没有此种神龙！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
rs.Close
set rs=nothing
connb.Close
set connb=nothing
Response.End 
else
id=rs("id")
rs.Close
rs.Open ("select * from myanimal where animalname='"&newname&"' and rest=0 and username='"&username&"'"),connb
if not(rs.BOF or rs.EOF) then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您己有此种名字的神龙！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
rs.Close
set rs=nothing
connb.Close
set connb=nothing
Response.End 
end if
connb.Execute ("update myanimal set animalname='"&newname&"' where username='"&username&"' and rest=0 and id="&id)
msg="<font color=FF0000>【消息】</font>["&username&"]将自己的神龙["&animalname&"]取了一个好听的名字["&newname&"]！"
call chat(says)
end if
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

Response.Write "<script language=JavaScript>{alert('  恭喜您！\n  您己将神龙"&animalname&"改名为"&newname&"！！！\n  按 [确定] 返 回！');location.href='myanimal.asp';}</script>"
%>
<html></html>
