<%@ LANGUAGE=VBScript%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
un=session("yx8_mhjh_username")
usercorp=session("yx8_mhjh_usercorp")
usergrade=session("yx8_mhjh_usergrade")
if usercorp<>"官府" then
Response.Write "<script language=JavaScript>{alert('  对不起！\n 非官府禁止进入！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
Response.End 
end if
if usergrade<40 then Response.End 
id=request("id")
if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"　")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 or InStr(id,"&")<>0 then Response.Redirect "../../error.asp?id=524"
if un="" then Response.Redirect "../../error.asp?id=016" %> 
<!--#include file="data.asp"--><%   
sql= "select username,animalname from myanimal where outtime=true and rest=0 and id="&id&""    
rs.open sql,connb,1,2    
if rs.EOF or rs.BOF then     
Response.Write "<script language=JavaScript>{alert('  对不起！\n  没有此神龙！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
rs.Close
set rs=nothing
connb.Close
set connb=nothing
Response.End 
end if
username=rs("username")
animalname=rs("animalname")
connb.Execute ("update myanimal set outtime=False where outtime=true and  rest=0 and id="&id)
msg="<font color=FF0000>【消息】</font>官府大人"&un&"将"&username&"的神龙["&animalname&"]无罪释放了！"
call chat(says)
rs.Close    
set rs=nothing    
connb.close    
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
   Response.Write "<script language=JavaScript>{alert(' 释放成功！！！\n  按 [确定] 返 回！');location.href='l.asp';}</script>"

%>     

