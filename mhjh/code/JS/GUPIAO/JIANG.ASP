<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if session("yx8_mhjh_inchat")<>"1" then %>
<script language="vbscript">
alert "请进入聊天室再进入股票！"
window.close()
</script>
<%Response.End
end if
if username=""  then
  Response.Redirect "../../error.asp?id=016"
elseif username<>Application("yx8_mhjh_admin")then
	Response.Redirect "../../error.asp?id=439"
else%>
<!--#include file="jhb.asp"-->
<%
for sj=1 to 20
sql= "select * from 股票 where sid="&sj        
set rs=conn.execute(sql) 
if (rs("当前价格")-rs("开盘价格"))/rs("开盘价格")>-0.01 then 
sql="update 股票 set 当前价格=当前价格*0.9 where sid="&sj
conn.execute sql 
end if
next 
mess="股票市场走向熊市所有股票低迷不振价格下挫10%"
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
		newtalkarr(599)="<font color=red>【股市】</font><b><font color=blank>"& mess &"</font></b>"	
        newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
        Application.Lock
        Application("yx8_mhjh_talkpoint")=talkpoint+1
        Application("yx8_mhjh_talkarr")=newtalkarr
        Application.UnLock
        erase newtalkarr
        erase talkarr
       conn.close
Response.Redirect "index.asp"
end if
%>

