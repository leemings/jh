<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
chatroomsn=session("yx8_mhjh_userchatroomsn")
if username="" then response.redirect "../error.asp?id=016"
if session("yx8_mhjh_usercorp")="官府"  then
ask=trim(Request.form("ask"))
reply=trim(Request.form("reply"))
silver=Trim(Request.Form("silver"))
ask=server.HTMLEncode(ask)
reply=server.HTMLEncode(reply)
if not isnumeric(silver) then Response.Redirect "../error.asp?id=024"
silver=clng(silver)
if silver<1 then
silver=1
elseif silver>500000 then
silver=500000
end if
if trim(Request.form("ask"))="" then Response.Redirect "../error.asp?id=060"
if trim(Request.form("reply"))="" then Response.Redirect "../error.asp?id=060"
ask=replace(ask,"\","\\")
ask=replace(ask,"/","\/")
ask=replace(ask,"'","\'")
ask=replace(ask,chr(34),"\"&chr(34))
reply=replace(reply,"\","\\")
reply=replace(reply,"/","\/")
reply=replace(reply,"'","\'")
reply=replace(reply,chr(34),"\"&chr(34))
Application("yx8_mhjh_ask")=ask
Application("yx8_mhjh_reply")=reply
Application("yx8_mhjh_askuser")=username
Application("yx8_mhjh_asksilver")=silver
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=int(Application("yx8_mhjh_talkpoint"))
Application("yx8_mhjh_talkpoint")=talkpoint+1
Dim newtalkarr(600)
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
newtalkarr(j+8)=talkarr(i+9)
j=j+10
next
newtalkarr(591)=talkpoint+1
newtalkarr(592)=1
newtalkarr(593)=0
newtalkarr(594)=username
newtalkarr(595)="大家"
newtalkarr(596)=""
newtalkarr(597)="660099"
newtalkarr(598)="660099"
newtalkarr(599)="<b>【出 题】<font color=balck>" & Application("yx8_mhjh_ask") & "？ </font>正确答案是什么？[提问人]:<font color=red>〖" & Application("yx8_mhjh_askuser")&"〗[奖金]"&silver&"</font></b><font class=timsty>("&now()&")</font>"
newtalkarr(600)=chatroomsn
Application("yx8_mhjh_talkarr")=newtalkarr
%>
<script LANGUAGE="JavaScript">
if(window!=window.top){top.location.href=location.href;}
window.close();
</script>
<%
else
response.redirect "../error.asp?id=046"
end if
%>
