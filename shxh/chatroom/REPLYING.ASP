<%
username=session("Ba_jxqy_username")
mycorp=session("Ba_jxqy_usercorp")
chatroomsn=session("Ba_jxqy_userchatroomsn")
if username=Application("Ba_jxqy_askuser") then response.redirect "../error.asp?id=061"
if username="" then response.redirect "../error.asp?id=016"
if mycorp="官府" then response.redirect "../error.asp?id=062"
reply=trim(Request.form("reply"))
reply=server.HTMLEncode(reply)
if trim(Request.form("reply"))="" then Response.Redirect "../error.asp?id=060"
reply=replace(reply,"\","\\")
reply=replace(reply,"/","\/")
reply=replace(reply,"'","\'")
reply=replace(reply,chr(34),"\"&chr(34))
if reply=Application("Ba_jxqy_reply") then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("Ba_jxqy_connstr")
conn.open connstr
 sql="SELECT 银两 FROM 用户 WHERE 姓名='" & username & "'"
 rs.open sql,conn,1,3
 rs("银两")=rs("银两")+Application("Ba_jxqy_asksilver")
 rs.Update
 rs.close
 conn.close
	talkarr=Application("Ba_jxqy_talkarr")
	talkpoint=int(Application("Ba_jxqy_talkpoint"))
	Application("Ba_jxqy_talkpoint")=talkpoint+1
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
	newtalkarr(599)="<b>【答 题】<font color=balck>" & Application("Ba_jxqy_ask") & "？</font>正确答案是：<font color=red>[ " & Application("Ba_jxqy_reply") & " ] </font>〖"& username &"〗因此获得了"&Application("Ba_jxqy_asksilver")&"两银子的奖励!</b><font class=\'timsty\'>("&time()&")</font>"
    newtalkarr(600)=chatroomsn
	Application("Ba_jxqy_talkarr")=newtalkarr
	Application("Ba_jxqy_ask")=""
	Application("Ba_jxqy_reply")=""
	Application("Ba_jxqy_askuser")=""
else
response.redirect "../error.asp?id=059"
response.end
end if
%>
<script LANGUAGE="JavaScript">
if(window!=window.top){top.location.href=location.href;}
window.close();
</script>

