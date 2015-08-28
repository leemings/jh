<%Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
st=LCase(trim(request.querystring("un")))
if InStr(st,"or")<>0 or InStr(st,"'")<>0 or InStr(st,"`")<>0 or InStr(st,"=")<>0 or InStr(st,"-")<>0 or InStr(st,",")<>0 or InStr(st,"&")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from s where ss='"&username&"' and gyr='"&st&"' and wc='3'",conn
if rst.eof or rst.bof then
%>
<script language=vbscript>
MsgBox "人家没有请你啊!"
window.close()
</script>
<%
rst.close
set rst=nothing
conn.close
set conn=nothing
Response.End
else
bc=rst("bc")
bsr=rst("bsr")
abc=1
qssok=""&username&"接受了"& st &"的雇用<img src=image/IMAGE/ws.gif>，佣金"&bc&"两,成功追杀后,"&username&"将得到这笔钱!"
sql="update s set wc='0' where ss='"&username&"' and wc='3' and gyr='"&st&"'"
conn.execute(sql)
sql="delete * from s where ss='"&username&"' and wc='3'"
conn.execute(sql)
rst.close
set rst=nothing
conn.close
set conn=nothing
if abc=1 then
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
	newtalkarr(595)=st
	newtalkarr(596)="" 
	newtalkarr(597)="000000" 
	newtalkarr(598)="000000" 
	newtalkarr(599)="<b><font color=FF0000>【接受雇用】</font></b>"&qssok&"<font class=timsty>("&time()&")</font>"
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
end if
%>
<script language=vbscript>
MsgBox "你接受了雇用，你要杀的目标是[<%=bsr%>],佣金<%=bc%>,成功追杀后你将得到这笔钱"
window.close()
</script>
<%
end if
%>