<%Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=106"
zj=LCase(trim(request.querystring("name")))
if zj<>username then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？人家向你求婚吗？！');}</script>"
	Response.End
end if
qiuhun=split(application("yx8_mhjh_qiuhun"),"|")
to1=qiuhun(0)
from1=qiuhun(1)
if Application("yx8_mhjh_qiuhun")="" or username<>to1 then%>
<script language=vbscript>
  MsgBox "别人没向你求婚啊!"
</script>
<%Response.End
else
message="<font color=FF0000>【求婚】</font><font color=000000>["&username& "]</font>断然拒绝了<font color=000000>[" & from1 & "]</font>的求婚！"
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
	newtalkarr(595)=from1
	newtalkarr(596)="" 
	newtalkarr(597)="000000" 
	newtalkarr(598)="000000" 
	newtalkarr(599)="<font color=red><marquee><b>"& message &"</b></marquee></font>"
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
application("yx8_mhjh_qiuhun")=""
Application.UnLock
erase newtalkarr
erase talkarr
%>
<script language=vbscript>
MsgBox "你拒绝了人家的求婚!"
</script>
<%end if%>

