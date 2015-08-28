<%Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Response.AddHeader "Pragma", "No-Cache"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016" 
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
 Response.Write "<script Language=javascript>alert('提示：禁止站外提交！');</script>"
 Response.Write "<script Language=javascript>location.href = 'option.asp';</script>"
 Response.end
end if 
dx=LCase(trim(request.querystring("dx")))
ddd=LCase(trim(request.querystring("ddd")))
yl=LCase(trim(request.querystring("yl")))
if ddd<>username then
 Response.Write "<script Language=javascript>alert('提示：不要开人家的红包嘛！');</script>"
 Response.Write "<script Language=javascript>location.href = 'option.asp';</script>"
 Response.end
end if 
if not isnumeric(yl) then 
 Response.Write "<script language=javascript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
 Response.End
end if
if yl<50000000 then
qssok="<font color=FF0000>【开红包】</font>"& username &"接过了"& dx &"的的红包,慢慢走回家,快快打开一看,有"&yl&"两<img src=../image/252.GIF>,还可以吧！"
else
qssok="<font color=FF0000>【开红包】</font>"& username &"接过了"& dx &"的的红包,打开一看,哇！竟有"&yl&"两<img src=../image/251.GIF>啊！"
end if
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
 newtalkarr(599)=""&qssok&"<font class=timsty>("&time()&")</font>"
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<script Language=javascript>location.href = "option.asp";</script>