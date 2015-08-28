<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
youname=session("yx8_mhjh_username")
if youname="" then  Response.End
chatroomsn=session("yx8_mhjh_userchatroomsn")
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
rs.open "select 操作时间,杀人,保护 from 用户 where 姓名='"&youname&"'",conn
sj=DateDiff("s",rs("操作时间"),now())
f=Minute(time())
ss=140-sj
if rs("保护")=True then
   Response.Write "<script language=JavaScript>{alert('你已经开了保护，不需要再开');}</script>"
   Response.End
elseif f>59 then
   Response.Write "<script language=JavaScript>{alert('现在是杀人时间[每个小时的后30分钟]，禁止打开保护！)，');}</script>"
   Response.End
elseif rs("杀人")>9 then
   Response.Write "<script language=JavaScript>{alert('你杀人总数到了10个,成了通缉犯，禁止打开保护！');}</script>"
   Response.End
elseif sj<140 then
   Response.Write "<script language=JavaScript>{alert('你刚刚开启过一次保护了，要再开启就再等["& ss &"]秒吧');}</script>"
   Response.End
else
   conn.Execute "update 用户 set 保护=True,操作时间=now(),体力=体力-10000 where 姓名='" & youname & "'"
   diaox="的保护模式开启，谁也打不了！体力因此下降10000"
   abc=1
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
if abc=1 then
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=Application("yx8_mhjh_talkpoint")
dim newtalkarr(600)
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
newtalkarr(599)="<font color=red>【保护】</font><b><font color=red>"&youname&""&diaox&"</font></b>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
   Response.Write "<script language=JavaScript>{alert('保护模式开启，谁也打不了');}</script>"
   Response.End
end if
%>
