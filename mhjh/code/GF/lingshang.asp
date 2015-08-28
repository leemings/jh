<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
name=request("name")
my=session("yx8_mhjh_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select 最后登录IP from 用户 where 姓名='" & my & "'"
set rs=conn.execute(sql)
ip=rs("最后登录IP")
if rs.eof or rs.bof then
%><script language=vbscript>
MsgBox "你不是剑侠中人!"
location.href = "../exit.ASP"
</script><%
conn.close
response.end
else
sql="select * from 用户 where 姓名='" & name & "'"
set rs=conn.execute(sql)
yin=rs("杀人")*20000
if rs("杀人者")<>my or rs("杀人者")="已结案" then%>
<script language=vbscript>
MsgBox "错误！这个通缉犯好象不是你杀的或是已经结案了。"
location.href = "tongji.asp"
</script>
<%else
if rs("姓名")=my then%>
<script language=vbscript>
MsgBox "错误！您自己杀自己啊，有没有搞错？"
location.href = "tongji.asp"
</script><%
else
if rs("最后登录IP")=ip then
sql="update 用户 set 杀人者='等候处砍' where 姓名='" & name & "'"
rs=conn.execute(sql)
%>
<script language=vbscript>
MsgBox "错误！有没有搞错，您杀的这个人正坐在你旁边呢，是不是想到这里来骗钱？"
location.href = "tongji.asp"
</script><%
conn.close
response.end
else
sql="update 用户 set 银两=银两+'" & yin & "' where 姓名='" & my & "'"
rs=conn.execute(sql)
sql="update 用户 set 杀人者='已结案',杀人=0 where 姓名='" & name & "'"
rs=conn.execute(sql)
mess=my & " 将被通缉的江湖恶棍<font color=ff0000> "& name &" </font>给宰了，领到了赏金"&yin&"两。"
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
newtalkarr(599)="<font color=red>【事件】</font><b><font color=red>" & mess & "</font></b>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
conn.close
Response.write ""&mess&"<b><p align=center><input type=button value=' 关 闭 ' onClick=javascript:window.close();></p>"
response.end
end if
end if
end if
rs.close
set rs=nothing
end if
%>



