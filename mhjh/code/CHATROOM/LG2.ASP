<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
youname=session("yx8_mhjh_username")
if youname="" then   Response.End
chatroomsn=Session("yx8_mhjh_userchatroomsn")
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT * FROM 用户 WHERE 姓名='" & youname & "'"
Set Rs=conn.Execute(sql)
if rs("保护")=true then
sql="update 用户 set 保护=false where 姓名='" & youname & "'"
Set Rs=conn.Execute(sql)
diaox="的保护模式关闭！"
abc=1
else
dx="保护模式已经关闭了，不需要再关闭！"
  Response.Write "<script language=JavaScript>{alert('" & dx & "');}</script>"
  Response.End
end if
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
newtalkarr(599)="<font color=red>【保护】</font><b><font color=red>" & youname & ""& diaox &"</font></b>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
  Response.Write "<script language=JavaScript>{alert('保护模式关闭');}</script>"
  Response.End
end if
%>