<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("awjh_l_usermdb")
if info(5)<>"官府" or info(2)<10  then Response.Redirect "../error.asp?id=439"
result=request("result")
id=request("id")
bg=request("bg")
my=request("my")
if result="罚款一万" then
sql="update 用户 set 银两=银两-10000 where 姓名='"&bg&"'"
conn.Execute(sql)
sql="update 用户 set 银两=银两+10000 where 姓名='"&my&"'"
conn.Execute(sql)
end if
if result="罚款十万" then
sql="update 用户 set 银两=银两-100000 where 姓名='"&bg&"'"
conn.Execute(sql)
sql="update 用户 set 银两=银两+100000 where 姓名='"&my&"'"
conn.Execute(sql)
end if
if result="坐牢" then
sql="update 用户 set 状态='牢',登录=now()+1/144 where 姓名='" & bg & "'"
conn.execute sql
end if

sql="update 申冤 set 结果='1' where id=" & id & ""
conn.execute sql
newer="<font color=blank>" & my & "</font>状告<font color=blue>" & bg & "</font>,要求" & result & ",通过申议！"
Application.Lock
sd=Application("awjh_l_sd")
line=int(Application("awjh_l_line"))
Application("awjh_l_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="660099"
sd(197)="660099"
sd(198)="对"
sd(199)="<font color=red>【申冤判决】</font><font color=blank><b>"& newer &"</b></font>"
sd(200)=session("nowinroom")
Application("awjh_l_sd")=sd
Application.UnLock

'踢人
sub boot(bg)
Application.Lock
onlinelist=Application("awjh_l_onlinelist")
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(bg) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("awjh_l_onlinelist")=listnull
else
Application("awjh_l_onlinelist")=newonlinelist
end if
Application("awjh_l_useronlinename")=useronlinename
Application("awjh_l_chatrs")=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
conn.close
set rs=nothing	
set conn=nothing
Response.Redirect "manage.asp"
%>