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
sql="update 申冤 set 结果='0' where id=" & id & ""
conn.execute sql
newer="<font color=blank>" & my & "</font>状告<font color=blue>" & bg & "</font>,要求" & result & ",因证据不足、理由不充分，未通过！"
conn.close
set rs=nothing	
set conn=nothing
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

Response.Redirect "manage.asp"
%>