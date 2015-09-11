<%@ LANGUAGE=VBScript%>
<%
Sub showchat(mess)
says=mess   '聊天数据
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & session("aqjh_name") & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","&session("nowinroom")&");<"&"/script>"
addmsg saystr
end sub
Function Yushu1(a)
 Yushu1=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu1(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>

<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
response.write "<br><br><center><table border=1 cellpadding=20 bordercolor=black bgcolor=#EEEEEE width=450>"
response.write "<tr><td style='font:9pt Verdana'>"
response.write "你提交的路径有误，禁止从站点外部提交数据请不要乱该参数！"
response.write "</td></tr></table></center>"
response.end
end if
useronlinename=Application("aqjh_useronlinename"&session("nowinroom"))
nickname=session("aqjh_name")
if Session("aqjh_name")="" then Response.Redirect "../error.asp?id=440"

id=Request.Form ("id")
if id="" then Response.end
towho=Request.Form ("towho")
if towho="" or towho=nickname or towho="大家" or Instr(useronlinename," "&towho&" ")=0 then 
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您好像送错了对像！！！\n  按 [确定] 返 回！');parent.history.go(-1);}</script>"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
rs.Open ("select * from myanimal where username='"&nickname&"' and rest=0 and id="&id),conn
if rs.BOF or rs.EOF then
Response.Write "<script language=JavaScript>{alert('  对不起！\n  您没有此种神龙，你是不是想作弊！！！\n  按 [确定] 返 回！');parent.ps.location.href='about:blank';}</script>"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.End 
else
animalname=rs("animalname")
conn.Execute ("update myanimal set username='"&towho&"',attack=attack/2 where username='"&nickname&"' and rest=0 and id="&id)
msg="<font color=black>["&nickname&"]将自己的神龙<font color=red>["&animalname&"]</font>送给了朋友<font color=red>["&towho&"]</font></font>！"
says=msg   '聊天数据
call showchat(says)

Response.Write "<script language=JavaScript>{alert('  赠送成功！！');parent.f3.location.href='myanimal.asp';parent.ps.location.href='about:blank';}</script>"
end if
%>
</body>
</html>
