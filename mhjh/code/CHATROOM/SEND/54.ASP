<%
function leave(un)
f=Minute(time())
if Instr(Application("yx8_mhjh_zlname"),";"&un&";")<>0 then
Response.Write "<script Language=Javascript>alert('你现在正是暂时离开的状态呀！ :P');</script>" 
Response.end 
elseif f>30 then
Response.Write "<script language=JavaScript>{alert('现在是杀人时间[每个小时的后30分钟]，禁止离线，不想死就去官府吧！)，');}</script>"
   Response.End
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
rs.open "select 操作时间,杀人,保护 from 用户 where 姓名='"&un&"'",conn
if rs("杀人")>9 then
Response.Write "<script language=JavaScript>{alert('你杀人总数到了10个,成了通缉犯，禁止离线！');}</script>"
   Response.End
exit function 
end if
if Instr(Application("yx8_mhjh_zlname"),";"&un&";")=0 then
Application.Lock
Application("yx8_mhjh_zlname")=Application("yx8_mhjh_zlname")&un&";"
Application.UnLock
leave="<font color=FF0000>【离开】</font>##暂时离开了座位......"
end if
end function
%>


