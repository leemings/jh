<%
function fengzhou(un,st,co,gr,sn,mg)
if co="官府" and gr>=19 then
fengzhou="<font color=FF0000>【封咒】</font>%%,你刚刚干的太过分了了,封你咒以示公正"&newaberrant(st,"大家","封咒",120)
conn.execute "insert into 聊务记录(管理类型,管理人员,被管理人,管理原因,管理时间,管理房间) values ('封咒','" & un & "','" & st & "','" & mg & "',now(),'"& Application("yx8_mhjh_systemname"&chatroomsn) &"')"
else
fengzhou="<font color=FF0000>【封咒】</font>%%,##向你封咒可惜无效"
end if
end function
%>
