<%
function gbilt(un,st,co,gr)
if gr>=50 and co="官府" then
if chatroomsn<>3 then
if  Application("yx8_mhjh_chatroomsnkb")<>0 then
Application("yx8_mhjh_chatroomsnkb")=0
gbilt="<font color=FF0000>【关打】</font>关闭聊天室战斗模式成功！"
else
gbilt="<font color=FF0000>【关打】</font>关闭聊天室战斗模式失败！原因：已经关闭"
end if
else
gbilt="<font color=FF0000>【关打】</font>关闭聊天室战斗模式失败，原因：比武大厅不能关闭动武！"
end if
else
gbilt="<font color=FF0000>【关打】</font>##不是官府的，呵呵！"
end if
end function
%>
