<%
function dakailt(un,st,co,gr)
if gr>=50 and co="官府" then
if Application("yx8_mhjh_chatroomsnkb")<>1 then
Application("yx8_mhjh_chatroomsnkb")=1
dakailt="<font color=FF0000>【开打】</font>##开启聊天室战斗模式成功！大家可以比武论英雄了"
else
dakailt="<font color=FF0000>【开打】</font>开启聊天室战斗模式失败！原因已经开启"
end if
else
dakailt="<font color=FF0000>【开打】</font>##不是官府的，呵呵！"
end if
end function
%>