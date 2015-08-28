<%
function comeback(un)
if Instr(Application("yx8_mhjh_zlname"),";"&un&";")=0 then
Response.Write "<script Language=Javascript>alert('你没有启动过暂时离开的功能呀！ :P');</script>" 
Response.end 
exit function 
end if
if Instr(Application("yx8_mhjh_zlname"),";"&un&";")<>0 then
zllist=Application("yx8_mhjh_zlname")
Application.Lock
Application("yx8_mhjh_zlname")=replace(zllist,";"&un&";",";")
Application.UnLock
comeback="<font color=FF0000>【回来】</font>##回到了自己座位,向大家高兴的喊道：各位好，我回来了!"
end if
end function
%>
