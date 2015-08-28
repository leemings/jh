<%
function bulletin(un,st,co,gr,sn,mg)
	if co="官府" and gr>=40  then
		bulletin="<font color=FF0000 size=3>【公告】</font><table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chatroom/image/rightct3.gif ></td><td background=../chatroom/image/rightct4.gif ></td> <td><img  src=../chatroom/image/rightct1.gif ></td></tr><tr><td background=../chatroom/image/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr> <td valign=center align=center>皇城公告</td> </tr> <tr><td valign=center align=center >"&msg&"</td></tr></table></td><td background=../chatroom/image/rightct08.gif></td></tr><tr><td><img src=../chatroom/image/rightct9.gif></td><td background=../chatroom/image/rightct10.gif></td><td><img src=../chatroom/image/rightct11.gif></td></tr></table></td></table>"
	else
		bulletin="<font color=FF0000>【公告】</font>##，你无权发告公告:"&mg
	end if
end function
%>

