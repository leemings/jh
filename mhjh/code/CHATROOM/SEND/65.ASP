<%
function zhangml(un,st,co,mg)
if len(mg)<40 then
  rst.Open "select id from 用户 where 姓名='"&un&"' and 门派='"&co&"' and 身份='掌门'",conn
if rst.EOF or rst.BOF then
zhangml="<font color=FF0000>【掌门令】</font>##,你不是本派掌门，无法发令箭"
else
rst.Close
rst.Open "select id from 用户 where 姓名='"&st&"' and 门派='"&co&"' and 身份<>'掌门'",conn
 if rst.EOF or rst.BOF then
zhangml="<font color=FF0000>【逐出】</font>%%不是本派弟子，所以##你无权发令箭！"
 else
		zhangml="<font color=FF0000>【掌门令】</font>##对堂下的%%发出号令<table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chatroom/image/rightct3.gif ></td><td background=../chatroom/image/rightct4.gif ></td> <td><img  src=../chatroom/image/rightct1.gif ></td></tr><tr><td background=../chatroom/image/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr> <td valign=center align=center><font color=red>掌门令</font></td> </tr> <tr><td valign=center align=center >"&msg&"</td></tr> <tr><td valign=center align=center ><font color=blue>"&un&"</font></td></tr></table></td><td background=../chatroom/image/rightct08.gif></td></tr><tr><td><img src=../chatroom/image/rightct9.gif></td><td background=../chatroom/image/rightct10.gif></td><td><img src=../chatroom/image/rightct11.gif></td></tr></table></td></table></body>"
 end if	
 rst.Close
end if
else
zhangml="<font color=FF0000>【掌门令】</font>##,掌门发令箭只能在四十字以内!"
end if
end function
%>