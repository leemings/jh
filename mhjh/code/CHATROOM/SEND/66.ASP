<%
function junzhul(un,st,mg)
if len(mg)<50 then
  rst.Open "select 国家 from 用户 where 姓名='"&un&"' and 职位='君主'",conn
if rst.EOF or rst.BOF then
junzhul="<font color=FF0000>【君主令】</font>##,你不是本国君主，无法发令箭"
else
guojia=rst("国家")
rst.Close
rst.Open "select id from 用户 where 姓名='"&st&"' and 国家='"&guojia&"' and 职位<>'君主'",conn
 if rst.EOF or rst.BOF then
junzhul="<font color=FF0000>【逐出】</font>%%不是本国子民，所以##你无权发令箭！"
 else
		junzhul="<font color=FF0000>【君主令】</font>##对%%发出圣旨<table align=center><td><table border=0 cellpadding=0 cellspacing=0 ><tr><td><img src=../chatroom/image/rightct3.gif ></td><td background=../chatroom/image/rightct4.gif ></td> <td><img  src=../chatroom/image/rightct1.gif ></td></tr><tr><td background=../chatroom/image/rightct8.gif ></td><td valign=center align=center><table align=center border=0 cellpadding=1 cellspacing=0 ><tr> <td valign=center align=center><font color=red>君主令</font></td> </tr> <tr><td valign=center align=center >"&msg&"</td></tr> <tr><td valign=center align=center ><font color=blue>"&guojia&"国君-"&un&"</font></td></tr></table></td><td background=../chatroom/image/rightct08.gif></td></tr><tr><td><img src=../chatroom/image/rightct9.gif></td><td background=../chatroom/image/rightct10.gif></td><td><img src=../chatroom/image/rightct11.gif></td></tr></table></td></table></body>"
 end if	
 rst.Close
end if
else
junzhul="<font color=FF0000>【君主令】</font>##,君主发令箭只能在五十字以内!"
end if
end function
%>