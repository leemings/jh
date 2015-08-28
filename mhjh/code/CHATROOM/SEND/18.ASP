<%
function jiebaich(un,st)
rst.Open "select 结拜兄弟 from 用户 where 姓名='"&st&"'",conn
jb=rst("结拜兄弟")
rst.Close
if jb=un then
	Response.Write "<script language=JavaScript>{alert('对方已经是你的结拜兄弟了');}</script>"
	Response.End
	exit function
end if
if trim(Application("yx8_mhjh_jb1"))<>trim(un) or trim(Application("yx8_mhjh_jb2"))<>trim(st) then
	Response.Write "<script language=JavaScript>{alert('对方也没有想和你结拜，你也太那个了吧！');}</script>"
	Response.End
	exit function
end if
conn.execute "update 用户 set 结拜兄弟='"&un&"' where 姓名='"&st&"'"
conn.execute "update 用户 set 结拜兄弟='"&st&"' where 姓名='"&un&"'"
jiebaich="<font color=FF0000>【接拜】</font>##和%%终于效仿古人实行了桃源结义，多么畅快阿！"
Application.Lock
Application("yx8_mhjh_jb1")=""
Application("yx8_mhjh_jb2")=""
Application.UnLock
end function
%>
