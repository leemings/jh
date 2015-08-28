<%
function jiebai(un,st,gr)
rst.Open "select 结拜兄弟 from 用户 where 姓名='"&st&"'",conn
jib=rst("结拜兄弟")
rst.Close
if jib<>"无" then
	Response.Write "<script language=JavaScript>{alert('你已经有兄弟了，你还结拜作什么？');}</script>"
	Response.End
	exit function
end if
if gr<2 then
	Response.Write "<script language=JavaScript>{alert('你的等级还不够2级，不能结拜');}</script>"
	Response.End
	exit function
end if
jiebai="<font color=FF0000>【结拜】</font>##向%%提出了结拜申请，实行桃源结义，同意不同意就要看%%了！"
Application.Lock
Application("yx8_mhjh_jb1")=st
Application("yx8_mhjh_jb2")=un
Application.UnLock
end function
%>
