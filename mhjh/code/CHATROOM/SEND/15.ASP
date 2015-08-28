<%
function bais(un,st,gr)
rst.Open "select 银两,师傅 from 用户 where 姓名='"&un&"'",conn
yl=rst("银两")
shfu=rst("师傅")
rst.Close
if yl<50000 then 
	Response.Write "<script language=JavaScript>{alert('你连5万两也没有，你还想拜师?');}</script>"
	Response.End
	exit function
end if
if shfu<>"无" then
	Response.Write "<script language=JavaScript>{alert('你已经有师傅了');}</script>"
	Response.End
	exit function
end if
rst.Open "select 等级,师傅,徒弟 from 用户 where 姓名='"&st&"'",conn
dej=rst("等级")
sf=rst("师傅")
td=rst("徒弟")
rst.Close
if td=un then
	Response.Write "<script language=JavaScript>{alert('对方是你的师傅了，你还拜师?');}</script>"
	Response.End
	exit function
end if
if gr<2 then
	Response.Write "<script language=JavaScript>{alert('你等级还不够2级，不能拜师');}</script>"
	Response.End
	exit function
end if
if gr>dej then
	Response.Write "<script language=JavaScript>{alert('你等级比对方高，不能拜师');}</script>"
	Response.End
	exit function
end if
if sf=un then
	Response.Write "<script language=JavaScript>{alert('对方是你徒弟啊，怎么能拜师');}</script>"
	Response.End
	exit function
end if
conn.execute "update 用户 set 银两=银两-50000 where 姓名='"&un&"'"
bais="<font color=FF0000>【拜师】</font>##准备了5万块拜师费，向%%提出了拜师申请，同意不同意就要看%%了！"
Application.Lock
Application("yx8_mhjh_sf")=st
Application("yx8_mhjh_td")=un
Application.UnLock
end function
%>
