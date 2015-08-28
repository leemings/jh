<%
function studi(un,st,gr)
rst.Open "select 徒弟 from 用户 where 姓名='"&un&"'",conn
sfa=rst("徒弟")
rst.Close
if sfa<>"无" then
	Response.Write "<script language=JavaScript>{alert('你已经有徒弟，不能再收了');}</script>"
	Response.End
	exit function
end if
rst.Open "select 等级,师傅,徒弟 from 用户 where 姓名='"&st&"'",conn
dej=rst("等级")
sf=rst("师傅")
td=rst("徒弟")
rst.Close
if sf=un then
	Response.Write "<script language=JavaScript>{alert('对方已经是你的徒弟，不能再收他了');}</script>"
	Response.End
	exit function
end if
if sf<>"无" then
	Response.Write "<script language=JavaScript>{alert('对方已经有师傅了，不能再收他了');}</script>"
	Response.End
	exit function
end if
if trim(Application("yx8_mhjh_sf"))<>trim(un) or trim(Application("yx8_mhjh_td"))<>trim(st) then
	Response.Write "<script language=JavaScript>{alert('对方也没有想拜你为师傅，你也太那个了吧！');}</script>"
	Response.End
	exit function
end if
if gr<4 then
	Response.Write "<script language=JavaScript>{alert('你的等级还不够4级');}</script>"
	Response.End
	exit function
end if
if gr<dej then
	Response.Write "<script language=JavaScript>{alert('你的等级比对方还要低，不能收他为徒弟');}</script>"
	Response.End
	exit function
end if
if td=un then
	Response.Write "<script language=JavaScript>{alert('对方你的师傅啊，你想玩什么啊？');}</script>"
	Response.End
	exit function
end if
conn.execute "update 用户 set 师傅='"&un&"'  where 姓名='"&st&"'"
conn.execute "update 用户 set 银两=银两+50000,徒弟='"&st&"'  where 姓名='"&un&"'"
studi="<font color=FF0000>【收徒】</font>%%向##交纳了5万块拜师费，又是点头又是哈腰的，终于求得收自己为徒,嘿嘿，有师傅在，啥也不怕！"
Application.Lock
Application("yx8_mhjh_sf")=""
Application("yx8_mhjh_td")=""
Application.UnLock
end function
%>
