<%
function card(un,st,mg,co)
if co="十八地狱" or mg="风铃" then 
       Response.Write "<script language=JavaScript>{alert('如果你形态是鬼魂或者你对别人使用风铃,都是禁止的！');}</script>"
       Response.End
       exit function
end if	
if Application("yx8_mhjh_fellow")=True then
rst.Open "select tc.* from card tc inner join 用户 tu ON tu.姓名=tc.username where tc.username='"&un&"' and tc.cname='"&mg&"' and tc.cnumber>0 and tu.会员=true",conn
else
rst.Open "select * from card where username='"&un&"' and cname='"&mg&"' and cnumber>0",conn
end if
if rst.EOF or rst.BOF then
card="<font color=FF0000>【道具】</font>##，你并没有"&mg&"可对%%使用！"
else
especial=trim(rst("cespecial"))
especialtime=rst("ctime")
esptime=dateadd("s",especialtime,now())
rst.Close
rst.Open "select 特技 from 用户 where 姓名='"&st&"'"
if rst.EOF or rst.bof then Response.End
if instr(rst("特技"),";抗"&especial&";")<>0 then
card="<font color=FF0000>【道具】</font>##对%%使用了"&mg&",但是%%具有抗"&especial&"特性，因而没有任何效果"
else
card="<font color=FF0000>【道具】</font>##对%%使用了"&mg&newaberrant(st,un,especial,especialtime)
end if
conn.execute "update card set cnumber=cnumber-1 where username='"&username&"' and cname='"&mg&"'"
if instr(";死亡;入狱;逮捕;",especial)<>0 then conn.execute "update 用户 set 状态='"&especial&"',最后登录时间='"&esptime&"' where 姓名='"&st&"'"
end if
rst.Close
end function
%><!--#include file="60.asp"-->