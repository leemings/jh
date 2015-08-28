<%
function dazuo(un,st)
rst.Open "select 操作时间,门派,内力 from 用户 where 姓名='"&un&"'",conn
mp=rst("门派")
nl=rst("内力")
cz=rst("操作时间")
rst.Close
if mp<>"出家" then
  		Response.Write "<script language=JavaScript>{alert('你并没有出家，呵呵还不是一个和尚，打什么坐呢');}</script>"
		Response.End
        exit function
 end if
if nl>50000000 then
  		Response.Write "<script language=JavaScript>{alert('晕，想吓死人啊，内力这么厉害了，还练');}</script>"
		Response.End
        exit function
 end if
if nl<50000 then
  		Response.Write "<script language=JavaScript>{alert('你的身体太差了，回去练练再来吧！体力不能少于50000哦');}</script>"
		Response.End
        exit function
 end if
sj=DateDiff("s",cz,now())
if sj<5 then
	ss=5-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒再打坐！');}</script>"
	Response.End
else
randomize()
nnn=clng(rnd()*200)+100
conn.execute "update 用户 set 内力=内力+"&nnn&",操作时间=now()  where 姓名='"&un&"'"
dazuo="<font color=FF0000>【打坐】</font>经过长时间的打坐，##的内力增加"&nnn&"点"
end if
end function
%>
