<%
function zhanshou(un,st,co,gr,sn,mg)
if gr>=50 and co="官府" then
if mg="" then
zhanshou="<font color=FF0000>【斩首】</font>##,必须输入斩首原因才可将%%在就地正法！"
else
rst.Open "select 状态 from 用户 where 姓名='"&st&"'",conn
cond=rst("状态")
if cond="死亡" then
zhanshou="<font color=FF0000>【斩首】</font>%%业已死亡，不需要再处斩了！呵呵"
else
	especial="死亡"
	especialtime=800
	esptime=dateadd("s",especialtime,now())
zhanshou="<font color=FF0000>【斩首】</font>由于%%实在罪大恶极，##将%%就地正法！原因是:"&mg&newaberrant(st,"大家",especial,especialtime)
if instr(";死亡;入狱;逮捕;",especial)<>0 then conn.execute "update 用户 set 状态='死亡',体力=体力-100000,内力=内力-100000,道德=道德-5000,等级=等级-1,最后登录时间='"&esptime&"',被杀=被杀+1,杀人者='"&un&"' where 姓名='"&st&"'"
conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','被斩首:"&mg&"')"
conn.execute "insert into 聊务记录(管理类型,管理人员,被管理人,管理原因,管理时间,管理房间) values ('斩首','" & un & "','" & st & "','" & mg & "',now(),'"& Application("yx8_mhjh_systemname"&chatroomsn) &"')"
end if
end if
else
zhanshou="<font color=FF0000>【斩首】</font>##，对不起，您输入的命令无法执行，原因不明命令或您无此命令的执行权限"
rst.Close
end if
end function
%>
