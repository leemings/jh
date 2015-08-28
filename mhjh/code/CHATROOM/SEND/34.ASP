<%
function impart(un,st,mg)
if not isnumeric(mg) then
impart="<font color=FF0000>【传功】</font>##想你的确是传功"&mg&"给%%吗？"
else
mg=clng(mg)
if mg<100 then
impart="<font color=FF0000>【传功】</font>##想你的确是传功"&mg&"给%%吗？也太少了点儿吧"
else
rst.Open "select 操作时间 from 用户 where 姓名='"&un&"' and 内力>="&mg,conn
if rst.EOF or rst.BOF then
impart="<font color=FF0000>【传功】</font>%%笑着对##说：'其实你只传给我几百我也会很开心的不用那么"&mg&"多！'"
else
  sj=DateDiff("s",rst("操作时间"),now())
if sj<60 then
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒再传！');}</script>"
	Response.End
	exit function
end if
conn.execute "update 用户 set 内力=内力+"&mg*0.8&" where 姓名='"&st&"'"
conn.execute "update 用户 set 内力=内力-"&mg&",道德=道德+"&mg*0.01&",操作时间=now() where 姓名='"&un&"'"
impart="<font color=FF0000>【传功】</font>##将"&clng(mg*0.8)&"的内力传授给了%%,魔幻之神奖励##"&mg*0.01&"点道德。<bgsound src='../mid/xixing.wav' loop=1>"
end if
end if
end if
end function
%>
