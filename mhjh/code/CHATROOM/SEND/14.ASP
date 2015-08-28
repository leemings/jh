<%
function sendmoney(un,st,mg)
rst.Open "select 等级 from 用户 where 姓名='"&st&"'",conn
  dj=rst("等级")
  rst.close
  if dj<=3 and mg>100000000 then
	Response.Write "<script language=JavaScript>{alert('对方等级太低，你不用送1亿这么大方吧！');}</script>"
	Response.End
	exit function
  end if
if not isnumeric(mg) then
sendmoney="<font color=FF0000>【送钱】</font>##想你的确是送钱"&mg&"给%%吗？"
else
mg=clng(mg)
if mg<1 then
sendmoney="<font color=FF0000>【送钱】</font>##想你的确是送钱"&mg&"给%%吗？"
elseif mg>100000000 and session("yx8_mhjh_usergrade")<50 and session("yx8_mhjh_usercorp")<>"官府"then
sendmoney="<font color=FF0000>【送钱】</font>##现在官府禁止送钱超过10000万，谢谢合作"
else
rst.Open "select 银两,等级,操作时间 from 用户 where 姓名='"&un&"' and 银两>="&mg,conn
if rst.EOF or rst.BOF then
sendmoney="<font color=FF0000>【送钱】</font>%%笑着对##说：'你的心情我理解，等你有了"&mg&"两银子再来送我好吗？'"
else
    sj=DateDiff("s",rst("操作时间"),now())
dj=rst("等级")
    rst.close
if sj<3 then
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('请你等上["& ss &"]秒再操作！');}</script>"
	Response.End
exit function
	end if
if dj<10 then
	Response.Write "<script language=JavaScript>{alert('你还是小孩子,你哪来的钱给别人?自己留着吧,起码也得到10岁再送钱吧！');}</script>"
	Response.End
exit function
	end if
conn.execute "update 用户 set 银两=银两+"&mg*0.97&" where 姓名='"&st&"'"
conn.execute "update 用户 set 操作时间=now(),银两=银两-"&mg&" where 姓名='"&un&"'"
sendmoney="<font color=FF0000>【送钱】</font>##将"&mg&"两银子送给了%%，并向官府救灾中心交纳税金3%<bgsound src='../mid/thanks.wav' loop=1>"
end if
end if
end if
end function
%>