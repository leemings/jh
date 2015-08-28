<%
function bangfu(un,st,mg)
rst.Open "select 积分,最后登录IP from 用户 where 姓名='"&st&"'",conn
  dj=rst("积分")
  dl=rst("最后登录IP")
  rst.close
  if dj>500 then
	Response.Write "<script language=JavaScript>{alert('对方不是新人了，你不能再救济他了！');}</script>"
	Response.End
	exit function
  end if
if not isnumeric(mg) then
bangfu="<font color=FF0000>【帮扶新人】</font>##想你的确是送钱"&mg&"给%%吗？"
else
mg=clng(mg)
if mg<1000000 then
bangfu="<font color=FF0000>【帮扶新人】</font>##想你的确是送钱"&mg&"给%%吗？至少100万，别太吝啬！"
elseif mg>3000000 then
bangfu="<font color=FF0000>【帮扶新人】</font>##现在官府禁止送钱给新人超过300万，谢谢你的好心！"
else
rst.Open "select 银两,操作时间,最后登录IP from 用户 where 姓名='"&un&"' and 银两>="&mg,conn
if rst.EOF or rst.BOF then
bangfu="<font color=FF0000>【帮扶新人】</font>%%笑着对##说：'你的心情我理解，等你有了"&mg&"两银子再来送我好吗？'"
else
    sj=rst("最后登录IP")
    rst.close
if sj=dl then
	Response.Write "<script language=JavaScript>{alert('相同或者相近的IP无法进行帮扶新人活动！');}</script>"
	Response.End
exit function
	end if
conn.execute "update 用户 set 银两=银两+"&mg*0.97&",积分=积分+500,体力=体力+"&mg/2&",内力=内力+"&mg/5&" where 姓名='"&st&"'"
conn.execute "update 用户 set 银两=银两-"&mg&",体力=体力-"&mg/2&",内力=内力-"&mg/5&",防御=防御+"&mg/50&",道德=道德+"&mg/50&" where 姓名='"&un&"'"
bangfu="<font color=FF0000>【帮扶新人】</font>##发挥魔幻江湖的优良传统，热心帮助新人，把自己的"&mg&"两银子、"&mg/2&"体力、"&mg/5&"内力送给了%%，而他自己获得官府奖励防御"&mg/50&"，道德"&mg/50&"，向官府救灾中心交纳税金3%<bgsound src='../mid/thanks.wav' loop=1>"
end if
end if
end if
end function
%>