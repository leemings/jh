<%
function touqian(un,st)
randomize()
odds=rnd()*4000
if odds>1999 then
rst.Open "select 积分 from 用户 where 姓名='"&un&"'",conn
uhp=rst("积分")
rst.Close
if odds*0.2>uhp then odds=uhp
conn.execute "update 用户 set 银两=银两-"&odds*0.2&",道德=道德-"&odds*1&" where 姓名='"&un&"'"
touqian="<font color=FF0000>【偷钱】</font>##竟然在官府人员的眼皮下对%%进行偷钱，有违江湖公道，被百姓们打得动弹不得！体力和道德下降"&clng(odds*1)&"，"&newaberrant(un,"大家","麻痹",120)
else
rst.Open "select 积分 from 用户 where 姓名='"&st&"'",conn
if rst.EOF or rst.BOF then
Response.Write "<script language=JavaScript>{alert('你真是疯了，连门派的金库都敢偷，别胡来，这是找死！');}</script>"
Response.End
      end if
smp=rst("积分")
rst.Close
if odds>smp then odds=smp
if odds>1999 then odds=smp*0.5
conn.execute "update 用户 set 银两=银两+"&odds*2&",道德=道德-"&odds*1&" where 姓名='"&un&"'"
conn.execute "update 用户 set 银两=银两-"&odds*2&" where 姓名='"&st&"'"
touqian="<font color=FF0000>【偷钱】</font>##竟然使用了江湖中失传已久的隔空取物秘技，偷走了%%"&clng(odds)*2&"的银两,道德降低"&odds*1&"点。<bgsound src='../mid/xixing.wav' loop=1>"
end if
end function

%>
