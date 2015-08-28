<%
function dianxue(user,sent,mycorp)
rst.open "select 精力,攻击,积分,内力,等级,身份,会员 from 用户 where 姓名='"&user&"'",conn
jingliu=rst("精力")
wugongu=rst("攻击")
jifenu=rst("积分")
nl=rst("内力")
dji=rst("等级")
sf=rst("身份")
hy=rst("会员")
rst.close
rst.open "select 攻击,门派 from 用户 where 姓名='"&sent&"'",conn
if rst.EOF or rst.BOF then
Response.Write "<script language=JavaScript>{alert('你真是疯了，连门派的金库都敢偷，别胡来，这是找死！');}</script>"
Response.End
      end if
wugongs=rst("攻击")
menpai=rst("门派")
rst.Close
if jingliu<100000 then
 Response.Write "<script language=JavaScript>{alert('精力不够!');}</script>"
 Response.end
elseif dji<30 then
 Response.Write "<script language=JavaScript>{alert('你等级小于30');}</script>"
 Response.end
elseif nl<2000000 then
 Response.Write "<script language=JavaScript>{alert('你连200万内力都没有，还在这里充什么高手，点穴功夫你还不会用呢！');}</script>"
 Response.end
elseif jifenu<100000 then
 Response.Write "<script language=JavaScript>{alert('你还是一个小辈,等你有了100000积分再来吧');}</script>"
 Response.end
elseif wugongu<=wugongs then
 Response.Write "<script language=JavaScript>{alert('就凭你也想点人家');}</script>"
 Response.end
elseif menpai="官府" then
 Response.Write "<script language=JavaScript>{alert('官府的都敢来，想造反吗');}</script>"
 Response.end
elseif not(mycorp="官府") and hy<>True then
 Response.Write "<script language=JavaScript>{alert('你不是官府的，不能点穴');}</script>"
 Response.end
else
 if instr(aberrantname,";"&sent&";")=0  then
 	especial="点穴"
	especialtime=180
 conn.execute "update 用户 set 精力=精力-2000000,内力=内力-20000000 where 姓名='" & user & "'"
 dianxue="<font color=green>【点穴】</font>##运动本身的200万内力，耗费200万精力，食指变得通红，一股厉气向%%射去，%%躲闪不及，乖乖的被点在那里，60秒无法动弹。%%乖乖的不动了,哎,聪明反被聪明误！"&newaberrant(sent,"大家",especial,especialtime)
 else
 Response.Write "<script language=JavaScript>{alert('他已经被点穴了，还想欺负他呀!');}</script>"
 Response.end
 end if
end if
rst.close
end function
%>