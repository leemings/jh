<%
function burglemoney(un,st)
randomize()
odds=rnd()*10000
rst.Open "select * from 用户 where 姓名='"&st&"'",conn
if rst.EOF or rst.BOF then
burglemoney="<font color=FF0000>【偷窃】</font>##你想对谁偷窃啊，"&newaberrant(un,"大家","中毒",360)
elseif rst("保护")=True then
burglemoney="<font color=FF0000>【偷窃】</font>##太傻了,居然在%%已经开启保护的情况下实施武力抢劫，被%%的保护真气震得晕头转向，"&newaberrant(un,"大家","中毒",360)
else
rst.Close
rst.Open "select 积分,保护,门派,形态 from 用户 where 姓名='"&un&"'",conn
uhp=rst("积分")
bh=rst("保护")
ump=rst("门派")
xt=rst("形态")
rst.Close
if uhp<10000 then
burglemoney="<font color=FF0000>【偷窃】</font>##还是个小孩子,居然敢在公开场合对%%实施武力抢劫，有违江湖公道，被大家打傻了，"&newaberrant(un,"大家","麻痹",180)
elseif ump="出家" then
burglemoney="<font color=FF0000>【偷窃】</font>##是个出家人,居然敢对%%实施武力抢劫，有违江湖公道，被方丈踢了一脚，"&newaberrant(un,"大家","麻痹",180)
elseif xt="鬼魂" then
burglemoney="<font color=FF0000>【偷窃】</font>##是个鬼魂,居然大白天敢对%%实施武力抢劫，有违江湖公道，被阎王打了一巴掌，"&newaberrant(un,"大家","麻痹",180)
elseif bh=true then
burglemoney="<font color=FF0000>【偷窃】</font>##处于保护中,居然蒙着眼睛敢对%%实施武力抢劫，真是愚蠢啊，被大家打傻了，"&newaberrant(un,"大家","麻痹",180)
elseif odds>3999 then
if odds>uhp then odds=uhp
conn.execute "update 用户 set 体力=体力-"&odds&",道德=道德-"&odds&" where 姓名='"&un&"'"
burglemoney="<font color=FF0000>【偷窃】</font>##因在公众场合对%%实施武力抢劫，有违江湖公道，被大家打傻了，不会出声了，体力和道德下降"&odds&"，"&newaberrant(un,"大家","麻痹",180)
elseif odds>0 and odds<4000 then
rst.Open "select * from 物品  where 所有者='"&st&"' and 数量>0 order by 价格 desc",conn
if rst.EOF or rst.BOF then
odds=clng(odds*0.2)
if odds>uhp then odds=uhp
conn.execute "update 用户 set 体力=体力-"&odds&",道德=道德-"&odds&" where 姓名='"&un&"'"
burglemoney="<font color=FF0000>【偷窃】</font>%%形容落魄，身无长物，##居然想对之进行偷窃，被大家打的遍体鳞伤，体力和道德下降"&odds&"，"&newaberrant(un,"大家","封咒",180)
else
cname=rst("名称")
set rst2=server.CreateObject("adodb.recordset")
rst2.Open "select * from 物品 where 所有者='"&un&"' and 名称='"&cname&"'",conn
if rst2.EOF or rst2.BOF then
conn.execute "insert into 物品(名称,属性,体力,内力,攻击,防御,特效,数量,价格,所有者,寄售,装备) values('"&cname&"','"&rst("属性")&"',"&rst("体力")&","&rst("内力")&","&rst("攻击")&","&rst("防御")&",'"&rst("特效")&"',1,"&rst("价格")&",'"&un&"',false,false)"
else
conn.execute "update 物品 set 数量=数量+1 where 所有者='"&un&"' and 名称='"&cname&"'"
end if
conn.execute "update 物品 set 数量=数量-1 where 所有者='"&st&"' and 名称='"&cname&"'"
rst2.Close
set rst2=nothing
burglemoney="<font color=FF0000>【偷窃】</font>##偷偷使出妙手空空的手段，偷走了%%<bgsound src='../mid/xiaotou.wav' loop=1>随身携带的"&cname
end if
rst.Close
else
rst.Open "select 银两 from 用户 where 姓名='"&st&"'",conn
umoney=rst("银两")
rst.Close
if odds>umoney then odds=umoney
if odds>1900 then odds=umoney*0.5
conn.execute "update 用户 set 银两=银两+"&odds&",道德=道德-"&odds*0.01&" where 姓名='"&un&"'"
conn.execute "update 用户 set 银两=银两-"&odds&" where 姓名='"&st&"'"
burglemoney="<font color=FF0000>【偷窃】</font>##偷偷使出妙手空空的手段，偷走了%%"&clng(odds)&"的银两<bgsound src='../mid/xiaotou.wav' loop=1>"
end if
end if
end function
%>
