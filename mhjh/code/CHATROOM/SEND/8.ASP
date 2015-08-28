<%
function cast(un,st,mg)
tt=now()
f=Minute(time())
rst.Open "select tc.攻击,tc.特效,tu.积分,tu.道德,tu.门派,tu.形态,tu.保护 from 物品 tc inner join 用户 tu on tc.所有者=tu.姓名 where tc.名称='"&mg&"' and tc.所有者='"&un&"' and 数量>0",conn
if rst.EOF or rst.BOF then
cast="<font color=FF0000>【投掷】</font>##，你并没有"&mg&"，所以无法投掷！"
elseif rst("积分")<4000 then
cast="<font color=FF0000>【投掷】</font>##你还是个新人，别惹祸上身！"
elseif rst("保护")=True then
cast="<font color=FF0000>【投掷】</font>##你现在处于保护阶段,无法投掷攻击%%！"
elseif rst("门派")="出家" then
cast="<font color=FF0000>【投掷】</font>##你是已经出家人了，无法投掷攻击%%！"
elseif rst("形态")="鬼魂" then
cast="<font color=FF0000>【投掷】</font>##你是已经鬼魂了，无法投掷攻击%%！赶紧去凤凰蘖盘吧。"
elseif f<30 and chatroomsn<>2 then
cast="<font color=FF0000>【攻击】</font>##用"&mg&"投掷%%失败，原因：这个房间打仗时间是每个小时的后30分钟!要解决个人恩怨请到论剑大厅，那里是杀人的战场！"
else
fmorcal=rst("道德")
athp=rst("攻击")
atespecial=rst("特效")
rst.Close
rst.Open "select 积分,体力,道德,门派,特技,状态,形态,保护 from 用户 where 姓名='"&st&"'",conn
if rst("积分")<4000 then
cast="<font color=FF0000>【投掷】</font>%%还是个新人，你怎么忍心下手呢！"
elseif rst("保护")=True  then
cast="<font color=FF0000>【投掷】</font>%%现在处于练功保护阶段，所以##无法对之使用"&mg
elseif rst("门派")="出家" then
cast="<font color=FF0000>【投掷】</font>%%已经出家了，无法投掷攻击%%！"
elseif st="武术家" then
cast="<font color=FF0000>【投掷】</font>%%是练功用的机器人，你无法用暗器投掷攻击%%！"
elseif rst("形态")="鬼魂" then
cast="<font color=FF0000>【投掷】</font>%%已经是鬼魂了，你无法投掷攻击%%！"
elseif rst("门派")="官府" then
cast="<font color=FF0000>【投掷】</font>%%是官府的老大，##无法投掷攻击！"
elseif Application("yx8_mhjh_chatroomsnkb")=0 and chatroomsn<>2 then
cast="<font color=FF0000>【投掷】</font>##用"&mg&"攻击%%失败，原因：这里不得动武，要解决个人恩怨请到比武大厅!"
else
stespecial=rst("特技")
stmoral=rst("道德")
condition=rst("状态")
sthp=rst("体力")
if isnull(atespecial) then atespecial="无"
conn.execute "update 用户 set 道德=道德-1 where 姓名='"&un&"'"
conn.execute "update 物品 set 数量=数量-1 where 名称='"&mg&"' and 所有者='"&un&"'"
if instr(stespecial,";抗"&atespecial&";")<>0 then atespecial="无"
if condition="死亡" then
atespecial="无"
conn.execute "update 用户 set 道德=道德-100000 where 姓名='"&un&"'"
cast="<font color=FF0000>【投掷】</font>%%业已死亡，##居然还向之投掷"&mg&"，道德因此减100000<bgsound src='../mid/anqi.wav' loop=1>"
elseif sthp>athp then
conn.execute "update 用户 set 体力=体力-"&athp&",防御=防御+1 where 姓名='"&st&"'"
especialtime=athp
cast="<font color=FF0000>【投掷】</font>##向%%投掷了"&mg&"，使之体力减少<bgsound src='../mid/anqi.wav' loop=1>"&athp
else
morcal=clng(fmorcal-tmorcal*0.01)
if morcal>2100000000 then
morcal=2100000000
elseif morcal<-2100000000 then
morcal=-2100000000
end if
atespecial="死亡"
especialtime=18
conn.execute "update 用户 set 道德="&morcal&",杀人=杀人+1,总杀人=总杀人+1 where 姓名='"&un&"'"
conn.execute "insert into 英烈堂(时间,死者,凶手,死因) values(now(),'"&st&"','"&un&"','"&mg&"')"
cast="<font color=FF0000>【投掷】</font>##向%%投掷了"&mg&"，并直接造成%%死亡"
end if
if atespecial<>"无" then cast=cast&newaberrant(st,un,atespecial,especialtime)
Response.Write "<script language=javascript>parent.resfrm.location.href="&chr(34)&"miniweapon.asp"&chr(34)&";</script>"
end if
end if
rst.Close
end function
%>
