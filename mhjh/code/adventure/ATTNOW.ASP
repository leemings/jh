<!--#include file="../config.asp"--><%
username=session("yx8_mhjh_username")
mygrade=Session("yx8_mhjh_usergrade")
biological=session("yx8_mhjh_userfight")
if username="" then
msg="top.location.replace('../error.asp?id=016');"
else
if biological="none" then
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【错误】</FONT>没有战斗对象，命令无效，你该行动了<br>');parent.confrm.document.form1.move.disabled=false;parent.behfrm.location.replace('Action.asp');"
else
attname=Request.QueryString("attack")
if instr(attname,"'")<>0 then 	Response.End
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select tu.体力,tu.内力,ta.基本攻击+tu.攻击 as 攻击,tu.防御,ta.消耗内力 from 攻击 ta inner join 用户 tu on tu.姓名=ta.姓名 where ta.招式='"&attname&"' and ta.姓名='"&username&"'",conn
if rst.EOF or rst.BOF then
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【错误】</FONT>对不起，您无法使用"&attname&"进行攻击！可能是您还没有学会"&attname&msg&"<br>');"
elseif rst("消耗内力")>rst("内力") then
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【错误】</FONT>对不起，您无法使用"&attname&"进行攻击！可能是您内力不足<br>');"
else
fattack=rst("攻击")
fdence=rst("防御")
fhp=rst("体力")
fmp=rst("内力")
ump=rst("消耗内力")
rst.Close
rst.Open "select tf.hp,tb.attack,tb.defence,tb.encourage from fight tf inner join biological tb on tb.biological=tf.biological where tf.biological='"&biological&"' and tf.username='"&username&"'",conn
if rst.EOF or rst.BOF then
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【错误】怪物不存在,请继续冒险</FONT><br>');parent.behfrm.location.href='action.asp';"
else
thp=rst("hp")
tattack=rst("attack")
tdefence=rst("defence")
encourage=rst("encourage")
attack=fattack-tdefence
Maxattack=Application("yx8_mhjh_Maxattack")*mygrade\50
randomize()
if attack<0 then
attack=clng(rnd()*10000)+1
elseif attack>Maxattack then
attack=Maxattack-clng(rnd()*1000)
end if
defence=tattack-fdence
if defence>0 then
defence=Maxattack-clng(rnd()*1000)
elseif defence<=0 then
defence=clng(rnd()*10000)
end if
fhp=fhp-defence
fmp=fmp-ump
thp=thp-attack
if thp>0 and fhp>0 then
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&" where 姓名='"&username&"'"
conn.Execute "update fight set hp="&thp&" where username='"&username&"'"
msg="parent.confrm.document.form1.move.disabled=true;parent.msgfrm.document.writeln('<FONT color=#ff0000>【战斗】</FONT>"&username&"使用"&attname&"攻击"&biological&"，使之生命减速少"&attack&"。<br><FONT color=#ff0000>【战斗】</FONT>"&username&"遭到"&biological&"反击，生命减少"&defence&"。<br>');parent.confrm.document.form1.hp.value='生命："&fhp&"';parent.confrm.document.form1.mp.value='内力："&fmp&"';parent.behfrm.form1.hp.value='生命:"&thp&"';"
elseif fhp>0 and thp<0 then
staple_M=mid(encourage,1,2)
staple_C=mid(encourage,4)
select case staple_M
case "积分"
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&",积分=积分+"&clng(staple_c)&" where 姓名='"&username&"'"
case "攻击"
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&",攻击=攻击+"&clng(staple_c)&" where 姓名='"&username&"'"
case "防御"
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&",防御=防御+"&clng(staple_c)&" where 姓名='"&username&"'"
case "资质"
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&",资质=资质+"&clng(staple_c)&" where 姓名='"&username&"'"
case "风铃"
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&",风铃=风铃+"&clng(staple_c)&" where 姓名='"&username&"'"
case "道德"
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&",道德=道德+"&clng(staple_c)&" where 姓名='"&username&"'"
case "银两"
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&",银两=银两+"&clng(staple_c)&" where 姓名='"&username&"'"
case "物品"
rst.Close
rst.Open "select * from 物品 where 所有者='"&username&"' and 名称='"&staple_C&"'",conn
if rst.EOF or rst.BOF then
rst.close
rst.Open "select * from 商店 where 名称='"&staple_C&"'",conn
conn.Execute "insert into 物品(名称,属性,体力,内力,攻击,防御,特效,数量,价格,所有者,寄售,装备) values('"&rst("名称")&"','"&rst("属性")&"',"&rst("体力")&","&rst("内力")&","&rst("攻击")&","&rst("防御")&",'"&rst("特效")&"',1,"&rst("价格")\2&",'"&username&"',false,false)"
else
conn.Execute "update 物品 set 数量=数量+1 where 所有者='"&username&"' and 名称='"&staple_C&"'"
end if
conn.Execute "update 用户 set 体力="&fhp&",内力="&fmp&" where 姓名='"&username&"'"
end select
conn.Execute "update fight set fnum=fnum+1 where username='"&username&"'"
session("yx8_mhjh_userfight")="none"
msg="parent.msgfrm.document.writeln('<FONT color=#ff0000>【战斗】</FONT>"&username&"使用"&attname&"攻击"&biological&"，使之死亡。得到"&staple_M&staple_C&"<br><FONT color=#ff0000>【战斗】</FONT>"&username&"遭到"&biological&"临死反击，生命减少"&defence&"。<br>');parent.confrm.document.form1.hp.value='生命："&fhp&"';parent.confrm.document.form1.mp.value='内力："&fmp&"';parent.confrm.document.form1.move.disabled=false;parent.behfrm.location.href='action.asp';"
elseif fhp<=0 then
logintime=dateadd("s",1800,now())
logintimetype="#"&month(logintime)&"/"&day(logintime)&"/"&year(logintime)&" "&hour(logintime)&":"&minute(logintime)&":"&second(logintime)&"#"
conn.Execute "update 用户 set 体力=0,状态='死亡',最后登录时间="&logintimetype&" where 姓名='"&username&"'"
session.Abandon
msg="top.location.replace('../error.asp?id=054');"
end if
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
end if
end if
Response.Write "<script language=javascript>"&msg&"</script>"
%>
