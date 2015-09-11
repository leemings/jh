<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'移动(北上)
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green><font color=" & saycolor & ">"+walk(mid(says,i+1))+"</font>"
towhoway=1
towho=aqjh_name
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'移动(北上)
function walk(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")

rs.open "select 门派,内力,武功,银两,轻功,等级 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn
mapai=rs("门派")
if rs("银两")<500000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的钱不够50万所以不能使用行走功能！');window.close();}</script>"
	response.end
end if
if rs("武功")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的武功不够1000所以不能使用行走功能！');window.close();}</script>"
	response.end
end if
if rs("内力")<10000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的内力不够1万所以不能使用行走功能！');window.close();}</script>"
	response.end
end if

if rs("轻功")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的轻功不够1000点了在去存好轻功点数在来吧！');window.close();}</script>"
	response.end
end if

if len(fn1)>3 then
	Response.Write "<script language=JavaScript>{alert('提示：行走方向错误请选择正确路迳 ');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 操作时间 from 用户 where 姓名='"&aqjh_name&"'",conn
sj=DateDiff("n",rs("操作时间"),now())
if sj<1 then
	ss=1-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('这样走你累不累的呀?请等上"&ss&"分再来行走吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

rs.close
conn.execute "update 用户 set 操作时间=now()  where 姓名='"&aqjh_name&"'"
randomize 
r=int(rnd*32)+1
select case r
case 1
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了江湖铁打馆,铁打馆老板</font><font color=red>【陆乘风】说道</font>:<font color=#000080>这里可是爱情江湖最有名的铁打馆什么好武器都有不知道"& mapai &"<font color=red>【##】 </font>想买些什么<font color=red>【##】</font>想了想真不知道买什么好最后有老板<font color=#ff0000>【陆乘风】</font>帮<font color=#ff0000>【##】</font>拿主意,卖了一把飞刀给<font color=#ff0000>【##】</font>算他2万银两而已.因为景气差只好半卖半送 "
	rs.open "SELECT w4 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	duyao=add(rs("w4"),"飞刀",1)
	conn.execute "update 用户 set  轻功=轻功-1000,银两=银两-20000,w4='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 2
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了爱情礼堂,往礼堂四处看了一下发觉完全无人存在只好继续行走</font><font color=red>【" & fn1 & "/直上】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
case 3
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/南下】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
case 4
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了风清扬之家,<font color=red>【##】 </font>高兴死了,想到终于有机会见<font color=red>【风老前辈】</font>一面.<font color=#ff0000>【##】</font>敲了一会儿的门可惜完全无人回应<font color=#ff0000>【##】:说道</font>可能<font color=#ff0000>【风老前辈】</font>有事出门了唉只好等下次在来拜候<font color=#ff0000>【风老前辈】</font>了.<font color=#ff0000>【##】</font>只好继续行走. "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
case 5
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了少林藏经阁,藏经阁大师</font><font color=red>【云尘大师】说道</font>:<font color=#000080>施主看你气色不好是否有内伤.希望施主能让老纳替你治疗<font color=red>【##】 </font>说道:因为这几天刚在华山比武所以身重内伤,所以来此希望能望大师出手相救.<font color=red>【云尘大师】:说道</font>出家人慈悲为怀施主有事当然会出手相救.<font color=#ff0000>【云尘大师】</font>立刻传送了2000的内力给<font color=#ff0000>【##】</font>****<font color=#ff0000>【##】</font>内伤立刻完全好了.<font color=#ff0000>【##】</font>休养了一下立刻行走. "
	conn.execute "update 用户 set  轻功=轻功-1000,内力=内力+2000 where 姓名='"&aqjh_name&"'"
case 6
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/东下】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
case 7
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了宠物商店,宠物店老板娘</font><font color=red>【贵娘】说道</font>:<font color=#000080>这里可是爱情江湖最有名的宠物商店喔"& mapai &"<font color=red>【##】 </font>想了一想还像没有什么好买的只好继续行走.临走前<font color=red>【贵娘】</font>送了三朵玫瑰给<font color=#ff0000>【##】</font>顺便说道:下次还要记得来看我贵娘喔嘻. "
	rs.open "SELECT w7 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	duyao=add(rs("w7"),"玫瑰",3)
	conn.execute "update 用户 set  轻功=轻功-1000,w7='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
case 8
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/直下】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
case 9
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/西南】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
case 10
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了剑术之家,看见众人正在比武练剑.觉得好精彩就往前几步看.却不小心被众人击了一掌.</font><font color=red>【##】</font>:<font color=#000080>立刻内力下降2000<font color=red>【##】 说道:</font>唉这是招谁惹谁了无缘无故被挨了一掌.<font color=red>【##】</font>没心思看下去了只好继续行走. "
	conn.execute "update 用户 set  轻功=轻功-1000,内力=内力-2000 where 姓名='"&aqjh_name&"'"
case 11
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/南下】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
case 12
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/南下】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
case 13
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了风清扬之家,<font color=red>【##】 </font>高兴死了,想到终于有机会见<font color=red>【风老前辈】</font>一面.<font color=#ff0000>【##】</font>敲了一会儿的门<font color=#ff0000>【风老前辈】出来说道:</font>这位大侠不知道有何事<font color=#ff0000>【##】</font>立刻说道:因有幸路过此地想说顺便拜候你老人家.也希望<font color=#ff0000>【风老前辈】</font>能指点晚辈一两招武功<font color=#ff0000>【风老前辈】说道:</font>看你这么有诚意就传你两招武功. <font color=#ff0000>【##】</font>武功立刻增进1000. <font color=#ff0000>【##】</font>学完后顺便跟<font color=#ff0000>【风老前辈】</font>拜别就继续行走."
	conn.execute "update 用户 set 轻功=轻功-1000,武功=武功+1000 where 姓名='"&aqjh_name&"'" 
case 14
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/南下】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 


case 15
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走走到了华山顶峰.<font color=#ff0000>【##】</font>说道:真是奇景想不到华山如此美丽 <font color=#ff0000>【##】</font>.见到一堆人在打猎于是看到地上有两块老虎肉立刻检走.<font color=#ff0000>【##】</font>笑道真是太好了检到2块虎肉."
	rs.open "SELECT w6 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	duyao=add(rs("w6"),"虎肉",2)
	conn.execute "update 用户 set  轻功=轻功-1000,w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 16
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走走到了三国药铺.药铺<img src='../pic/dz45.gif' width='30' height='30'><font color=#ff0000>【赵老板】</font>立刻出来招呼 <font color=#ff0000>【##】</font>.客棺不知道想买点什么这里什么药都有还很便宜喔.<font color=#ff0000>【##】</font>想了一会说道给我来个九花雨露丸.<font color=#ff0000>【赵老板】</font>立刻把九花雨露丸拿出来.客棺200000而已便宜吧<font color=#ff0000>【##】</font>拿了付了钱就立刻在度行走."
	rs.open "SELECT w1 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	duyao=add(rs("w1"),"九花雨露丸",1)
	conn.execute "update 用户 set  轻功=轻功-1000,银两=银两-200000,w1='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close

case 17
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走走到了泉源澡堂.澡堂老板<font color=#ff0000>【胖子王】</font>立刻出来招呼<font color=#ff0000>【##】</font>客棺.客棺进来洗个澡吧但可没有女人陪喔.<font color=#ff0000>【##】</font>想说好吧反正好几天每洗澡了进来冲个凉的.洗好了<font color=#ff0000>【##】</font>付了钱给<font color=#ff0000>【胖子王】</font>之后就继续上路. "
	conn.execute "update 用户 set 轻功=轻功-1000,银两=银两-5000 where 姓名='"&aqjh_name&"'" 

case 18
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/南下】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 

case 19
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走走到了爱情药铺.药铺<img src='../images/picc/man5.gif' width='30' height='30'><font color=#ff0000>【赵老板】</font>立刻出来招呼 <font color=#ff0000>【##】</font>.客棺不知道想买点什么这里什么药都有还很便宜喔.<font color=#ff0000>【##】</font>想了一会说道给我来个大白菜.<font color=#ff0000>【赵老板】</font>立刻把大白菜拿出来.客棺200000而已便宜吧<font color=#ff0000>【##】</font>拿了付了钱就立刻在度行走."
	rs.open "SELECT w1 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	duyao=add(rs("w1"),"大白菜",1)
	conn.execute "update 用户 set  轻功=轻功-1000,银两=银两-200000,w1='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close


case 20
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/北上】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 

case 21
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走走到了混世客栈.客栈老板<font color=#ff0000>【龙老板】</font>立刻出来招呼<font color=#ff0000>【##】</font>客棺.客棺想吃点什么这里的东西吃了可是会增加你的内力喔.<font color=#ff0000>【##】</font>说道:那就来半斤牛肉半斤猪肉五个馒头顺便在来一坛女儿红<font color=#ff0000>【龙老板】</font>立刻把酒菜端出来<font color=#ff0000>【##】</font>吃完了顺便付了5000两给<font color=#ff0000>【龙老板】</font>*******.<font color=#ff0000>【龙老板】</font>说道:下次要记得在来光顾喔<font color=#ff0000>【##】</font>走前发觉内力突然上升2000.真是精神百倍 "
	conn.execute "update 用户 set 轻功=轻功-1000,内力=内力+2000,银两=银两-5000 where 姓名='"&aqjh_name&"'" 

case 22
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/南下】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
	
case 23
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/东】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 	

case 24
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走走到了大街钱庄.钱庄老板<font color=#ff0000>【钱万里】</font>立刻出来招呼<font color=#ff0000>【##】</font>客棺.客棺想来存点钱吗这里利息可高了.<font color=#ff0000>【##】</font>说道:现在爱情这么乱那有钱存<font color=#ff0000>【钱老板】</font>看了<font color=#ff0000>【##】</font>一下觉得他很可怜就从钱庄拿出一个金币给<font color=#ff0000>【##】</font>*******.<font color=#ff0000>【##】</font>高兴死了连忙跟<font color=#ff0000>【钱老板】</font>道谢..谢过完继续上路 "
	conn.execute "update 用户 set 轻功=轻功-1000,金币=金币+1 where 姓名='"&aqjh_name&"'" 
	
case 25
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/北上】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 

case 26
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/北上】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
	
case 27
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走走到了炼丹房.炼丹房老板<font color=#ff0000>【药子仙】</font>立刻出来招呼 <font color=#ff0000>【##】</font>.这位侠客不知道来丹庐房要买什么我这什么好药都有.<font color=#ff0000>【##】</font>想了一会说道给我来个大力丸.<font color=#ff0000>【药子仙】</font>立刻把大力丸拿出来.客棺200000而已便宜吧<font color=#ff0000>【##】</font>拿了付了钱就立刻在度行走."
	rs.open "SELECT w8 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	duyao=add(rs("w8"),"大力丸",1)
	conn.execute "update 用户 set  轻功=轻功-1000,银两=银两-200000,w8='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
			
case 28
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/左西】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 
	
case 28
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 探了探路,却完全没发现路口只好继续行走.<font color=#ff0000>【" & fn1 & "/东南】</font> "
	conn.execute "update 用户 set 轻功=轻功-1000 where 姓名='"&aqjh_name&"'" 

case 30
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了藏经阁,觉得太好了可以来偷学点武功.</font><font color=red>【##】</font>:<font color=#000080>立刻拿起易经筋开始习武<font color=red>【##】 </font>习了一段时间觉得差不多了可以立刻上路.<font color=red>【##】</font>走前发觉武功大涨1000点发觉真是学无止尽. "
	conn.execute "update 用户 set  轻功=轻功-1000,武功=武功+1000 where 姓名='"&aqjh_name&"'"

case 31
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了剑术之家,看见众人正在比武练剑.觉得好精彩就往前几步看.却不小心被众人击了一掌.</font><font color=red>【##】</font>:<font color=#000080>立刻内力下降2000<font color=red>【##】 说道:</font>唉这是招谁惹谁了无缘无故被挨了一掌.<font color=red>【##】</font>没心思看下去了只好继续行走. "
	conn.execute "update 用户 set  轻功=轻功-1000,内力=内力-2000 where 姓名='"&aqjh_name&"'"
				
case 32
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走到了剑术之家,看见众人正在比武练剑.觉得好精彩就往前几步看.却不小心被众人击了一掌.</font><font color=red>【##】</font>:<font color=#000080>立刻内力下降2000<font color=red>【##】 说道:</font>唉这是招谁惹谁了无缘无故被挨了一掌.<font color=red>【##】</font>没心思看下去了只好继续行走. "
	conn.execute "update 用户 set  轻功=轻功-1000,内力=内力-2000 where 姓名='"&aqjh_name&"'"
	
case 33
walk="<font color=red>【行走消息】</font><font color=#000080>位于"& texin &""& mapai &"<font color=red>【##】</font> 往" & fn1 & "<img src='img/ki17.gif' width='20' height='20'> 行走走到了爱情药铺.药铺<img src='../images/picc/man5.gif' width='30' height='30'><font color=#ff0000>【赵老板】</font>立刻出来招呼 <font color=#ff0000>【##】</font>.客棺不知道想买点什么这里什么药都有还很便宜喔.<font color=#ff0000>【##】</font>想了一会说道给我来个大白菜.<font color=#ff0000>【赵老板】</font>立刻把大白菜拿出来.客棺200000而已便宜吧<font color=#ff0000>【##】</font>拿了付了钱就立刻在度行走."
	rs.open "SELECT w1 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
	duyao=add(rs("w1"),"大白菜",1)
	conn.execute "update 用户 set  轻功=轻功-1000,银两=银两-200000,w1='"&duyao&"' where 姓名='"&aqjh_name&"'"
	rs.close
end select
set rs=nothing	
conn.close
set conn=nothing
end function
%>