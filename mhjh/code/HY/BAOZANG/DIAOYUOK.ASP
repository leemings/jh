<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
username=session("yx8_mhjh_username")
sja=session("wabao")
if username="" then Response.Redirect "../../error.asp?id=016"
if  sja="" then Response.Redirect "../../error.asp?id=526"
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('你不能进入,请先进入聊天室再来。谢谢合作'); window.close();</script>"
Response.End 
end if
randomize timer
r=int(rnd*20)+1
  if sja<>"yes" then
     Response.Redirect "diao.asp"
     response.end                                                           
  end if
session("wabao")=""
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
select case r
case 1
	mess="恭喜"&username&"挖到一把飞凤剑，可作为暗器使用，杀伤体力800000、内力200000"
sql="select * from 物品 where 名称='飞凤剑' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,攻击,防御,数量) values ('飞凤剑','" & username & "','暗器',800000,200000,1)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+1 where 名称='飞凤剑' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
case 2
	mess="恭喜"&username&"挖到一包金器，到集市卖得十万银子"
    conn.execute("update 用户 set 银两=银两+100000 where 姓名='" & username & "'")

case 3
	mess="恭喜"&username&"挖到到一把金钥匙，变卖得到银两500000"
   conn.execute("update 用户 set 银两=银两+500000 where 姓名='" & username & "'")

case 4
	mess="恭喜"&username&"挖到到一盒死人骨头，气死了，道德下降3000"
    conn.execute("update 用户 set 道德=道德-3000 where 姓名='" & username & "'")

case 5
	mess="倒霉的"&username&"宝藏没找到，而且不小心踩到陷阱,内力减少20000,但因祸得福，意外得到精力10000！"
   conn.execute("update 用户 set 内力=内力-20000,精力=精力+10000 where 姓名='" & username & "'")

case 6
   mess="强盗又来抢劫,"&username&"反抗遭到毒打,体力下降500"
   conn.execute("update 用户 set 体力=体力-500 where 姓名='" & username & "'")

case 7
     mess="恭喜"&username&"，运气真是好的不得了呀！挖到了一大堆黄金啊,足足值300000两银子！！转换为防御5000"
     conn.execute("update 用户 set 防御=防御+5000 where 姓名='" & username & "'")

case 8
        mess=""&username&"运气真是好，居然挖出了江湖至宝寻梦天使！精力加10000！"
        conn.execute("update 用户 set 精力=精力+10000 where 姓名='" & username & "'")
case 9
        mess=""&username&"做梦都想成为高手，怎奈心急吃不成热豆腐，苦苦修炼独孤九剑破气式，结果造成自残，攻击减少2000点！"
        conn.execute("update 用户 set 攻击=攻击-1000 where 姓名='" & username & "'")
case 10
        mess=""&username&"意外在一个洞穴中挖到一本神秘的古代秘籍，依照上面的口诀练了几天！居然把自己的防御给提高了20000点！"
        conn.execute("update 用户 set 防御=防御+10000 where 姓名='" & username & "'")
case 11
        mess=""&username&"运气真是好，居然挖出了江湖至宝寻梦天使！精力加10000！"
        conn.execute("update 用户 set 精力=精力+10000 where 姓名='" & username & "'")
case 12
	mess="恭喜"&username&"，有够幸运的，从石头缝隙里拣到一块宝石，找人做成了一枚法力无边的戒指。"
sql="select * from 物品 where 名称='魔戒' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,攻击,防御,特效,价格,数量) values ('魔戒','" & username & "','饰品',200000,20000,'抗麻痹',400000,1)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+1 where 名称='魔戒' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
case 13
        mess=""&username&"很是不幸！挖掘时居然被官府当作通缉犯给抓进大牢，官府黑暗啊，硬是说"&username&"杀了人，惨，惨！"
        conn.execute("update 用户 set 杀人=杀人+1,总杀人=总杀人+1 where 姓名='" & username & "'")
case 14
	mess="恭喜"&username&"，学有所成，从强盗手里夺下一把绝世好剑，剑尚未出鞘，已经让人感到寒气逼迫人，赶紧装备上啊，能抵抗疯狂的。"
sql="select * from 物品 where 名称='魔剑' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,攻击,防御,特效,价格,数量) values ('魔剑','" & username & "','武器',100000,20000,'抗疯狂',500000,1)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+1 where 名称='魔剑' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
case 15
        mess=""&username&"在一个小山沟里抠了三天三夜，什么也没弄到，却捡到一个别人丢下的精力袋，自己使用后觉得神气十足！"
        conn.execute("update 用户 set 精力=精力+20000 where 姓名='" & username & "'")
case 16
	mess="恭喜"&username&"挖到一个精致的小檀香木合，小心的打开一看，里面藏着一包神秘药品，"&username&"也不管什么，拿来放在自己的腰包里。"
sql="select * from 物品 where 名称='黄花草' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,体力,内力,价格,数量) values ('黄花草','" & username & "','药品',0,500000,1000000,3)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+3 where 名称='黄花草' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
case 17
	mess="恭喜"&username&"挖到一个远古医圣留下的草药包,实在是美极了,好东西啊,能解麻痹呢。"
sql="select * from 物品 where 名称='麦门冬' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,体力,内力,价格,特效,数量) values ('麦门冬','" & username & "','药品',0,50,400,解麻痹,20)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+20 where 名称='麦门冬' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
case 18
	mess="怪事"&username&"居然挖到一个玻璃容器,里面有两个东西还在来回游动呢,什么东西呢?算了,装起来再说吧。"
sql="select * from 物品 where 名称='小鲤鱼' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,体力,内力,价格,数量) values ('小鲤鱼','" & username & "','药品',0,300000,200000,2)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+2 where 名称='小鲤鱼' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
case 19
        mess=""&username&"在一个小山沟里抠了三天三夜，什么也没弄到，累得吐了血,弄的一点力气都没有了。哎,人为财死,鸟为食亡,干什么都不能太贪婪啊！"
        conn.execute("update 用户 set 内力=内力-500000 where 姓名='" & username & "'")
case 20
	mess="神奇,古怪"&username&"居然挖到一条大草鱼,不知道谁这么倒霉给弄丢到这里了,可惜的是,鱼已经死了,没什么用,只好拿去当药品了,不过,卖还魂丹那家伙很小气,只给换了两个药丸。"
sql="select * from 物品 where 名称='还魂丹' and 所有者='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof and rs.bof then
			sql="insert into 物品(名称,所有者,属性,体力,内力,价格,数量) values ('还魂丹','" & username & "','药品',500000,0,500000,2)"
			conn.execute(sql)
                        else 
				sql="update 物品 set 数量=数量+2 where 名称='还魂丹' and 所有者='" & username & "'"
				conn.execute(sql)
		        end if
                        rs.close
                        set rs=nothing
end select
conn.close
set conn=nothing

dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="大家" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>【挖掘宝藏】</font><font color=blue>"&username&"</font><font color=red>在后山挖宝藏："&mess&"</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr

%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../dg/setup.css" rel=stylesheet type="text/css">
<title>挖宝贝OK</title></head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#000000">
<div align="center"> <br>
<br>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr>
<td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr>
<td valign="top">
<div align="center">
<p><%=mess%></p>
</div>
</table>
<div align="center"><br>
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</div>
</td>
</tr>
</table>
</div>
</body>
</html>


