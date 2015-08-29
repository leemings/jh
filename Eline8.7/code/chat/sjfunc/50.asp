<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<%'在线赌博押大♀wWw.51eline.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不能够赌博！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
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
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【在线赌博】<font color=" & saycolor & ">"+xiazhu(mid(says,i+1))+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'在线赌博押大
function xiazhu(fn1)
randomize ()
m1 = Int(6 * Rnd + 1)
if instr(fn1,"&")=0 or right(fn1,1)="&" then
	Response.Write "<script language=JavaScript>{alert('操作错误，格式如下：[赌局操作&押操作&银两]');}</script>"
	Response.End 
end if
zt=split(fn1,"&")
abcc=trim(zt(2))
if not isnumeric(abcc) then
	Response.Write "<script language=JavaScript>{alert('["&abcc&"]操作错误，数量请使用数字！');}</script>"
	Response.End 
end if
yinjindian=trim(zt(0))
yacz=zt(1)
yinliang=abs(int(zt(2)))
if yinjindian<>"金" and yinjindian<>"银" and yinjindian<>"点" then
	Response.Write "<script language=JavaScript>{alert('选择有误：银、金、点！');}</script>"
	Response.End 
end if
select case yacz
	case "大"
		duboimg="<img src='../jhimg/da.gif'>"
	case "小"
		duboimg="<img src='../jhimg/xiao.gif'>"
	case "单"
		duboimg="<img src='../jhimg/dan.gif'>"
	case "双"
		duboimg="<img src='../jhimg/shuang.gif'>"
case else
	Response.Write "<script language=JavaScript>{alert('押操作为：大、小、单、双！');}</script>"
	Response.End 
end select
	if (yinliang<1000 or yinliang>2000000) and yinjindian="银" then
		Response.Write "<script language=JavaScript>{alert('[银]最少押：1000两，最多200万！');}</script>"
		Response.End 
	end if
	if (yinliang<10 or yinliang>200) and (yinjindian="点" or yinjindian="金") then
		Response.Write "<script language=JavaScript>{alert('["&yinjindian&"]最少押：10，最多200"&yinjindian&"！');}</script>"
		Response.End 
	end if

'开始判断
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
select case yinjindian
'对赌博银的判断
case "银"
	randomize(timer())
	m2 = Int(6 * Rnd + 1)
		rs.open "select 银两 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
		if rs("银两")<yinliang then
			rs.close
			set rs=nothing	
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('你有这么多的钱吗？！');}</script>"
			Response.End 		
		end if
		rs.close
	rs.open "select top 1 a FROM g WHERE c='银' and b='庄家' and f=true",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('现在[银局]没有庄家，不能在线赌博！！');}</script>"
		Response.End 		
	end if
	rs.close
	rs.open "select top 1 * FROM g WHERE c='银' and a='"& sjjh_name &"' and f=true",conn,2,2
	if not(rs.eof) or not(rs.bof) then
		if rs("b")="庄家" then
			temp=sjjh_name&"你现在是[银局]庄家，你要搞什么呀!"
		else
			temp=sjjh_name&"你在[银局]为闲压["&rs("e")&"]"&rs("d")&"等着开局吧！"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 id,a FROM g WHERE c='银' and b='玩家' and f=false",conn,2,2
	tempdu=0
	if rs.eof or rs.bof then
		Response.Write "<script language=JavaScript>{alert('提示：在银局里现在没有空位正在开局!');}</script>"
		tempdu=1
		rs.close
	end if
	if 	tempdu=0 then
	id=rs("id")
	conn.execute "update 用户 set 银两=银两-"&yinliang&" where 姓名='" & sjjh_name &"'"
	conn.execute "update g set a='"&sjjh_name&"',f=true,d="&yinliang&",e='"&yacz&"' where id="&id&" and f=false"
	rs.close
	end if
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and e='大' and f=true")
	dars=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and e='小' and f=true")
	xiaors=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and e='单' and f=true")
	danrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and e='双' and f=true")
	shuangrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='银' and d<>0 and f=true")
	durs=tmprs("数量")

'开局了
if durs>=10 then
	sj=DateDiff("s",Application("sjjh_du"),now())
	if sj<5 then
		s=5-sj
		Response.Write "<script language=JavaScript>{alert('提示：网络在正处理数据，请您等上["&s&"秒钟]再操作！');}</script>"
		Response.End
	end if	
	Application.Lock
	Application("sjjh_du")=now()
	Application.UnLock
	randomize()
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'查找庄家
rs.open "select * FROM g WHERE c='银' and b='庄家' and f=true",conn,2,2
zhuangjia=rs("a")
rs.close

'豹子处理
if m1=m2 and m2=m3 and m3=m1 then
	rs.open "select * FROM g WHERE c='银' and d<>0 and b<>'庄家' and f=true",conn,2,2
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("d")
	rs.movenext
	loop
	rs.close
	'给庄加钱
	conn.execute "update 用户 set 体力=体力+10000,存款=存款+"&yingyin&" where 姓名='"& zhuangjia &"'"
	'清除所有赌博状态
	conn.execute "update g set f=false where f=true and c='银'"
	xiazhu="[银两局]庄家开：<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>计："& sjdubo &"点!庄家开出豹子……通杀！庄家："&zhuangjia&"收入："&yingyin&"两！气势如宏，体力+1万点！"
		set rs=nothing
		conn.close
		set conn=nothing
		exit function
end if

'初始化数据
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'开单双处理
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src='../jhimg/shuang.gif'>"
	rs.open "select * FROM g WHERE c='银' and e='双' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("d")
		shuangname=shuangname&rs("a")&" "
		conn.execute "update 用户 set 存款=存款+"&rs("d")*2&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		conn.execute "update 用户 set 存款=存款-"& rs("d") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	'清除所有赌博状态
	conn.execute "update g set f=false where f=true and c='银' and e='双'"
end if
if sjdubo/2<>int(sjdubo/2) then
	danshuang="<img src='../jhimg/dan.gif'>"
	rs.open "select * FROM g WHERE c='银' and e='单' and f=true",conn
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("d")
		danname=danname&rs("e")&" "
		conn.execute "update 用户 set 存款=存款+"&rs("d")*2&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		conn.execute "update 用户 set 存款=存款-"& rs("d") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='银' and e='单'"
end if

'对开大小处理
if sjdubo<=10 then
	daxiao="<img src='../jhimg/xiao.gif'>"
	rs.open "select * FROM g WHERE c='银' and e='小' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("d")
		xiaoname=xiaoname&rs("a")&" "
		conn.execute "update 用户 set 存款=存款+"&rs("d")*2&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		conn.execute "update 用户 set 存款=存款-"& rs("d") &" where 姓名='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='银' and e='小'"
end if
if sjdubo>10 then
	daxiao="<img src='../jhimg/da.gif'>"
	rs.open "select * FROM g WHERE c='银' and e='大' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("d")
		daname=daname&rs("a")&" "
		conn.execute "update 用户 set 存款=存款+"&rs("d")*2&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		conn.execute "update 用户 set 存款=存款-"& rs("d") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='银' and e='大'"
end if

'对剩下输的用户银两
	rs.open "select * FROM g WHERE c='银' and d<>0 and b<>'庄家' and f=true",conn,2,2
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("d")
	rs.movenext
	loop
	rs.close
	conn.execute "update 用户 set 存款=存款+"&yingyin&" where 姓名='"& zhuangjia &"'"
	conn.execute "update g set f=false where f=true and c='银'"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	duboname=shuangname&danname&xiaoname&daname
	xiazhu="[银两局]庄家开：<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>计：<font color=blue>"& sjdubo &"</font>点!"&danshuang&daxiao&"总下注：<font color=blue>"&zong&"</font>两，庄家：["&zhuangjia&"]收入：<font color=blue>"&yingyin &"</font>两,赔出：<font color=blue>"&pei&"</font>两，合计：<font color=blue>"&yingyin-pei&"</font>两,共有：<font color=red>"&duboname&"</font>玩家幸运押中,所有押中玩家内力、体力各加+500！"
	set rs=nothing
	conn.close
	set conn=nothing
	exit function

end if
xiazhu="[银两局]##从自己的小荷包里拿出:"& yinliang &"两，我买"& duboimg &"！，一定中的！现在下注如下：押大：<font color=blue>"& dars &"</font>个，押小:<font color=blue>"& xiaors &"</font>个,押单：<font color=blue>"&danrs&"</font>个,押双:<font color=blue>"&shuangrs&"</font>个！还差:<font color=blue>"&(10-durs)&"</font>个开局！"
	set rs=nothing	
	conn.close
	set conn=nothing

'对存点赌博判断
case "点"
		randomize(timer())
		m2 = Int(6 * Rnd + 1)
		rs.open "select allvalue,会员等级 FROM 用户 WHERE 姓名='" & sjjh_name &"'",conn
		select case rs("会员等级")
		case 1
			hydian=31250
		case 2
			hydian=90000
		case 3
			hydian=250000
		case 4
			hydian=490000
		case else
			hydian=0
		end select
		if (rs("allvalue")-hydian)<yinliang then
			rs.close
			set rs=nothing	
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('你有这么多的存点吗？！');}</script>"
			Response.End 		
		end if
		if sjjh_jhdj<10 then
			rs.close
			set rs=nothing	
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('提示：江湖赌博存点需要10级以上才可以操作！');}</script>"
			Response.End 		
		end if

		rs.close
	rs.open "select top 1 a FROM g WHERE c='点' and b='庄家' and f=true",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('现在[存点局]没有庄家，不能在线赌博！！');}</script>"
		Response.End 		
	end if
	rs.close
	rs.open "select top 1 * FROM g WHERE c='点' and a='"& sjjh_name &"' and f=true",conn
	if not(rs.eof) or not(rs.bof) then
		if rs("b")="庄家" then
			temp="##你现在是[存点局]庄家，你要搞什么呀!"
		else
			temp="##你在[存点局]为闲压["&rs("e")&"]"&rs("d")&"等着开局吧！"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 id,a FROM g WHERE c='点' and b='玩家' and f=false",conn,2,2
	tempdu=0
	if rs.eof or rs.bof then
		rs.close
		tempdu=1
		Response.Write "<script language=JavaScript>{alert('提示：在点局里现在没有空位正在开局!');}</script>"
	end if
	if tempdu=0 then
	id=rs("id")
	conn.execute "update 用户 set allvalue=allvalue-"&yinliang&" where 姓名='" & sjjh_name &"'"
	conn.execute "update g set a='"&sjjh_name&"',f=true,d="&yinliang&",e='"&yacz&"' where id="&id&" and f=false"
	rs.close
	end if
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and e='大' and f=true")
	dars=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and e='小' and f=true")
	xiaors=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and e='单' and f=true")
	danrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and e='双' and f=true")
	shuangrs=tmprs("数量")
	tmprs=conn.execute("Select count(*) As 数量 from g where c='点' and d<>0 and f=true")
	durs=tmprs("数量")

'开局了
if durs>=10 then
	sj=DateDiff("s",Application("sjjh_du"),now())
	if sj<5 then
		s=5-sj
		Response.Write "<script language=JavaScript>{alert('提示：网络在正处理数据，请您等上["&s&"秒钟]再操作！');}</script>"
		Response.End
	end if	
	Application.Lock
	Application("sjjh_du")=now()
	Application.UnLock
	randomize()
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'查找庄家
rs.open "select * FROM g WHERE c='点' and b='庄家' and f=true",conn,2,2
zhuangjia=rs("a")
rs.close

'豹子处理
if m1=m2 and m2=m3 and m3=m1 then
	rs.open "select * FROM g WHERE c='点' and d<>0 and b<>'庄家'",conn,2,2
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("d")
	rs.movenext
	loop
	rs.close
	'给庄加钱
	conn.execute "update 用户 set allvalue=allvalue+"&yingyin&",体力=体力+10000 where 姓名='"& zhuangjia &"'"
	'清除所有赌博状态
	conn.execute "update g set f=false where f=true and c='点'"
	xiazhu="[存点局]庄家开：<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>计："& sjdubo &"点!庄家开出豹子……通杀！庄家："&zhuangjia&"收入："&yingyin&"存点！鸿运当头体力+1万点！"
		set rs=nothing
		conn.close
		set conn=nothing
		exit function
end if

'初始化数据
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'开单双处理
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src='../jhimg/shuang.gif'>"
	rs.open "select * FROM g WHERE c='点' and e='双' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("d")
		shuangname=shuangname&rs("a")&" "
		conn.execute "update 用户 set allvalue=allvalue+"&rs("d")*2&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		conn.execute "update 用户 set allvalue=allvalue-"& rs("d") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='点' and e='双'"
end if
if sjdubo/2<>int(sjdubo/2) then
	danshuang="<img src='../jhimg/dan.gif'>"
	rs.open "select * FROM g WHERE c='点' and e='单' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("d")
		danname=danname&rs("a")&" "
		conn.execute "update 用户 set allvalue=allvalue+"&rs("d")*2&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		conn.execute "update 用户 set allvalue=allvalue-"& rs("d") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='点' and e='单'"
end if


'对开大小处理
if sjdubo<=10 then
	daxiao="<img src='../jhimg/xiao.gif'>"
	rs.open "select * FROM g WHERE c='点' and e='小' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("d")
		xiaoname=xiaoname&rs("a")&" "
		conn.execute "update 用户 set allvalue=allvalue+"&rs("d")*2&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		conn.execute "update 用户 set allvalue=allvalue-"& rs("d") &" where 姓名='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='点' and e='小'"
end if
if sjdubo>10 then
	daxiao="<img src='../jhimg/da.gif'>"
	rs.open "select * FROM g WHERE c='点' and e='大' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("d")
		daname=daname&rs("a")&" "
		conn.execute "update 用户 set allvalue=allvalue+"&rs("d")*2&",体力=体力+500,内力=内力+500 where 姓名='"& rs("a") &"'"
		conn.execute "update 用户 set allvalue=allvalue-"& rs("d") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "update g set f=false where f=true and c='点' and e='大'"
end if

'对剩下输的用户点处理
	rs.open "select * FROM g WHERE c='点' and d<>0 and b<>'庄家' and f=true",conn,2,2
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("d")
	rs.movenext
	loop
	rs.close
	conn.execute "update 用户 set allvalue=allvalue+"&yingyin&" where 姓名='"& zhuangjia &"'"
	conn.execute "update g set f=false where f=true and c='点'"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	duboname=shuangname&danname&xiaoname&daname
	xiazhu="[存点局]庄家开：<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>计：<font color=blue>"& sjdubo &"</font>点!"&danshuang&daxiao&"总下注：<font color=blue>"&zong&"</font>点，庄家：["&zhuangjia&"]收入：<font color=blue>"&yingyin &"</font>存点,赔出：<font color=blue>"&pei&"</font>存点，合计：<font color=blue>"&yingyin-pei&"</font>存点,共有：<font color=red>"&duboname&"</font>玩家幸运押中！所有押中玩家内力、体力各加+500点！"
	set rs=nothing
	conn.close
	set conn=nothing
	exit function

end if
xiazhu="[存点局]##把自己努力泡的存点拿出:"& yinliang &"点，我押"& duboimg &"！，一定中的！现在下注如下：押大：<font color=blue>"& dars &"</font>个，押小:<font color=blue>"& xiaors &"</font>个,押单：<font color=blue>"&danrs&"</font>个,押双:<font color=blue>"&shuangrs&"</font>个！还差:<font color=blue>"&(10-durs)&"</font>个开局！"

	set rs=nothing	
	conn.close
	set conn=nothing

end select
end function
%>
