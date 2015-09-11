<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'江湖神兽系统攻击文件，只需要指向被攻击者，我招式名字
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0

for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
onlinekill=onlinenow/40+1
if onlinenow<0 and chatinfo(0)<>"恩怨情仇" then
	Response.Write "<script language=JavaScript>{alert('江湖在线人数低于0人不能动武！想打去恩怨情仇.那里没有限制');}</script>"
	Response.End
end if
useronlinename=Application("aqjh_useronlinename"&nowinroom)

next
%>
<%'发招(比拼内力)
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")


if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以发招！');}</script>"
	Response.End
end if
f=Minute(time())
if  f>60 and chatinfo(0)<>"恩怨情仇" then
Response.Write "<script language=JavaScript>{alert('神兽动武时间为每小时的前30分钟！');window.close();}</script>"
	Response.End 
end if
call dbdz(1,"发招")
if aqjh_jhdj<30 then
	Response.Write "<script language=JavaScript>{alert('提示：发招需要30级以上才可以操作！');}</script>"
	Response.End
end if
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
'对暂离开、点哑穴判断
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","＆")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【神兽攻击】<font color=" & saycolor & ">"+attack(mid(says,i+1),towho)+"</font>"
call chatsay("发招",towhoway,towho,saycolor,addwordcolor,addsays,says)

'发招(比拼内力)
function attack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,等级,保护,grade,宝物,死亡时间,身份,保留1,lasttime,通缉,属性,姓名 from 用户 where 姓名='" & to1 &"'",conn,2,2
tomp=rs("门派")
tobliu=rs("保留1")
tosf=rs("身份")
npc=rs("属性")
if to1="大家"  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：神兽大叫,想叫我累死啊??！');}</script>"
	Response.End
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他是出家人不可以操作！');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<10 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：他刚刚复活还不足10分钟，还是先放过他吧！');}</script>"
	Response.End
end if
if rs("grade")>=6  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！！');}</script>"
	Response.End
end if
if rs("等级")<=30 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对30级以下的江湖新手操作！');}</script>"
	Response.End
end if
if rs("保护")="true"  and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：对方正在练功保护请不要偷袭!');}</script>"
	Response.End
end if

sj=DateDiff("n",rs("lasttime"),now())
jfj=rs("lasttime")
if sj<1 and rs("通缉")=false then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	attack="<font color=red><b>##对%%动武失败，原因：##刚刚登录聊天室还不到1分钟，不要偷袭。</b></font>"
	exit function
end if
rs.close
rs.open "select 等级,门派,身份,保护,防御,攻击,防御加,攻击加,武功,内力,门派,保留1,杀人数,暴豆时间,grade,死亡时间,z1,z2,z3,z4,z5,z6,lasttime,转生,道德,召唤兽1 from 用户 where 姓名='" & aqjh_name &"'",conn
zaohuan1=rs("召唤兽1")
if InStr(LCase(useronlinename)," " & LCase(zaohuan1) & "")=0 then 
Response.Write "<script language=javascript>alert('神兽在睡觉呢,请先召唤神兽！');</script>"
Response.End 
end if
if to1=zaohuan1  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：神兽大叫,想叫我自杀??！');}</script>"
	Response.End
end if
if to1=Application("figo") then 

rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	attack="<font color=red><b>##支使神兽对%%动武失败，原因："&Application("figo")&"不适合神兽攻击</b></font>"
	exit function
end if

if to1=aqjh_name  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：神兽大叫,我靠,i服了you??！');}</script>"
	Response.End
end if

if rs("等级")<=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你还是个新手，等练到30级的时候再打架吧！');}</script>"
	Response.End
end if
sj=datediff("n",rs("lasttime"),now())
if sj<1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	attack="<font color=red><b>##对%%动武失败，原因：##刚刚登录聊天室还不到1分钟，先准备一下吧。</b></font>"
	exit function
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<10 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你刚刚复活还不到10分钟，还是先练练吧！');}</script>"
	Response.End
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是出家人不可以操作！');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<9 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府中人请不要发招攻击！！！');}</script>"
	Response.End
end if

if rs("保护")="true" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你正在练功保护请不要发招!');}</script>"
	Response.End
end if
if rs("杀人数")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
	conn.execute "update 用户 set 保护='false' where 姓名='" & aqjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('提示：你作恶太多，超过江湖杀人上限"& Application("aqjh_killman") &"不能再发招了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

sj=DateDiff("n",rs("暴豆时间"),now())
xinxi=""
baodoudj=1
if sj<=60 then
	xinxi="<font color=red><b>威力爆满</b></font>"
	baodoudj=3
end if
yxjfyfrom=(rs("防御")+rs("防御加"))*rs("转生")
yxjgjfrom=rs("攻击")+rs("攻击加")
zhuan=rs("转生")
daode=rs("道德")
if daode>50000 then 
yxjwgfrom=rs("武功")*(1+zhuan/50)
end if
if daode<=-5000 then 
yxjwgfrom=rs("武功")*(1-zhuan/50)
end if
if daode>-5000 and daode<=50000 then 
yxjwgfrom=rs("武功")
end if
bliu=rs("保留1")
nlfrom=rs("内力")
menpai=rs("门派")
mysf=rs("身份")
if bliu<>"保留" and menpai<>"出家" and menpai<>"游侠" and menpai<>"天网" and menpai<>"官府" and (mysf="掌门" or mysf="护法" or mysf="长老" or mysf="堂主") then
	zbyhdata=split(bliu,"|")
	zbsj=ubound(zbyhdata)
	if zbsj=4 then
		zbsj=DateDiff("n",zbyhdata(1),now())
		if zbsj<60 then
			xinxi=xinxi&"<font color=red><b>，##振臂一呼有效，在线下忍为其提高功力。</b></font>"
			yxjfyfrom=yxjfyfrom+clng(zbyhdata(3))
			yxjgjfrom=yxjgjfrom+clng(zbyhdata(2))
			yxjwgfrom=yxjgjfrom+clng(zbyhdata(4))
		end if
	end if
	erase zbyhdata
end if
'取出自己装备武器的攻击及特性(并且所有武器耐久-1

rs.close
rs.open "select * from 用户 where 姓名='" & zaohuan1 &"'",conn
sstl=rs("体力")
'开始武功
'rs.open "select id,c,d,e,f,g FROM y WHERE a='" & trim(fn1) & "' and b='" & menpai & "'",conn
'if rs.eof or rs.bof then
	'rs.close
	'set rs=nothing	
	'conn.close
	'set conn=nothing
	'Response.Write "<script language=JavaScript>{alert('提示：你的门派:"& menpai &" 并没有这样的武功["& fn1 &"]呀！');}</script>"
	'Response.End
'end if
wgid=3
wgwg=10000
wgnl=10000
wgdj=20
wgsm="'"&zaohuan1&"'使用神兽法术攻击%%,%%被吓的一生冷汗"
'if rs("g")="随机" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*129)+1 &".gif width=80 height=80  >"
'else
	'gif="<img src=../jhmp/gif/"& rs("g")&" width=80 height=80>"
'end if

if rs("内力")<10000 or rs("武功")<10000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：神兽内力不够或武功不足!');}</script>"
	Response.End
end if
rs.close
'对方的判断
rs.open "select 姓名,攻击,攻击加,门派,身份,防御,防御加,武功,会员等级,z1,z2,z3,z4,z5,z6,属性,道德,转生 from 用户 where 姓名='" & to1 &"'",conn
to1=rs("姓名")
daode=rs("道德")
zhuan=rs("转生")
yxjgjto=rs("攻击")+rs("攻击加")
yxjfyto=rs("防御")+rs("防御加")
if daode>50000 then 
yxjwgto=rs("武功")*(1+zhuan/50)
end if
if daode<=-5000 then 
yxjwgto=rs("武功")*(1-zhuan/50)
end if
if daode>-5000 and daode<=50000 then 
yxjwgto=rs("武功")
end if
jhhy=rs("会员等级")
zbyh=""
if tobliu<>"保留" and tomp<>"出家" and tomp<>"游侠" and tomp<>"天网" and (tosf="掌门" or tosf="护法" or tosf="长老" or tosf="堂主") then
	zbyhdata=split(tobliu,"|")
	zbsj=ubound(zbyhdata)
	if zbsj=4 then
		zbsj=DateDiff("n",zbyhdata(1),now())
		if zbsj<60 then
			zbyh="<font color=red><b>，%%振臂一呼有效，在线弟子为其提高功力。</b></font>"
			yxjfyto=yxjfyto+clng(zbyhdata(3))
			yxjgjto=yxjgjto+clng(zbyhdata(2))
			yxjwgto=yxjgjto+clng(zbyhdata(4))
		end if
	end if
	erase zbyhdata
end if
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
wpsmto=zbyh
'取出对方装备武器的攻击及特性(并且所有武器耐久-1)
if not(isnull(rs("z1"))) and instr(rs("z1"),"|")<>0 then
	data=split(rs("z1"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z1=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z1=null
			wpsmto="%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z2"))) and instr(rs("z2"),"|")<>0 then
	data=split(rs("z2"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z2=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z2=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z3"))) and instr(rs("z3"),"|")<>0 then
	data=split(rs("z3"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z3=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z3=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z4"))) and instr(rs("z4"),"|")<>0 then
	data=split(rs("z4"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z4=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z4=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z5"))) and instr(rs("z5"),"|")<>0 then
	data=split(rs("z5"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		yxjtotx=data(4)
		if (clng(data(1))-1)>0 then
			z5=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z5=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(rs("z6"))) and instr(rs("z6"),"|")<>0 then
	data=split(rs("z6"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z6=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z6=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
conn.execute "update 用户 set z1='"&z1&"',z2='"&z2&"',z3='"&z3&"',z4='"&z4&"',z5='"&z5&"',z6='"&z6&"' where  姓名='"&to1&"'"
rs.close
randomize timer
ki=int(rnd*6)+4
killer=int(((yxjgjfrom+yxjwgfrom+yxjfyfrom)-(yxjfyto+yxjgjto+yxjwgto))/ki)*baodoudj+wgwg+wgnl+wgdj*100
'if killer<=100 then
'	randomize timer
'	killer=int(rnd*99)+1
'end if
randomize timer
'对武功次数计数

randomize timer
ki=int(rnd*500)
if yxjtx<>"无" and yxjtx<>"" then
	if yxjtx<>yxjtotx then
		conn.execute "update 用户 set "&yxjtx&"="&yxjtx&"+"&ki&" where 姓名='"&zaohuan1&"'"
		conn.execute "update 用户 set  "&yxjtx&"="&yxjtx&"-"&ki&" where 姓名='" & to1 &"'"
		gjtx=gjtx&"武器克制对方，吸取%%[<font color=blue><b>"&yxjtx&"</b></font>]共计:+<font color=red>"&ki&"</font>点!"
	else
		conn.execute "update 用户 set "&yxjtx&"="&yxjtx&"-"&ki&" where 姓名='"&zaohuan1&"'"
		conn.execute "update 用户 set  "&yxjtx&"="&yxjtx&"+"&ki&" where 姓名='" & to1 &"'"
		gjtx=gjtx&"武器被对方所克制，'"&zaohuan1&"'被吸取[<font color=blue><b>"&yxjtx&"</b></font>]共计:-<font color=red>"&ki&"</font>点"
	end if
end if
gjtx=gjtx&wpsm&wpsmto
if killer>0 then
	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& int(killer/10) & " where 姓名='"&zaohuan1&"'"
	conn.execute "update 用户 set 体力=体力-"  & killer & " where 姓名='" & to1 &"'"
	djinfo="'"&zaohuan1&"'武功高强"
	mekill=abs(int(killer/10))
else
	if abs(killer)>5000 then killer=5000
	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& abs(killer) & " where 姓名='"&zaohuan1&"'"
	mekill=abs(int(killer/10))
	randomize timer
	killer=int(rnd*400)+1
'	killer=0
	djinfo="'"&zaohuan1&"'武功技不如人"
end if
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*13)+1
mywav="wav"&wav&".wav"
if sstl<-100 then
	attack=xinxi & "'"&zaohuan1&"'<bgsound src=wav/gj/"&mywav&" loop=1>用神兽招式修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点,<br>由于%%的内力深厚，<font color=red>"&zaohuan1&"</font>只感觉眼前一黑，双腿一蹬，灵魂飞往封神台报到去了……"&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call boot(zaohuan1,"发招，操作者："&zaohuan1&",["&menpai&"]唉，技不如人")
	exit function
end if
'--------------------------------------------
randomize()
	myxy = Int(Rnd*1000)	'此处的50000即为设置随机得到装备的概率，数越大，机率越小
	ddwp=""
	tp=""
	lb1=""
	zdsj=""
	otherwp="骨头|太阳保护|僵尸牙|硝酸|硫磺|僵尸血|绿宝石|蓝宝石|木炭|水晶石|红宝石|大草鱼"
	select case myxy
		case 10,22,323,350,411,534,978	'此处为得到极品装备
			dim tmprs
			tmprs=conn.execute("Select count(*) As 极品装备 from b where lb='极品'")
			jpzs=tmprs("极品装备")-1
			set tmprs=nothing
			if isnull(jpzs) then jpzs=0
			sql="select * from b where lb='极品'"
			set rs1=conn.execute(sql)
			randomize()
			xy=int(rnd*jpzs)
			i=0
			mywpsj=""
			do while i<=xy and not(rs1.eof or rs1.bof)
				if i=xy then
					ddwp=rs1("a")
					tp=rs1("i")
					lx=rs1("b")
					select case lx
						case "左手","右手","头盔","盔甲","双脚","装饰"
							lb1="w3"
							zdsj=rs("w3")
						case "药品"
							lb1="w1"
							zdsj=rs("w1")
						case "卡片"
							lb1="w5"
							zdsj=rs("w5")
						case "毒药"
							lb1="w2"
							zdsj=rs("w2")
						case "暗器"
							lb1="w4"
							zdsj=rs("w4")
						case "鲜花"
							lb1="w7"
							zdsj=rs("w7")
						case "配药"
							lb1="w8"
							zdsj=rs("w8")
						case else
							lb1="w6"
							zdsj=rs("w6")
					end select
				end if
				i=i+1
				rs1.movenext
			loop
			rs1.close
			set rs1=nothing
		case else	'此为得到上面的otherwp中定义的物品
			js=split(otherwp,"|")
			n=ubound(js)
			randomize()
			xy=int(rnd*n)
			ddwp=js(xy)
			erase js
			lb1="w6"
			zdsj=rs("w6")
			
	end select
'------------------------------------------



'----------------------------------------------
rs.close
rs.open "select 体力,宝物,门派,通缉,等级,属性 from 用户 where 姓名='" & to1 &"'",conn,2,2

randomize timer
npcr=int(rnd*8)+1
npcr1=int(rnd*41)+1
select case npcr
case 1
conn.execute "update 用户 set 体力=体力-"& mekill &" where 姓名='" &  zaohuan1 &"'"
npc="<br><font color=red>【NPC还击】</font>%%用[NPC家族]招式{绝对诱惑}攻击"&zaohuan1&",杀伤"&zaohuan1&"体力:<font color=red>-"& mekill *2&"</font>点，"&zaohuan1&"气的直发晕……"
case 2
conn.execute "update 用户 set 体力=体力+"&npcr1&"*3583,内力=内力+"&npcr1&"*3228 where 姓名='" & to1 &"'"
npc="<br><font color=red>【补充体力】</font>%%感觉体力不支，急忙使用了<font color=red>[NPC专用特效补药]100个</font>,体力马上暴涨了<font color=red>"&npcr1*35833&"</font>点,内力暴涨<font color=red>"&npcr1*32282&"</font>点……"
case 3
conn.execute "update 用户 set 体力=体力-"& mekill &"*3 where 姓名='" &  zaohuan1&"'"

npc="<br><font color=red>【NPC必杀技】</font>%%用[NPC家族]招式{一掌四式}攻击"&zaohuan1&",杀伤"&zaohuan1&"体力:<font color=red>-"& mekill*6 &"</font>点，打的"&zaohuan1&"晕头转向……"
case 4
conn.execute "update 用户 set 体力=体力-"& mekill &"*2 where 姓名='" &  zaohuan1e &"'"

npc="<br><font color=red>【NPC绝招】</font>%%用[NPC家族]招式{葵花宝典}攻击"&zaohuan1&",杀伤"&zaohuan1&"体力:<font color=red>-"& mekill*4 &"</font>点，打的"&zaohuan1&"直转……"

case 5

npc="<br><font size=4 color=red>%%觉得一阵头晕，真想休息</font>"

case 6
conn.execute "update 用户 set 体力=体力-"& mekill &"*3 where 姓名='" &  zaohuan1 &"'"

npc="<br><font color=red>【NPC必杀技】</font>%%用[NPC家族]招式{一掌四式}攻击"&zaohuan1&",杀伤"&zaohuan1&"体力:<font color=red>-"& mekill*6 &"</font>点，打的"&zaohuan1&"晕头转向……"
case 7
conn.execute "update 用户 set 体力=体力+"&npcr1&"*3583,内力=内力+"&npcr1&"*3228 where 姓名='" & to1 &"'"
npc="<br><font color=red>【补充体力】</font>%%感觉体力不支，急忙使用了<font color=red>[NPC专用特效补药]100个</font>,体力马上暴涨了<font color=red>"&npcr1*35833&"</font>点,内力暴涨<font color=red>"&npcr1*32282&"</font>点……"

case 8

npc="<br><font size=4 color=red>%%觉得好困，真想休息</font>"

end select



'npc="<br><FONT color=green>【npc攻击】</FONT>npc %% 被打了一下,惊醒过来,急忙攻击"&zaohuan1&","&zaohuan1&"被野兽咬了一下,体力减少"& mekill &""
if  rs("属性")<>"npc" then
    aa=xinxi &"'"&zaohuan1&"'<bgsound src=wav/gj/"&mywav&" loop=1>攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点."&gjtx  
	else 
	    aa=xinxi &"'"&zaohuan1&"'<bgsound src=wav/gj/"&mywav&" loop=1>攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点."&gjtx&npc

end if 
if rs("体力")>-100 then
    attack=aa
	rs.close 
	set rs=nothing 
	conn.close 
	set conn=nothing 
	exit function 
	else 
end if 
if rs("属性")<>"npc" then
	conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & aqjh_name &"'" 
	else
		conn.execute "update 用户 set 总杀人=总杀人+1 where 姓名='" & aqjh_name &"'" 
	
		end  if

	if rs("宝物")=Application("aqjh_baowuname") then 
		conn.execute "update 用户 set 宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where 姓名='" & aqjh_name &"'" 
		conn.execute "update 用户 set 宝物修练=0,宝物='无',保护='闭关',内力=100,体力=2000 where 姓名='" & to1 &"'" 
		attack="'"&zaohuan1&"'把%%的<img src=img/z1.gif>宝物:"& Application("aqjh_baowuname") &"抢走，交给主人##因得到此宝。江湖宝物需要进行修练才可以得到更多的东西！"&gjtx
		rs.close 
		set rs=nothing	 
		conn.close 
		set conn=nothing 
		exit function 
	end if 
wpname=""&ddwp&""
'randomize timer
	ii=int(rnd()*4)+1
iii=int(rnd()*20)+1
wusl=""&ii&""

		Application("aqjh_diuqi1")=lb1&"|"&wpname&"|"&wusl&"|"&now()&"|"&npc
	Application.UnLock
'------------------------------------------------------------
randomize
s=int(rnd*500000)+1
sjy=int(rnd*200)+1
snjy=int(rnd*200)+1
yn=0
  jstl=int(rnd*200)+1
  jstl2=int(rnd*200)+1
  jstl3=int(rnd*200)+1
  jstl4=int(rnd*200)+1
  jstl5=int(rnd*200)+1
  jstl6=int(rnd*200)+1
  jstl7=int(rnd*200)+1
q4=int(rnd*8898)+1000
q1=int(rnd*8898)+100
q2=int(rnd*9898)+10
q3=int(rnd*9898)+1
	Application.Lock
Application("aqjh_zlp1")=q4
Application("aqjh_zlp2")=q1
Application("aqjh_zlp3")=q2

  Application("aqjh_ww1")=jstl
  Application("aqjh_ww2")=jstl2
  Application("aqjh_ww3")=jstl3
  Application("aqjh_ww4")=jstl4
  Application("aqjh_ww5")=jstl5
  Application("aqjh_ww6")=jstl6
  Application("aqjh_ww7")=jstl7

	Application("aqjh_by1")=iii+1
Application("aqjh_yb8")=sjy+100
	Application("aqjh_yb9")=s+100
Application("aqjh_yb7")=srjy+100
	Application.UnLock
'--------------------------------------------------------------

dlyp="    <font color=red><input type=button value='药品???个' onClick=javascript:disabled=1;window.open('by1.asp?tl1="&Application("aqjh_by1")&"','d') name=tongyi"&regjm+1&"></font>"
dlyl="     <input type=button value='银两 "&Application("aqjh_yb9")&"' onClick=javascript:disabled=1;window.open('yb9.asp?tl9="&Application("aqjh_yb9")&"','d') name=tongyi"&regjm&">"
dlwp="     <input type=button value='"&wpname&" "&wusl&"个' onClick=javascript:disabled=1;window.open('dq1.asp?tl9="&Application("aqjh_yb9")&"','d') name=tongyi"&regjm&">"
dlwq="     <input type=button value='武器' onClick=javascript:disabled=1;window.open('ww1.asp?tl="&Application("aqjh_ww1")&"','d') name=tongyi"&regjm&">"
dlaq="<input type=button value='暗器' onClick=javascript:disabled=1;window.open('ww3.asp?tl="&Application("aqjh_ww3")&"','d') name=tongyi"&regjm&">"
dlkp="<input type=button value='卡片' onClick=javascript:disabled=1;window.open('ww4.asp?tl="&Application("aqjh_ww4")&"','d') name=tongyi"&regjm&">"
dlxh="<input type=button value='鲜花' onClick=javascript:disabled=1;window.open('ww5.asp?tl="&Application("aqjh_ww5")&"','d') name=tongyi"&regjm&">"
dldgxh="<input type=button value='动感鲜花' onClick=javascript:disabled=1;window.open('ww6.asp?tl="&Application("aqjh_ww6")&"','d') name=tongyi"&regjm&">"
dldy="<input type=button value='毒药' onClick=javascript:disabled=1;window.open('ww2.asp?tl="&Application("aqjh_ww2")&"','d') name=tongyi"&regjm&">"
dljy="<input type=button value='经验 "&Application("aqjh_yb8")&"' onClick=javascript:disabled=1;window.open('yb8.asp?tl8="&Application("aqjh_yb8")&"','d') name=tongyi"&regjm&">"
rdljy="<input type=button value='存点 "&Application("aqjh_yb7")&"' onClick=javascript:disabled=1;window.open('yb7.asp?tl7="&Application("aqjh_yb7")&"','d') name=tongyi"&regjm&">"


randomize
dlxz=int(rnd*500000)+1
if dlxz>499900 or dlxz<=90 then 
dlzs=dlyp&dlyl&dlwp&dlaq&dlkp&dlxh&dldgxh&dldy&dljy
else if dlxz<=490900 and dlxz>=300000 then
dlzs=dlyp&dldy&dlwp&dlaq

else if dlxz<300000 and dlxz>=186000 then

dlzs=dlyp&dlaq&dldy&dldgxh

else dlzs=dlaq&dldgxh&dlyl&dldy

end if
end if
end if
if  rs("属性")<>"npc" then
  ' bb=xinxi & ""&zaohuan1&"<bgsound src=wav/daipu.wav loop=1>攻击<img src=sjfunc/003.gif>%%,%%慢慢倒下..江湖从此又少了一位大虾!"&gjtx&"%%泡点掉了一地，抢啊！~"&rdljy
	else 
	    bb=xinxi & ""&zaohuan1&"<bgsound src=wav/daipu.wav loop=1>攻击<img src=sjfunc/003.gif>%%,%%慢慢倒下..江湖从此又少了一位大虾!"&gjtx&"<br><FONT color=red>【暴物品】</FONT>%%化做一丝清烟,留下物品一堆在路边,谁检到是谁的……"&dlzs
	    
end if 


if rs("通缉")=False then 
		conn.execute "update 用户 set 状态='死',死亡时间=now(),事件原因='"&aqjh_name&"|发招：神兽攻击' where 姓名='" & to1 &"'"
		attack=bb
end if 
if rs("通缉")=True then 
 conn.execute "update 用户 set 通缉=False,状态='死',总杀人=总杀人-1,道德=0,体力=0,内力=0,武功=0,魅力=0,银两=0,存款=int(存款/2),事件原因='"&aqjh_name&"|发招：神兽攻击' where 姓名='" & to1 &"'" 
 conn.execute "update 用户 set 银两=银两+10000000 where 姓名='" & aqjh_name &"'" 
 attack="攻击通缉犯，"&xinxi & ""&zaohuan1&"攻击"&gif&"%%,%%慢慢倒下..为此：主人##得到一千万元官府奖励!"&gjtx 
end if 

call boot(to1,"发招，操作者："&zaohuan1&",["&menpai&"]") 
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & zaohuan1 & "','神兽攻击','人命')" 
rs.close 
set rs=nothing	 
conn.close 
set conn=nothing 
end function 
%>
