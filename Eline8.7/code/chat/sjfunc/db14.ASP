<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'发招(比拼内力)♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_jhdj<jhdj_fz then
	Response.Write "<script language=JavaScript>{alert('提示：发招需要["&jhdj_fz&"]级以上才可以操作！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
call roompd("夺宝动武")
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(towho)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：你要打的人不在聊天室里！');}</script>"
	Response.End
end if
if towho="大家" or towho=sjjh_name or towho=application("sjjh_automanname") then
	Response.Write "<script language=JavaScript>{alert('提示：不可以对自已、大家或机器人使用！');}</script>"
	Response.End
end if
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【夺宝动武】<font color=" & saycolor & ">"+attack(mid(says,i+1),towho)+"</font>"
call chatsay("发招",towhoway,towho,saycolor,addwordcolor,addsays,says)

'发招(比拼内力)
function attack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
tonamedb=to1
rs.open "select id,门派,grade from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：请不要对官府进行攻击！！');}</script>"
	Response.End
end if
tonameid=rs("id")
rs.close
rs.open "select id,门派,保护,防御,攻击,武功,内力,门派,杀人数,暴豆时间,grade,z1,z2,z3,z4,z5,z6 from 用户 where 姓名='" & sjjh_name &"'",conn
if rs("grade")>=6 and rs("grade")<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是官府中人请不要发招攻击！！！');}</script>"
	Response.End
end if
mynameid=rs("id")
sj=DateDiff("n",rs("暴豆时间"),now())
xinxi=""
baodoudj=1
if sj<=60 then
	xinxi="<font color=red><b>威力爆满</b></font>"
	baodoudj=3
end if
yxjfyfrom=rs("防御")
yxjgjfrom=rs("攻击")
yxjwgfrom=rs("武功")
nlfrom=rs("内力")
menpai=rs("门派")
myid=rs("id")
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
myz1=rs("z1")
myz2=rs("z2")
myz3=rs("z3")
myz4=rs("z4")
myz5=rs("z5")
myz6=rs("z6")
rs.close
'武功判断
rs.open "select id,c,d,e,f,g FROM y WHERE a='" & trim(fn1) & "' and b='" & menpai & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的门派并没有这样的武功！！！');}</script>"
	Response.End
end if
wgid=int(rs("id"))
wgwg=int(rs("c"))
wgnl=int(rs("d"))
wgdj=int(sqr(rs("e")/100))+1
wgsm=rs("f")
wgsm=replace(wgsm,"##",sjjh_name)
wgsm=replace(wgsm,"%%",to1)
if rs("g")="随机" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*23)+1 &".gif>"
else
	gif="<img src=../jhmp/gif/"& rs("g")&">"
end if
if rs("c")>nlfrom or rs("d")>yxjwgfrom then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：您的内力或武功不足!');}</script>"
	Response.End
end if
rs.close

'取出自己装备武器的攻击及特性(并且所有武器耐久-1
if not(isnull(myz1)) and instr(myz1,"|")<>0 then
	data=split(myz1,"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z1=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z1=null
			wpsm="##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(myz2)) and instr(myz2,"|")<>0 then
	data=split(myz2,"|")
	if ubound(data)=5 then	
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z2=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z2=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(myz3)) and instr(myz3,"|")<>0 then
	data=split(myz3,"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z3=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z3=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(myz4)) and instr(myz4,"|")<>0 then
	data=split(myz4,"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z4=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z4=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(myz5)) and instr(myz5,"|")<>0 then
	data=split(myz5,"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		yxjtx=data(4)
		gjtx="<font color=red><b>["&data(0)&"]</b></font>"
		if (clng(data(1))-1)>0 then
			z5=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z5=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
if not(isnull(myz6)) and instr(myz6,"|")<>0 then
	data=split(myz6,"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z6=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z6=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]坏掉了"
		end if
	end if
end if
conn.execute "update 用户 set z1='"&z1&"',z2='"&z2&"',z3='"&z3&"',z4='"&z4&"',z5='"&z5&"',z6='"&z6&"' where  姓名='"&sjjh_name&"'"

'对方的判断
rs.open "select id,姓名,攻击,防御,武功,会员等级,z1,z2,z3,z4,z5,z6 from 用户 where 姓名='" & to1 &"'",conn
to1=rs("姓名")
yxjgjto=rs("攻击")
yxjfyto=rs("防御")
yxjwgto=rs("武功")
jhhy=rs("会员等级")
toid=rs("id")
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
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
randomize timer
'对武功次数计数
if menpai<>"游侠" then
	conn.execute "update y set e=e+1 where id="&wgid
end if
randomize timer
ki=int(rnd*500)
if yxjtx<>"无" and yxjtx<>"" then
	if yxjtx<>yxjtotx then
		conn.execute "update 用户 set "&yxjtx&"="&yxjtx&"+"&ki&" where 姓名='" & sjjh_name &"'"
		conn.execute "update 用户 set  "&yxjtx&"="&yxjtx&"-"&ki&" where 姓名='" & to1 &"'"
		gjtx=gjtx&"武器克制对方，吸取对方[<font color=blue><b>"&yxjtx&"</b></font>]共计:+<font color=red>"&ki&"</font>点!"
	else
		conn.execute "update 用户 set "&yxjtx&"="&yxjtx&"-"&ki&" where 姓名='" & sjjh_name &"'"
		conn.execute "update 用户 set  "&yxjtx&"="&yxjtx&"+"&ki&" where 姓名='" & to1 &"'"
		gjtx=gjtx&"武器被对方所克制，被吸取[<font color=blue><b>"&yxjtx&"</b></font>]共计:-<font color=red>"&ki&"</font>点"
	end if
end if
gjtx=gjtx&wpsm&wpsmto
if killer>0 then
	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& int(killer/10) & " where 姓名='" & sjjh_name &"'"
	conn.execute "update 用户 set 体力=体力-"  & killer & " where 姓名='" & to1 &"'"
	djinfo="##武功高强"
	mekill=abs(int(killer/10))
else
	if abs(killer)>5000 then killer=5000
	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& abs(killer) & " where 姓名='" & sjjh_name &"'"
	mekill=abs(int(killer/10))
	randomize timer
	killer=int(rnd*400)+1
'	killer=0
	djinfo="##武功技不如人"
end if
rs.open "select 体力 from 用户 where 姓名='" & sjjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*13)+1
mywav="wav"&wav&".wav"
if rs("体力")<-100 then
	conn.execute "update 用户 set 状态='死',死亡时间=now(),宝物='无',宝物修练=0,事件原因='"&sjjh_name&"|发招："&fn1&"被"&to1&"震死' where 姓名='"& sjjh_name &"'"
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点,<br>由于%%的内力深厚，##只感觉眼前一些，双腿一蹬，灵魂飞往封神台报到去了……"&gjtx
	newdjinfo=replace(djinfo,"##",sjjh_name)
	xx=xinxi & "<font color=blue>"& sjjh_name &"</font><bgsound src=wav/gj/"&mywav&" loop=1>为夺取宝物，用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"<font color=blue>"& to1 &"</font>,[<font color=red><b>"&newdjinfo&"</b></font>],自己被<font color=blue>"& to1 &"</font>的武功震得体力:<font color=red>-"& mekill &"</font>点,<b><font color=red>"& sjjh_name &"只感觉眼前一些，双腿一蹬，灵魂飞往封神台报到去了……</font></b>"
	rs.close
	call boot(sjjh_name,"发招，操作者："&sjjh_name&",["&menpai&"]唉，技不如人"&fn1)
	bbb=zxrsbd(tonamedb,tonameid)
	attack=attack &"<br><br>"&bbb
	xx=xx &"<br><br>"& bbb
	set rs=nothing
	conn.close
	set conn=nothing
	call dbxx(xx)
	towho=tonamedb
	exit function
end if
rs.close
rs.open "select 体力,宝物,门派,等级 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("体力")>-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点."&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & sjjh_name &"'"
if rs("等级")>=10 then
	all=rs("等级")*2
else
	all=2
end if
conn.execute "update 用户 set 状态='死',死亡时间=now(),宝物='无',宝物修练=0,allvalue=allvalue-"&all&",事件原因='"&sjjh_name&"|发招："&fn1&"' where 姓名='" & to1 &"'"
conn.execute "update 用户 set allvalue=allvalue+"&int(all/2)&" where 姓名='" & sjjh_name &"'"
attack=xinxi & "##<bgsound src=wav/daipu.wav loop=1>用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"%%,%%慢慢倒下..江湖从此又少了一位大虾!"&gjtx&"%%泡点掉了<font color=red>"&all&"</font>点,有一半被##抢走……"
xx=xinxi & "<font color=blue>"& sjjh_name &"</font><bgsound src=wav/daipu.wav loop=1>为夺取宝物，用["& menpai &"]招式("& fn1 &")修练["&wgdj&"重]{"&wgsm&"}攻击"&gif&"<font color=blue>"& to1 &"</font>,<font color=blue>"& to1 &"</font>慢慢倒下..江湖从此又少了一位大虾!"
call boot(tonamedb,"发招，操作者："&sjjh_name&",["&menpai&"]"&fn1)
bbb=zxrsbd(sjjh_name,mynameid)	'判断在线人数
attack=attack&"<br>"&bbb
xx=xx&"<br>"&bbb
call dbxx(xx)
towho=tonamedb
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','"&fn1&"','人命')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>