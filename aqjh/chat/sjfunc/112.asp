<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../const3.asp"-->
<%'轩辕
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
onlinekill=aqjh_onlinekill
if onlinenow<onlinekill and chatinfo(0)<>aqjh_chat1 and chatinfo(0)<>aqjh_chat2 then
	Response.Write "<script language=JavaScript>{alert('提示:非恩怨房在线低于"&onlinekill&"人不得动武！');}</script>"
	Response.End
end if
next
%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<jhdj_fz then
	Response.Write "<script language=JavaScript>{alert('提示：发招需要["&jhdj_fz&"]级以上才可以操作！');}</script>"
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
says=Replace(says,"&amp;","")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
iii=int(rnd()*8)+1
says="<font color=green><img src='xx/image/p"&iii&".gif' border='0'>【轩辕】<font color=" & saycolor & ">"+attack(mid(says,i+1),towho)+"</font>"
call chatsay("轩辕",towhoway,towho,saycolor,addwordcolor,addsays,says)

'发招(比拼内力)
function attack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,等级,保护,grade,宝物,死亡时间,登录 from 用户 where 姓名='" & to1 &"'",conn,2,2
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<2 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！%%上线还没到2分钟啊！"
        exit function
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！%%已经出家了，不问江湖是非！"
        exit function
end if
if rs("门派")="官府" and aqjh_grade<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对官府的[%%]进行攻击！"
        exit function
end if
sj=DateDiff("n",rs("死亡时间"),now())
if sj<3 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！%%刚刚被别人杀死！"
        exit function
end if
if rs("等级")<30 and rs("宝物")="无" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行轩辕攻击！%%等级太低了！"
        exit function
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！%%正在练功保护请不要偷袭！"
        exit function
end if
rs.close
rs.open "select 会员等级,门派,保护,防御,攻击,武功,内力,门派,杀人数,暴豆时间,grade,死亡时间,登录,z1,z2,z3,z4,z5,z6 from 用户 where 姓名='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("死亡时间"),now())
if sj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！你刚刚被别人杀死，还是先练3分钟再报仇吧！"
        exit function
end if
dlsj=DateDiff("n",rs("登录"),now())
if dlsj<2 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！你上线还不到2分钟就不要再杀人了！"
        exit function
end if
if rs("会员等级")<1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！你不是等级会员,没有权利使用轩辕招式！"
        exit function
end if
if rs("门派")="出家" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！你是出家人不可以操作！"
        exit function
end if
if rs("保护")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！你正在练功保护请不要发招！"
        exit function
end if
if rs("杀人数")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
	conn.execute "update 用户 set 保护=false where 姓名='" & aqjh_name &"'"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##，你不能对[%%]进行攻击！你作恶太多，超过江湖杀人上限"&Application("aqjh_killman")&"不能再发招了！"
        exit function
end if
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
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
'取出自己装备武器的攻击及特性(并且所有武器耐久-1)
if not(isnull(rs("z1"))) and instr(rs("z1"),"|")<>0 then
	data=split(rs("z1"),"|")
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
if not(isnull(rs("z2"))) and instr(rs("z2"),"|")<>0 then
	data=split(rs("z2"),"|")
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
if not(isnull(rs("z3"))) and instr(rs("z3"),"|")<>0 then
	data=split(rs("z3"),"|")
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
if not(isnull(rs("z4"))) and instr(rs("z4"),"|")<>0 then
	data=split(rs("z4"),"|")
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
if not(isnull(rs("z5"))) and instr(rs("z5"),"|")<>0 then
	data=split(rs("z5"),"|")
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
if not(isnull(rs("z6"))) and instr(rs("z6"),"|")<>0 then
	data=split(rs("z6"),"|")
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
conn.execute "update 用户 set z1='"&z1&"',z2='"&z2&"',z3='"&z3&"',z4='"&z4&"',z5='"&z5&"',z6='"&z6&"' where  姓名='"&aqjh_name&"'"
rs.close
'开始轩辕
rs.open "select id,c,d,e,f,g FROM n WHERE a='" & trim(fn1) & "' and b='" & aqjh_name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你的轩辕并没有这样的武功["& fn1 &"]呀！');}</script>"
	Response.End
end if
wgid=int(rs("id"))
wgwg=int(rs("c"))
wgnl=int(rs("d"))
wgdj=int(sqr(rs("e")/20))+1
if wgdj>=0 and wgdj<2 then
xxtx="法力"
end if
if wgdj>=2 and wgdj<4 then
xxtx="轻功"
end if
if wgdj>=4 and wgdj<6 then
xxtx="防御"
end if
if wgdj>=6 and wgdj<8 then
xxtx="攻击"
end if
if wgdj>=8 and wgdj<10 then
xxtx="体力"
end if
if wgdj>=10 then
xxtx="金币"
end if
wgsm=rs("f")
if rs("g")="随机" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*130)+1 &".gif height=70>"
else
	gif="<img src=../jhmp/gif/"& rs("g")&" height=70>"
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
'对自己判断
rs.open "select * from 用户 where 姓名='" & aqjh_name &"'",conn
mytodel=rs(xxtx)
rs.close
'对方的判断
rs.open "select 姓名,攻击,防御,武功,会员等级,法力,轻功,体力,金币,z1,z2,z3,z4,z5,z6 from 用户 where 姓名='" & to1 &"'",conn
to1=rs("姓名")
yxjgjto=rs("攻击")
yxjfyto=rs("防御")
yxjwgto=rs("武功")
jhhy=rs("会员等级")
todel=rs(xxtx)
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
xxkill=int(killer/10000)
txkill=int(killer/10000)
if killer>0 then
 if todel<xxkill then
	xxkill=todel
	txkill=todel
 end if
 xxtt="吸取对方"
else
 if mytodel<abs(xxkill) then
	xxkill=mytodel
	txkill=mytodel
 else
	xxkill=abs(xxkill)
	txkill=abs(txkill)
 end if
 xxtt="被反吸"
end if
randomize timer
'对武功次数计数
if aqjh_name=Application("aqjh_user") then
	conn.execute "update n set e=e+1 where id="&wgid
end if
randomize timer
ki=int(rnd*500)
if yxjtx<>"无" and yxjtx<>"" then
	if yxjtx<>yxjtotx then
		conn.execute "update 用户 set "&yxjtx&"="&yxjtx&"+"&ki&" where 姓名='" & aqjh_name &"'"
		conn.execute "update 用户 set  "&yxjtx&"="&yxjtx&"-"&ki&" where 姓名='" & to1 &"'"
		gjtx=gjtx&"武器克制对方，吸取对方[<font color=blue><b>"&yxjtx&"</b></font>]共计:+<font color=red>"&ki&"</font>点!"
	else
		conn.execute "update 用户 set "&yxjtx&"="&yxjtx&"-"&ki&" where 姓名='" & aqjh_name &"'"
		conn.execute "update 用户 set  "&yxjtx&"="&yxjtx&"+"&ki&" where 姓名='" & to1 &"'"
		gjtx=gjtx&"武器被对方所克制，被吸取[<font color=blue><b>"&yxjtx&"</b></font>]共计:-<font color=red>"&ki&"</font>点"
	end if
end if
gjtx=gjtx&wpsm&wpsmto
if killer>0 then
	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& int(killer/10) & " where 姓名='" & aqjh_name &"'"
	conn.execute "update 用户 set 体力=体力-"  & killer & " where 姓名='" & to1 &"'"
	conn.execute "update 用户 set "&xxtx&"="&xxtx&"+"  & txkill & " where 姓名='" & aqjh_name &"'"
	conn.execute "update 用户 set "&xxtx&"="&xxtx&"-"  & xxkill & " where 姓名='" & to1 &"'"
	djinfo="##武功高强"
	mekill=abs(int(killer/10))
else
	if abs(killer)>5000 then killer=5000
	conn.execute "update 用户 set 道德=道德-20,内力=内力-" & abs(wgnl) & ",武功=武功-"& abs(wgwg) &",体力=体力-"& abs(killer) & " where 姓名='" & aqjh_name &"'"
	conn.execute "update 用户 set "&xxtx&"="&xxtx&"-"  & xxkill & " where 姓名='" & aqjh_name &"'"
	conn.execute "update 用户 set "&xxtx&"="&xxtx&"+"  & xxkill & " where 姓名='" & to1 &"'"
	mekill=abs(int(killer/10))
	randomize timer
	killer=int(rnd*400)+1
	djinfo="##武功技不如人"
end if
rs.open "select 体力 from 用户 where 姓名='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*23)+1
mywav="wav"&wav&".wav"
ii=int(rnd()*8)+1
if rs("体力")<-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>用轩辕招式<img src='xx/image/img"&ii&".gif' border='0'>("& fn1 &")修练["&wgdj&"重]特效[<font color=red>被反吸"&xxtx&":"& txkill &"</font>]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点,<br>由于%%的内力深厚，##只感觉眼前一些，双腿一蹬，灵魂飞往封神台报到去了……"&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call boot(aqjh_name,"发招，操作者："&aqjh_name&",["&menpai&"]唉，技不如人"&fn1)
	exit function
end if
rs.close
rs.open "select 体力,宝物,门派,通缉,等级 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("体力")>-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=3>用轩辕招式<img src='xx/image/img"&ii&".gif' border='0'>("& fn1 &")修练["&wgdj&"重]特效[<font color=red>"&xxtt&xxtx&":"& txkill &"</font>]{"&wgsm&"}攻击"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]使得%%体力下降:<font color=red>-"& killer &"</font>点,自己被%%的武功震得体力:<font color=red>-"& mekill &"</font>点."&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
	conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1 where 姓名='" & aqjh_name &"'"
	if rs("宝物")=Application("aqjh_baowuname") then
		conn.execute "update 用户 set 宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where 姓名='" & aqjh_name &"'"
		conn.execute "update 用户 set 宝物修练=0,宝物='无',保护=true,内力=100,体力=2000 where 姓名='" & to1 &"'"
		attack="##<img src='xx/image/img"&ii&".gif' border='0'>把%%的宝物:"& Application("aqjh_baowuname") &"抢走，因得到此宝。江湖宝物需要进行修练才可以得到更多的东西！"&gjtx
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		exit function
	end if
if rs("等级")>=20 then
	all=rs("等级")*2
else
	all=2
end if
if rs("通缉")=False then
		conn.execute "update 用户 set 状态='死',死亡时间=now(),allvalue=allvalue-"&all&",事件原因='"&aqjh_name&"|发招："&fn1&"' where 姓名='" & to1 &"'"
		conn.execute "update 用户 set allvalue=allvalue+"&int(all/2)&" where 姓名='" & aqjh_name &"'"
		attack=xinxi & "##<bgsound src=wav/daipu.wav loop=1>用轩辕招式<img src='xx/image/img"&ii&".gif' border='0'>("& fn1 &")修练["&wgdj&"重]特效[<font color=red>吸取对方"&xxtx&":"& txkill &"</font>]{"&wgsm&"}攻击<img src=xx/gif/WG14.GIF>%%,%%慢慢倒下..江湖从此又少了一位大虾!"&gjtx&"%%泡点掉了<font color=red>"&all&"</font>点,有一半被##抢走……"
else
	conn.execute "update 用户 set 通缉=False,状态='死',死亡时间=now(),道德=0,体力=0,内力=0,武功=0,魅力=0,银两=0,存款=int(存款/2),事件原因='"&aqjh_name&"|发招："&fn1&"' where 姓名='" & to1 &"'"
	conn.execute "update 用户 set 银两=银两+3000000 where 姓名='" & aqjh_name &"'"
	attack="攻击通缉犯，"&xinxi & "##<bgsound src=wav/daipu.wav loop=1>用轩辕招式<img src='xx/image/img"&ii&".gif' border='0'>("& fn1 &")修练["&wgdj&"重]特效[<font color=red>吸取对方"&xxtx&":"& txkill &"</font>]{"&wgsm&"}攻击<img src=xx/gif/WG14.GIF>%%,%%慢慢倒下..为此：##得到300万元官府奖励!"&gjtx
end if
call boot(to1,"发招，操作者："&aqjh_name&",["&menpai&"]"&fn1)
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','轩辕:"&fn1&"','人命')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>