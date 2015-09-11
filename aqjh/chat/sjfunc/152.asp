<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="../../mywp.asp"-->
<%'寻找魔器
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")

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
says=Replace(says,"&amp;","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>【寻找魔器】<font color=" & sayscolor & ">"+xunfaqi(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)


function xunfaqi(fn1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 等级,w6,w9,法力,操作时间,职业,门派 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,3,3
if DateDiff("s",rs("操作时间"),now())<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('为了防止在乱刷，寻找魔器等300秒钟操作!');}</script>"
	Response.End 
end if
dj=rs("等级")
fla=rs("法力")
w6w=rs("w9")
if fla<5000 then
Response.Write "<script language=JavaScript>{alert('你的法力不够无法施展呀，至少也得5000点啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("门派")="出家"  then
Response.Write "<script language=JavaScript>{alert('失败，你是出家人想做什么，找死啊!');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("职业")<>"冒险家" then
	Response.Write "<script language=JavaScript>{alert('你非冒险家，不能找宝物！！请去职业转换为冒险家！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj<50 then
Response.Write "<script language=JavaScript>{alert('此功能需要50级战斗等级呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if


conn.execute "update 用户 set 法力=法力-5000 where 姓名='" & aqjh_name &"'"
randomize()
r=int(rnd*25)+1
select case r
case 1
fq=add(w6w,"狼牙棒",1)
conn.execute "update 用户 set  w9='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你在家里的床底下偷到了<font color=red>法器狼牙棒</font>，随后钻地洞跑了."
case 2
fq=add(w6w,"破天锥",1)
conn.execute "update 用户 set  w9='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你闯进仙女洞把<font color=red>法器破天锥</font>从其手中抢走，真是世事不公啊."
case 3
fq=add(w6w,"血滴子",1)
conn.execute "update 用户 set  w9='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你在江湖里的一个破庙中发现了被别人丢弃的<font color=red>法器血滴子</font>."
case 4
fq=add(w6w,"抢劫令",1)
conn.execute "update 用户 set  w9='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你真是幸运，难得一见的<font color=red>法器抢劫令</font>都能被你找，真是有福之人啊."
case 5
w6w=rs("w6")
fq=add(w6w,"红宝石",1)
conn.execute "update 用户 set  w6='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你发现%%口袋里红光闪闪,顺手一摸原来是一颗<font color=red>红宝石</font>,于是捡了块碎石头塞进%%的口袋,把红宝石给盗走了."
case 6
w6w=rs("w6")
fq=add(w6w,"绿宝石",1)
conn.execute "update 用户 set  w6='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你发现口袋里绿光闪闪,顺手一摸原来是一颗<font color=red>绿宝石</font>,于是捡了块碎石头塞进%%的口袋,把绿宝石给盗走了."
case 7
w6w=rs("w6")
fq=add(w6w,"蓝宝石",1)
conn.execute "update 用户 set  w6='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你发现口袋里蓝光闪闪,顺手一摸原来是一颗<font color=red>蓝宝石</font>,于是捡了块碎石头塞进%%的口袋,把蓝宝石给盗走了."
case 8
fq=add(w6w,"魔力钻石",1)
conn.execute "update 用户 set  w9='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"

xunfaqi="##你在江湖里的一棵千年古树的树枝上发现一颗<font color=red>魔力钻石</font>，##你的眼光真是尖啊."
case 9
w6w=rs("w6")
fq=add(w6w,"生日蛋糕",1)
conn.execute "update 用户 set  w6='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##过生日，大家为##买了一盒生日蛋糕."
case 10
fq=add(w6w,"绝情刀",1)
conn.execute "update 用户 set  w9='"&fq&"',操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你去绝情门聊天,在门后发现<font color=red>一把绝情刀</font>."
case else
conn.execute "update 用户 set 操作时间=now() where 姓名='"&aqjh_name&"'"
xunfaqi="##你运气真是差差呀，什么都没找到."
end select
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>