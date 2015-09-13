<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="../../const3.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####房间处理#####
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if chatinfo(8)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以使用卡片！');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
f=Minute(time())
aqjh_kptime=aqjh_chatkptime
if f>aqjh_kptime and chatinfo(0)<>aqjh_chat1 then
 Response.Write "<script language=javascript>{alert('提示：非恩怨房间用卡请在每小时的前"&aqjh_kptime&"分');}</script>"
 Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'对暂离开、点哑穴判断
'call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&","&")
'对系统禁止字符处理
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=trim(mid(says,i+1))
if fnn1<>"归来卡" then call dianzan(towho) 
says=kapian(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'卡片
function kapian(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if
if aqjh_name=to1 and instr("捣乱卡;逮捕卡;踢人卡;冬眠卡;抢亲卡;情人卡;查税卡;穷神卡;衰神卡;死神卡;分手卡;降级卡;怪兽卡;归来卡;解除卡;陷害卡;嫁祸卡;吃豆卡;抱抱卡;真爱卡;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能自己进行操作！');</script>"
	Response.End
end if
if to1="大家" and instr("捣乱卡;逮捕卡;踢人卡;冬眠卡;抢亲卡;情人卡;查税卡;穷神卡;衰神卡;死神卡;分手卡;降级卡;怪兽卡;归来卡;解除卡;陷害卡;嫁祸卡;吃豆卡;抱抱卡;真爱卡;种子卡;种花卡;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能大家进行操作！');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select w5,门派 from 用户 where  姓名='"&aqjh_name&"'",conn,2,2
mycard=abate(rs("w5"),fn1,1)
select case fn1
case "福神卡"
 rs.close
 rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("sl")<>"" and rs("sl")<>"无" then
	kapian="<font color=blue>【卡片】<font color=" & saycolor & ">##，你已经有神灵附身了，不能使用["&fn1&"]，请先把身上的神灵送走吧！</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
	conn.Execute ("update 用户 set sl='福神',slsj=now()+3 where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]，神啊，您与我同在...</font>"
case "会员卡"
 rs.close
 rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("等级")<90 and rs("转生")<1 then
	kapian="<font color=blue>【魔法卡片】<font color=" & saycolor & ">##，你的等级太低了，不够90级，不能使用["&fn1&"]，请多泡点吧！</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
 if rs("会员等级")>2 then
	kapian="<font color=blue>【魔法卡片】<font color=" & saycolor & ">##，你不是免费会员，不能使用["&fn1&"]，请联系站长吧！</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
 conn.Execute ("update 用户 set 会员等级=2,会员日期=会员日期+60 where 姓名='"&aqjh_name&"'")
 kapian="<font color=green>【魔法卡片】<font color=" & saycolor & ">##到了90级，使用了["&fn1&"]，会员等级变为2级，会员日期增加2个月，真是爽啊！！</font>"
case "冰神卡"
 rs.close
 rs.open "select allvalue,会员等级,等级,转生,sl from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("等级")>90 or rs("转生")>0 then
		Response.Write "<script language=javascript>alert('提示：90级以上或转生人或2级以上会员无需使用冰神卡！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
 if rs("sl")<>"" and rs("sl")<>"无" then
	kapian="<font color=blue>【卡片】<font color=" & saycolor & ">##，你已经有神灵附身了，不能使用["&fn1&"]，请先把身上的神灵送走吧！</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
	conn.Execute ("update 用户 set sl='冰神',slsj=now()+1 where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]，神啊，您与我同在，愿我泡点速度增加3.5倍...</font>"
case "封杀卡"
 rs.close
 rs.open "select allvalue,会员等级,等级,转生,杀人数 from 用户 where 姓名='"&to1&"'",conn,2,2
 if rs("杀人数")<=2 then
		kapian="<font color=red><bgsound src=wav/003.mid loop=1>【魔法卡片】<font color=" & saycolor & ">"&aqjh_name&","&to1&"今天杀人数没有<img src='picwords/3.gif'>人，不能对他使用封杀卡哦~!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
	conn.Execute ("update 用户 set 杀人数=5,总杀人=总杀人+10 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]，现在对方已经被封杀了！');}</script>"
	kapian="<font color=red><b>【魔法卡片】</b><font color=" & saycolor & ">##偷偷对%%使用["&fn1&"]，使得%%今天再也杀不了人了，还为%%冠上了杀人魔的称号，杀人记录增加<img src='picwords/1.gif'><img src='picwords/0.gif'>人...</font>"
case "财神卡"
 rs.close
 rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("sl")<>"" and rs("sl")<>"无" then
	kapian="<font color=blue>【卡片】<font color=" & saycolor & ">##，你已经有神灵附身了，不能使用["&fn1&"]，请先把身上的神灵送走吧！</font>"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
    end if
	conn.Execute ("update 用户 set sl='财神',slsj=now()+3 where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]，神啊，您与我同在...</font>"
case "穷神卡"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update 用户 set sl='穷神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵与["&to1&"]赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]，神啊，与你同在...</font>"
case "衰神卡"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update 用户 set sl='衰神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵与["&to1&"]赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]，神啊，与你同在...</font>"
case "死神卡"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update 用户 set sl='死神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在神灵与["&to1&"]赋身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷对%%使用了["&fn1&"]，神啊，与你同在...</font>"
case "送神卡"
	conn.Execute ("update 用户 set sl='无',slsj=now() where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]成功解除了神灵赋身！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了["&fn1&"]，神啊，88！see tomorrow!</font>"
case "抱抱卡"
 rs.close
 rs.open "select 性别,等级 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("等级")<18 then
	Response.Write "<script language=JavaScript>{alert('你的等级太低了！');}</script>"
	response.end
 end if
 my_xb=rs("性别")
 rs.close
 rs.open "select 性别 from 用户 where 姓名='"&to1&"'",conn,2,2
 to_xb=rs("性别")
 if my_xb=to_xb then
	Response.Write "<script language=JavaScript>{alert('异性朋友你也抱？是不是同性恋啊？');}</script>"
	response.end
 end if
 conn.Execute ("update 用户 set 体力=体力-100,内力=内力+500,魅力=魅力+500 where 姓名='"&to1&"'")
 conn.Execute ("update 用户 set 体力=体力-100,内力=内力+500,道德=道德-100,allvalue=allvalue+30 where 姓名='"&aqjh_name&"'")
 kapian="<bgsound src=wav/baobao.wav loop=1><font color=green>【卡片】<font color=" & saycolor & ">##对%%使用了抱抱卡，终于如愿以偿的与%%在江湖的大厅疯狂的<img src=card/baobao.gif>起来！真叫人羡慕啊！同时##的经验值增加了30点,你看，把##乐得三天三夜睡不着觉！大叫：<font color=red><b>抱抱卡在手帅哥美女别想走!</b></font>"
case "真爱卡"
 rs.close
 rs.open "select 性别,等级 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("等级")<18 then
	Response.Write "<script language=JavaScript>{alert('你的等级太低了！');}</script>"
	response.end
 end if
 my_xb=rs("性别")
 rs.close
 rs.open "select 性别 from 用户 where 姓名='"&to1&"'",conn,2,2
 to_xb=rs("性别")
 if my_xb=to_xb then
	Response.Write "<script language=JavaScript>{alert('异性朋友你爱什么啊？是不是同性恋啊？');}</script>"
	response.end
 end if
 conn.Execute ("update 用户 set 道德=道德+50,魅力=魅力+50,allvalue=allvalue+50 where 姓名='"&to1&"'")
 conn.Execute ("update 用户 set 道德=道德+50,魅力=魅力+50,allvalue=allvalue-50 where 姓名='"&aqjh_name&"'")
 kapian="<bgsound src=wav/zhenai.wav loop=1><font color=green>【卡片】<font color=" & saycolor & ">##对%%使用了真爱卡，终于如愿以偿的与%%在江湖的大厅疯狂的<img src=card/zhenai.gif>给真爱的%%50经验,爱是无私的奉献，真叫江湖人士眼红啊！同时双方的道德和魅力都上升50点，你看，把##乐得跳起来了，大叫：<font color=red><b>下辈子我还要你爱我可以吗？</b></font>"
case "请神卡"
	dim sl(4)
	sl(0)="福神"
	sl(1)="财神"
	sl(2)="衰神"
	sl(3)="死神"
	sl(4)="穷神"
	randomize timer
	sss=int(rnd*4)+1
	if sss=5 then sss=4
	conn.Execute ("update 用户 set sl='"&sl(sss)&"',slsj=now()+3 where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]现在得到神灵["&sl(sss)&"]赋身！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]，神啊，您与我同在...请得大神["&sl(sss)&"]到自己的身上…</font>"
	erase sl
case "解除卡"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update 用户 set 保护=false,操作时间=now() where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功解决了["&to1&"]的练功状态！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了解除卡片，也不知道哪一位大虾要倒霉了...</font>"
case "陷害卡"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update 用户 set 内力=int(内力/3),体力=int(体力/3) where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]的体力体力只剩原来的1/3了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##终于忍不可忍，拿出自己陷害卡，向%%的头上打去,%%只觉眼前一黑，体力、内力损失大半..</font>"
case "补血卡"
	conn.Execute ("update 用户 set 体力=体力+50000,内力=内力+50000  where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]的体力、体力都增加了5万点！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了补血卡，这一下可是有福了，体力、内力暴涨5万点！</font>"
case "涨钱卡"
	conn.Execute ("update 用户 set 存款=存款+88880000  where  姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]的存款上涨了8888万！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了涨钱卡，自己的小荷包都装不下了，8888万呀！</font>"
case "升点卡"
 rs.close
 rs.open "select 等级,转生,门派 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("门派")="游侠" or rs("门派")="出家" or rs("门派")="天网" or rs("等级")>90 or rs("转生")>0 then
		Response.Write "<script language=javascript>alert('提示：你不是江湖门派的人或者你是转生人，等级高90级的人，不能使用这种卡片！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	conn.Execute ("update 用户 set allvalue=allvalue+1000  where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]使用了升点卡，泡点数上涨1000点！');}</script>"
	kapian="<font color=red>【门派卡片】<font color=" & saycolor & ">##使用了升点卡，呵。。不用泡点也加点，真是有福气，积分增加了1000点！</font>"
case "加金卡"
 rs.close
 rs.open "select 等级,转生,门派,会员金卡,金币,grade from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("门派")="游侠" or rs("会员金卡")>1000 or rs("门派")<>"官府" or rs("金币")>2000 or rs("门派")="出家" or rs("门派")="天网" then
		Response.Write "<script language=javascript>alert('提示：你的金卡金币太多或者你不是官府的人，所以不能使用这种卡片！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	conn.Execute ("update 用户 set 会员金卡=会员金卡+100  where  姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]的金卡上涨了100个！');}</script>"
	kapian="<font color=red>【门派卡片】<font color=" & saycolor & ">##使用了门派专有卡片加金卡，自己的金库都装不下了，100个金卡呀！</font>"
case "加钱卡"
 rs.close
 rs.open "select 等级,转生,门派,存款 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("门派")="游侠" or rs("门派")="出家" or rs("门派")="天网" or rs("存款")>500000000 then
		Response.Write "<script language=javascript>alert('提示：你不是江湖门派或者你的钱多5亿，不能使用这种卡片！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	conn.Execute ("update 用户 set 存款=存款+60000000  where  姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]的存款上涨了6000万！');}</script>"
	kapian="<font color=red>【门派卡片】<font color=" & saycolor & ">##使用了门派卡片加钱卡，自己的小荷包都装不下了，6000万呀！</font>"
case "练功卡"
	conn.Execute ("update 用户 set 武功=武功+10000  where  姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]使用了练功卡，武功上涨1万！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了练功卡，武功可是大幅度上涨，看来江湖又要不太平了！</font>"
case "珍宝卡"
	conn.Execute ("update 用户 set 保护=false,宝物修练=0,宝物='"& Application("aqjh_baowuname") &"' where 姓名='"&aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('您使用了珍宝卡，获得快乐至宝1枚！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##运气真是好，居然使用了传说中的珍宝卡，获得了江湖至宝"&Application("aqjh_baowuname")&"<img src=img/z1.gif>1枚！大家还不快抢！……</font>"
case "攻防卡"
	conn.Execute ("update 用户 set 攻击=攻击+300,防御=防御+300  where  姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]使用了攻防卡，攻击和防御各涨300点！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了攻防卡，攻击和防御大幅度上涨，看来江湖又多出一位高手了！</font>"
case "加点卡"
	conn.Execute ("update 用户 set allvalue=allvalue+5000  where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]使用了加点卡，泡点数上涨5000点！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了加点卡，呵。。不用泡点也加点，真是有福气！</font>"
case "归来卡"
%>
<!--#include file="wnkif.asp"-->
<%
 if Instr(LCase(application("aqjh_zanli")),LCase("!"&to1&"!"))=0 then
  kapian="<img src=card/youyi.jpg><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & sayscolor & ">"&aqjh_name&"对"&to1&"使用了归来卡,可惜{"&to1&"}并不是在暂离状态!白白浪费了一张归来卡！</font>"
 else
  rs.close
  rs.open "select id,会员等级,等级,身份,门派,名单头像,性别,好友名单,配偶,通缉,转生 from 用户 where 姓名='" & to1 &"'",conn,2,2
  jhid=rs("id")
  hydj=rs("会员等级")
  jhdj=rs("等级")
  jhsf=rs("身份")
if Instr(Application("aqjh_guibin"),"|" & to1 & "|")<>0 then
 jhsf="贵宾"
end if
if Instr(Application("aqjh_admin_send"),"|" & to1 & "|")<>0 then 
 jhsf="财神"
end if
if rs("配偶")=Application("aqjh_user") and rs("性别")="女" then
 jhsf="站长夫人"
end if
if Application("aqjh_mengzhu")=to1 then 
 jhsf="武林盟主"
end if
  if rs("通缉")=True then
   jhmp="通缉犯"
  else
   jhmp=rs("门派")
  end if
  jhtx=rs("名单头像")
  sex=rs("性别")
   myzs=rs("转生")
   mypeiou=rs("配偶")
'  nowinroom=session("nowinroom")
  Application.Lock
  onlinelist=Application("aqjh_onlinelist"&nowinroom)
  onlinenum=UBound(onlinelist)
  for i=1 to onlinenum step 1
   onlinexx=split(onlinelist(i),"|")
   if onlinexx(0)=to1 then
   aqjh_zm=to1&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&jhdj&"|"&jhid&"|"&hydj&"|0"&"|"&onlinexx(9)&"|"&mypeiou&"|"&myzs
   onlinelist(i)=aqjh_zm
   exit for
   end if
  next
  Application("aqjh_onlinelist"&nowinroom)=onlinelist
  aqjh_zanli=split(application("aqjh_zanli"),";")
  for x=0 to UBound(aqjh_zanli)
   dxname=split(aqjh_zanli(x),"|")
   if dxname(0)="!"&to1&"!" then
    dxcz=aqjh_zanli(x)&";"
    application("aqjh_zanli")=replace(application("aqjh_zanli"),dxcz,"")
    kapian="<img src=card/youyi.jpg><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & sayscolor & ">"&aqjh_name&"对"&to1&"使用了归来卡,成功地把{"&to1&"}从暂离状态拖回来!</font>"
    exit for
   end if
  next
  Application.UnLock
 end if
case "吃豆卡"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update 用户 set 暴豆时间=now()  where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('对["&to1&"]使用了吃豆卡，他不能再暴豆了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##实在对%%的行为看不过去，使用了吃豆卡，%%大叫一声，晕死过去。豆豆没有了...</font>"
case "暴豆卡"
	conn.Execute ("update 用户 set 暴豆时间=now()  where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('对["&aqjh_name&"]使用了暴豆卡，现在暴豆成功了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##:大叫着，我暴，我暴....从口袋中拿出暴豆卡，暴豆成功！</font>"
case "捣乱卡"
%>
<!--#include file="wnkif.asp"-->
<%
	conn.Execute ("update 用户 set 武功=int(武功/3)  where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('对["&to1&"]使用了捣乱卡，他武功只剩1/3了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##:大叫着，谁也不要拦着我，我要为民除害！使用了捣乱卡，"&to1&"的武功失去大半....</font>"
case "怪兽卡"
%>
<!--#include file="wnkif.asp"-->
<%
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
        if rs("门派")="官府" then 
            kapian="<img src=card/guaishou.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用怪兽卡!"
           	rs.close
		set rs=nothing
		conn.close
	    	set conn=nothing
	        exit function
	end if
        if rs("身份")="掌门" then 
            kapian="<img src=card/guaishou.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&",你不能对门派的掌门人<font color=red>[["&to1&"]]</font>使用怪兽卡!"
           	rs.close
		set rs=nothing
		conn.close
	    	set conn=nothing
	        exit function
	end if
	conn.Execute ("update 用户 set 身份='僵尸王',状态='僵尸王' where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('对["&to1&"]使用了怪兽卡，他已经变成了僵尸了！');}</script>"
	kapian="<img src=card/guaishou.jpg><font color=green>【卡片】<font color=" & saycolor & ">##:大叫着，"&to1&"变成了僵尸，请大家小心哦，如果怕的话请来我这边，我可以保护大家....</font>"
case "清纯卡"
	conn.Execute ("update 用户 set 结婚次数=0  where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('您使用了清纯卡，结婚次数清0完成！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##好怀念过去的时光呀……</font>"
case "亲亲卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" &saycolor& ">##偷偷对%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
    if rs("门派")="出家" then
		Response.Write "<script language=javascript>alert('【"&aqjh_name&"】你是出家人，搞错了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from 用户 where 姓名='"&to1&"'",conn
    if rs("门派")="出家" then
		Response.Write "<script language=javascript>alert('【"&aqjh_name&"】人家是出家人，搞错了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing

    end if
    if rs("性别")=sex then
        kapian="<table border=0><tr><td><img src=card/qinren.jpg></td><td><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" &saycolor& ">"&aqjh_name&"对"&to1&"使用亲亲卡失败:原因是"&to1&"与"&aqjh_name&"是同性!没爱人了玻璃也想啊？</font></td></tr></table>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.Execute ("update 用户 set allvalue=allvalue+50  where 姓名='" &aqjh_name&"'")
	Response.Write "<script language=JavaScript>{alert('["&aqjh_name&"]使用了加点卡，泡点数上涨50点！此卡由回首当年制作！');}</script>"
    kapian="<table border=0><tr><td><img src=card/ai1.gif></td><td><font color=green><bgsound src=123.wav loop=1>【卡片】<font color=" &saycolor& ">"&aqjh_name&"对"&to1&"使用亲亲卡,终于如愿以尝的与"&to1&"在江湖的大厅疯狂的<img src=card/ai.gif>起来!真叫人羡慕啊！同时"&aqjh_name&"的泡点上升50点！你看！把"&aqjh_name&"乐得都跳起来了！大叫：<img src=card/qin.gif></font></td></tr></table>" 
case "变性卡"
	rs.close
	rs.open "select 性别,配偶 FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
	sex=rs("性别")
	pl=rs("配偶")
	rs.close
    if pl="无" then 
        if sex="男" then
               sql="update 用户 set 性别='女' WHERE 姓名='" & aqjh_name & "'"
               xb="漂亮的女生"
         end if 
          if sex="女" then 
              sql="update 用户 set 性别='男' WHERE 姓名='" & aqjh_name & "'"
              xb="英俊的男孩"
          end if
          Set Rs=conn.Execute(sql)  
            bianxi=aqjh_name&"使用变性卡后,终于如愿以尝,变成了"&xb&"!" 
        else
          bianxi="使用变性卡失败!原因:"&aqjh_name&"是有家室的人呢,怎么还想变性呀!这样不是乱套了!为了惩戒像"&aqjh_name&"你这种道德败坏的人,特此没收"&aqjh_name&"的变性卡!"
        end if
	kapian="<img src=card/bxk.jpg></td><td><font color=green>【卡片】<font color=" & saycolor & ">##：这年头，克隆技术真先进呀，我变，我变……(不是在变成克琳钝，在作变性手术！）<br>【结果】"&bianxi&"</font>"
case "离婚卡"
	rs.close
	rs.open "select 配偶,boy FROM 用户 WHERE 姓名='" & aqjh_name &"'",conn,2,2
	peiou=rs("配偶")
	myboy=rs("boy")
	if isnull(myboy) or myboy="" then
		myboy=""
	else
		zt=split(myboy,"|")
		if UBound(zt)<>7 then
			myboy=""
		end if
	end if
	if peiou="无" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('【"&aqjh_name&"】你有钱呀，根本没有配偶呀！');</script>"
		Response.End
	end if
if myboy<>"" then
 conn.execute "insert into gry(boy,fm1,fm2) values ('"&myboy&"','"&aqjh_name&"','"&peiou&"')"
end if
	conn.Execute ("update 用户 set 配偶='无',boy='',boysex='' where 姓名='" & aqjh_name &"'")
	rs.close
	rs.open "select 配偶 FROM 用户 WHERE 姓名='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update 用户 set 配偶='无',boy='',boysex='' where 姓名='"&peiou&"'")
	end if
	rs.close
	kapian="<img src=card/lifen.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">##：想前想后，经过自己一番思想斗争，使用了离婚卡,终于想好了与["&peiou&"]离婚了……</font>"
if myboy<>"" then kapian=kapian&"<br>【孤儿院】双方的小孩已经送到孤儿院！孩子是无辜的啊！唉！"
case "抢亲卡"
%>
<!--#include file="wnkif.asp"-->
<%
    if rs("门派")="出家" then
		Response.Write "<script language=javascript>alert('【"&aqjh_name&"】你是出家人，搞错了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from 用户 where 姓名='"&to1&"'",conn,2,2
    if rs("门派")="出家" then
		Response.Write "<script language=javascript>alert('【"&aqjh_name&"】人家是出家人，搞错了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("性别")
    if rs("配偶")<>"无" then
	    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"对"&to1&"使用抢亲卡失败:原因是"&to1&"是有家室的人</font>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("门派")="官府" then 
           kapian="不可以对管理员使用抢亲卡……"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
    if rs("配偶")<>"无" then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"对"&to1&"使用抢亲卡失败:原因是自己已经是有家室的人了!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("性别")=sex then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"对"&to1&"使用抢亲卡失败:原因是"&to1&"与"&aqjh_name&"是同性!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update 用户 set 配偶='"&aqjh_name&"' where 姓名='"&to1&"'"
    conn.execute "update 用户 set 配偶='"&to1&"' where 姓名='"&aqjh_name&"'"
    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"对"&to1&"使用抢亲卡,终于如愿以尝的与"&to1&"结为夫妇!</font>"
case "逮捕卡"
%>
<!--#include file="wnkif.asp"-->
<%
        rs.close
        rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
        if rs("门派")="官府" then 
            kapian="<img src=card/xianhao.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用逮捕令!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        kapian="<img src=card/xianhao.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"拿出江湖特许的逮捕令,把"&to1&"给抓了!"
        mzky=mzk(to1)
        if mzky="ok" then   
           conn.execute "update 用户 set 状态='狱',登录=now()+3 where 姓名='" & to1 & "'"
            call boot(to1,to1&"被"&aqjh_name&"使用了逮捕令")
        else
           kapian=kapian&mzky
        end if
case "踢人卡"
%>
<!--#include file="wnkif.asp"-->
<%
      rs.close
      rs.open "select 等级,转生,门派,存款,登录 FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
      dlsj=DateDiff("n",rs("登录"),now())
       if dlsj<1 and aqjh_grade<10 then
        kapian="<img src=card/dz04.gif></td><td><font color=red><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&",你不能连续对刚上线的人用踢人卡,请等1分钟吧，别伤了和气!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
end if
      if rs("门派")="官府" then 
        kapian="<img src=card/dz04.gif></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用踢人卡!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
	  end if
       mtky=mtk(to1)
       if mtky="ok" then   
	      kapian="<img src=card/dz04.gif></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"使用踢人卡,飞起一脚，结果把"&to1&"踢了出去!"
    	  call boot(to1,to1&"被"&aqjh_name&"使用了踢人卡")     
    	else
			kapian=mtky
    	end if
case "冬眠卡"      
%>
<!--#include file="wnkif.asp"-->
<%
  rs.close
  rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
  if rs("门派")="官府" then
     kapian="<img src=card/shuimian.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用冬眠卡!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
  end if
   rs.close
   qxky=qxk(to1)
   if qxky="ok" then   
	   conn.execute "update 用户 set 登录=now()+12/(4*144),状态='眠' where 姓名='" & to1 & "'"
	   kapian="<img src=card/shuimian.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"对"&to1&"使用冬眠卡!使"&to1&"睡着了!"
	   call boot(to1,to1&"被"&aqjh_name&"使用了冬眠卡")
   else
   		kapian=qxky
   end if
case "查税卡" 
%>
<!--#include file="wnkif.asp"-->
<%
  rs.close   
  rs.open "SELECT * FROM 用户 WHERE  姓名='" & to1 & "'",conn,2,2
  if rs("门派")="官府" then
    kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&",你不能对官府人员使用查税卡!"
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    exit function
  end if
    yl=rs("银两")+rs("存款")
    if yl<=10000 then
    	kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"对"&to1&"使用查税卡失败:原因"&to1&"身上的银子小于10000两!"
	   	rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	    exit function
	end if
    mhky=mhk(to1)
    if mhky="ok" then   
      yl=int(yl*0.02)
      conn.execute "update 用户 set 银两=银两+"&yl&" where 姓名='"&aqjh_name&"'"
      if rs("银两")>=rs("存款") then
        conn.execute "update 用户 set 银两=银两-"&yl&" where 姓名='"&to1&"'"
      else
        conn.execute "update 用户 set 存款=存款-"&yl&" where 姓名='"&to1&"'"
      end if
      kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"使用查税卡,查得"&to1&"共"&yl&"两银子,全部归"&aqjh_name&"所有!"
    else
       kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"对"&to1&"使用查税卡失败!"
       kapian=kapian&mhky
    end if
case "升级卡"
 rs.close
 rs.open "select allvalue,会员等级,等级 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
 if rs("会员等级")>2 or rs("等级")>90 then
		Response.Write "<script language=javascript>alert('提示：90级以上或3级以上会员无需使用升级卡！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
 jhjy=rs("allvalue")
 jhdj=int(sqr(jhjy/50))
 jhadd=((jhdj+1)*(jhdj+1)-jhdj*jhdj)*50
 jhdj=jhdj+1
 jhjy=jhjy+jhadd
 conn.Execute ("update 用户 set allvalue="&jhjy&",等级="&jhdj&" where 姓名='"&aqjh_name&"'")
 Response.Write "<script language=Javascript>{alert('对["&aqjh_name&"]使用了升级卡！江湖经验值上升了"&jhadd&"点，战斗等级上升1级，现为"&jhdj&"级');}</script>"
 kapian=aqjh_name&"使用了升级卡，"&aqjh_name&"的泡点增加"&jhadd&"点，战斗等级上升1级...真是好福气呀，不泡点也升级！"
case "健身卡"
 conn.Execute ("update 用户 set 武功加=武功加+500,内力加=内力加+500,体力加=体力加+500 where 姓名='"&aqjh_name&"'")
 Response.Write "<script language=Javascript>{alert('对["&aqjh_name&"]使用了健身卡，锻练成功！武功、内力、体力上限值各涨500点！！！');}</script>"
 kapian="【卡片】"&aqjh_name&"大叫着，我要更强，我要更强....从口袋中拿出健身卡，在"&Application("aqjh_user")&"的精心调教下，"&aqjh_name&"经过一翻艰苦训练，武功、内力、体力上限值各涨500点！！！"
case "好人卡"
	conn.Execute ("update 用户 set 杀人数=0  where 姓名='" & aqjh_name &"'")
	Response.Write "<script language=JavaScript>{alert('您使用了好人卡，杀人数清0完成！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##说，我是好人，我今天没有杀人呀……</font>"
case "天眼卡"
	Response.Write "<script>parent.slbox=true;<"&"/script>"	
	Response.Write "<script language=JavaScript>{alert('您使用了天眼卡，现在可以看到别人私聊了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷地天启了天眼，十万百千里的地方都能看见了……</font>"
case "降级卡"
%>
<!--#include file="wnkif.asp"-->
<%
     rs.close
     rs.open "select * from 用户 where 姓名='"&to1&"'",conn,2,2
     jhdj=rs("等级")
     del1=jhdj*jhdj*50
     del2=(jhdj+1)*(jhdj+1)*50
     delok=del2-del1
     jhjf=rs("allvalue")-delok
 	conn.Execute ("update 用户 set 等级=等级-1,allvalue="&jhjf&" where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]的等级降为"&jhdj-1&"级了！');}</script>"
	kapian="<img src=card/jiangji.jpg><font color=green><bgsound src=wav/003.mid loop=1><font color=green>【卡片】<font color=" & saycolor & ">##对%%使用了降级卡,%%的等级降为"&jhdj-1&"级，积分也大为减少.</font>"
case "嫁祸卡"
rs.close
rs.open "select * from 用户 where 姓名='"&to1&"'",conn,2,2
if rs("通缉")="True" then
kapian="<bgsound src=plus/wav/wav/003.mid loop=1><font color=green>【卡片】</font><font color=blue>"&aqjh_name&"</font>对<font color=green>"&to1&"</font>使用<font color=red><b>嫁祸卡</b></font>失败："&to1&"也是通辑犯，已经不用你嫁祸了！！"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  exit function
end if
if rs("门派")="官府" then 
  kapian=""&aqjh_name&"你不可以对管理员使用嫁祸卡……"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  exit function
end if
rs.close
rs.open "select 通缉 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs("通缉")="False" then
kapian="<bgsound src=wav/wav/003.mid loop=1><font color=green>【卡片】</font><font color=blue>"&aqjh_name&"</font>对<font color=green>"&to1&"</font>使用<font color=red><b>嫁祸卡</b></font>失败："&aqjh_name&"不是通辑犯，怎么嫁祸呀！！"
        rs.close
set rs=nothing
conn.close
set conn=nothing
exit function
end if
%>
<!--#include file="wnkif.asp"-->
<%
conn.execute "update 用户 set 通缉=True where 姓名='"&to1&"'"
conn.execute "update 用户 set 通缉=False where 姓名='"&aqjh_name&"'"
kapian="<bgsound src=wav/wav/003.mid loop=1><font color=green>【卡片】</font>##偷偷使用卡片：现在%%是通辑犯大家快杀呀！"
Case "种子卡"
    rs.close
    rs.Open "select hua from 用户 where  姓名='" & aqjh_name & "'", conn,2,2
    myhua = rs("hua")
    If myhua = "" or IsNull(myhua) Or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>【卡片】</font>##你并没有鲜花,不能使用种子卡，请去开垦!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>【卡片】</font>##你的鲜花数据有问题，请重新开垦!"
                Exit Function
        End If
        '增加种子
        newmyhua = ""
        zt = Split(myhua, "|")
        newmyhua = CLng(zt(0)) + 5 & "|"
        For I = 1 To 26
                newmyhua = newmyhua & zt(I) & "|"
        Next
        conn.Execute "update 用户 set hua='" & newmyhua & "' where 姓名='" & aqjh_name & "'"
        kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了种子卡,得到了5颗花种，快去种花吧！</font>"
Case "种花卡"
    rs.close
    rs.Open "select hua from 用户 where  姓名='" & aqjh_name & "'", conn,2,2
    myhua = rs("hua")
    If myhua = "" Or IsNull(myhua) or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>【卡片】</font>##你并没有鲜花,不能使用种花卡，请去开垦!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>【卡片】</font>##你的鲜花数据有问题，请重新开垦!"
                Exit Function
        End If
        newmyhua = ""
        For I = 0 To 26
                If I >= 2 Then
                temp = Split(zt(I), ";")
                If CLng(temp(0)) > 0 And CLng(temp(0)) < 94 Then
                ss = CLng(temp(0)) + 5 & ";" & temp(1) & ";0"
                newmyhua = newmyhua & ss & "|"
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
        Next
        conn.Execute "update [用户] set hua='" & newmyhua & "' where 姓名='" & aqjh_name & "'"
        kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了种花卡,真是爽啊，眼看着花迅速地上涨！</font>"
Case "猪崽卡"
    rs.close
    rs.Open "select zhu from 用户 where  姓名='" & aqjh_name & "'", conn,2,2
    myhua = rs("zhu")
    If myhua = "" or IsNull(myhua) Or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>【卡片】</font>##你并没有养猪啊,请先去重建一下农场再来使用吧!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>【卡片】</font>##你的猪猪数据有问题，请去重建一下农场!"
                Exit Function
        End If
        '增加种子
        newmyhua = ""
        zt = Split(myhua, "|")
        newmyhua = CLng(zt(0)) + 2 & "|"
        For I = 1 To 26
                newmyhua = newmyhua & zt(I) & "|"
        Next
        conn.Execute "update 用户 set zhu='" & newmyhua & "' where 姓名='" & aqjh_name & "'"
        kapian="<font color=green>【卡片】<font color=" & saycolor & ">##掏出腰包中的猪崽卡,大吼道:猪猪超生术！好不容易得到了2只猪崽！</font>"
Case "养猪卡"
    rs.close
    rs.Open "select zhu from 用户 where  姓名='" & aqjh_name & "'", conn,2,2
    myhua = rs("zhu")
    If myhua = "" Or IsNull(myhua) or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>【卡片】</font>##你并没有养猪啊,请先去重建一下农场再来使用吧!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>【卡片】</font>##你的猪猪数据有问题，请去重建一下农场!"
                Exit Function
        End If
        newmyhua = ""
        For I = 0 To 26
                If I >= 2 Then
                temp = Split(zt(I), ";")
                If CLng(temp(0)) > 0 And CLng(temp(0)) < 94 Then
                ss = CLng(temp(0)) + 5 & ";" & temp(1) & ";0"
                newmyhua = newmyhua & ss & "|"
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
        Next
        conn.Execute "update [用户] set zhu='" & newmyhua & "' where 姓名='" & aqjh_name & "'"
        kapian="<font color=green>【卡片】<font color=" & saycolor & ">##摇着手中的养猪卡,对着猪猪唱起了猪猪成长歌，在音乐声中猪猪慢慢长大着！</font>"
case else
	Response.Write "<script language=JavaScript>{alert('系统并没有["&fn1&"]这种卡片,或不能使用你搞错了！');}</script>"
	Response.End
end select
'删除自己卡片，记录
conn.execute "update 用户 set w5='"&mycard&"' where  姓名='"&aqjh_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& aqjh_name &"','"& to1 &"','操作','"& fn1 & "')"
set rs=nothing	
conn.close
set conn=nothing
end function
'免税卡
function mhk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免税卡")=0 then
		rs.close
	    mhk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免税卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   'mhk="<br><font color=green>【免税卡】</font>"&to1&"身上的免税卡生效,因此不能抓他!"
	   mhk="<img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"正准备喜滋滋的从"&to1&"的口袋中拿钱,就在此时,"&to1&"掏出身上的免税卡说,慢着,我身上有免税卡,嘿嘿!"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'免罪卡
function mzk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免罪卡")=0 then
		rs.close
	    mzk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免罪卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   'mzk="<br><font color=green>【免罪卡】</font>"&to1&"身上的免罪卡生效,因此不能抓他!"
	   mzk="<img src=card/myk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&"官府的人准备用根铁索套在"&to1&"的脖子上,就在此时,"&to1&"掏出身上的免罪卡说,慢着,我身上有免罪卡,嘿嘿!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'清醒卡
function qxk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"清醒卡")=0 then
		rs.close
	    qxk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"清醒卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   qxk="><img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"拿出水晶球，睡吧，睡吧，在催眠……"&to1&"睁着个大眼睛，傻嘻嘻的看着他……在说我吗？"&to1&"掏出身上的清醒卡,慢着,我身上有清醒卡,嘿嘿!"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'免踢卡
function mtk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5 from 用户 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免踢卡")=0 then
		rs.close
	    mtk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免踢卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   mtk="<img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&aqjh_name&"使出国产臭脚，准备来个国际行动，却不小心踢到石头……"&to1&"在一边嘿嘿的笑，就你呀，还要踢我，再来20年……"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'万能卡
function wnk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  rs.open "select w5,转生 from 用户 where 姓名='"&to1&"'",conn,2,2
	if rs("转生")>=12 then
		rs.close
		wnk="zstt"
		exit function
	end if
	if iswp(rs("w5"),"万能卡")=0 then
		rs.close
	    wnk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"万能卡",1)
		conn.execute "update 用户 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   wnk="<img src=card/wangneng.jpg></td><td><font color=green style='font-size=9pt'><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&to1&"在一边嘿嘿的笑，就你呀，万能卡，万能卡，一卡在手，走遍天下……"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>