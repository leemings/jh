<%
	if chatroomsn=1 or chatroomsn=2 or Application("yx8_mhjh_systemname"&chatroomsn)=mycorp or mycorp="官府"  then
	act=Trim(mid(msg,3,3))
	msg=Trim(mid(msg,6))
	if instr(msg,"'")<>0 or (msg="" and  instr(";基金;收尸;天网;暂离;回来;解穴;点穴;对赌;抢亲;买题;断义;关闭;接拜;结拜;开打;关打;废除;帮扶;精力;逐徒;离师;拜师;收徒;传位;冷饮;拉人;查看;警告;发钱;逐出;偷钱;吸星;赠送;偷窃;斩首;结盟;毁盟;升级;降级;",";"&act&";")=0) then
		Response.Write "<script language=javascript>alert('"&act&"动作的后面应该带有合适的语句，如：//"&act&" ×××');</script>"
		Response.End
	elseif instr(onlinename,";"&sendto&";")=0 and instr(";奖励;求婚;君令;掌令;红包;追杀;轩辕;神龙;解穴;对赌;抢亲;断义;接拜;结拜;废除;逐徒;离师;拜师;收徒;传位;冷饮;拉人;攻击;查看;投掷;送钱;警告;封位;册封;逐出;传功;偷钱;吸星;赠送;偷窃;道具;封咒;帮扶;结盟;毁盟;刑罚;罚款;斩首;升级;降级;",";"&act&";")<>0 and sendto<>"武术家" and sendto<>"魔法师" and sendto<>"爱情鸟" then
		Response.Write "<script language=javascript>alert('"&act&"？对谁？？');</script>"
		Response.End
	elseif sendto=username and instr(";加金;奖励;求婚;君令;掌令;红包;追杀;轩辕;神龙;帮扶;对赌;封位;抢亲;断义;接拜;结拜;废除;打坐;精力;逐徒;离师;拜师;收徒;传位;拉人;冷饮;攻击;投掷;送钱;警告;册封;逐出;传功;偷钱;吸星;赠送;偷窃;道具;放大;封咒;刑罚;罚款;斩首;结盟;毁盟;升级;降级;",act)<>0 then
		Response.Write "<script language=javascript>alert('"&act&"？对自己？？');</script>"
		Response.End
	end if
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("yx8_mhjh_connstr")
	set rst=server.CreateObject("adodb.recordset")
	isprivacy=0
	select case act
case "查看"
msg=seeequip(username,sendto)
isprivacy=1
%><!--#include file="send/4.asp"--><%
case "攻打"
msg=gd(username,sendto,msg)
%><!--#include file="send/71.asp"--><%
case "拉人"
msg=huanying(username,sendto)
%><!--#include file="send/5.asp"--><%
case "冷饮"
msg=nenyin(username,sendto,mygrade)
%><!--#include file="send/6.asp"--><%
case "传言"
msg=chuanyan(username,sendto)
%><!--#include file="send/46.asp"--><%
case "心动"
msg=xindong(username,mygrade,msg)
%><!--#include file="send/47.asp"--><%
case "抢亲"
msg=qiqi(username,sendto,msg)
%><!--#include file="send/9.asp"--><%
case "结盟"
msg=jmen(username,sendto,mycorp)
%><!--#include file="send/48.asp"--><%
case "毁盟"
msg=hmen(username,sendto,mycorp)
%><!--#include file="send/49.asp"--><%
case "怒吼"
msg=nuhou(username,sendto,mygrade,msg)
%><!--#include file="send/50.asp"--><%
case "心跳"
msg=xintiao(username,mygrade,msg)
%><!--#include file="send/51.asp"--><%
case "攻击"
msg=attack(username,sendto,msg,mygrade)
%><!--#include file="send/7.asp"--><%
case "投掷"
msg=cast(username,sendto,msg)
%><!--#include file="send/8.asp"--><%
case "送钱"
msg=sendmoney(username,sendto,msg)
%><!--#include file="send/14.asp"--><%
case "拜师"
msg=bais(username,sendto,mygrade)
%><!--#include file="send/15.asp"--><%
case "收徒"
msg=studi(username,sendto,mygrade)
%><!--#include file="send/16.asp"--><%
case "结拜"
msg=jiebai(username,sendto,mygrade)
%><!--#include file="send/17.asp"--><%
case "接拜"
msg=jiebaich(username,sendto)
%><!--#include file="send/18.asp"--><%
case "断义"
msg=duanyi(username,sendto)
%><!--#include file="send/45.asp"--><%
case "帮扶"
msg=bangfu(username,sendto,msg)
%><!--#include file="send/73.asp"--><%
case "离师"
msg=tuoli(username,sendto)
%><!--#include file="send/20.asp"--><%
case "逐徒"
msg=zhutu(username,sendto)
%><!--#include file="send/21.asp"--><%
case "打坐"
msg=dazuo(username,sendto)
%><!--#include file="send/22.asp"--><%
case "放大"
msg=fangda(username,sendto,mygrade,msg)
%><!--#include file="send/23.asp"--><%
case "封咒"
msg=fengzhou(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/24.asp"--><%
case "刑罚"
msg=xingfa(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/26.asp"--><%
case "罚款"
msg=fakuan(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/27.asp"--><%
case "警告"
msg=alert(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/28.asp"--><%
case "斩首"
msg=zhanshou(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/29.asp"--><%
case "发钱"
msg=distributemoney(username,msg)
%><!--#include file="send/30.asp"--><%
case "册封"
msg=degree(username,sendto,mycorp,msg)
%><!--#include file="send/32.asp"--><%
case "封位"
msg=fengwei(username,sendto,msg)
%><!--#include file="send/53.asp"--><%
case "传位"
msg=chuanwei(username,sendto,mycorp)
%><!--#include file="send/31.asp"--><%
case "逐出"
msg=dislodge(username,sendto,mycorp)
%><!--#include file="send/33.asp"--><%
case "传功"
msg=impart(username,sendto,msg)
%><!--#include file="send/34.asp"--><%
case "偷钱"
msg=touqian(username,sendto)
%><!--#include file="send/35.asp"--><%
case "吸星"
msg=burglemp(username,sendto)
%><!--#include file="send/36.asp"--><%
case "偷窃"
msg=burglemoney(username,sendto)
%><!--#include file="send/37.asp"--><%
case "赠送"
msg=present(username,mygrade,sendto,msg)
%><!--#include file="send/38.asp"--><%
case "道具"
msg=card(username,sendto,msg,mycorp)
%><!--#include file="send/39.asp"--><%
case "公告"
msg=bulletin(username,sendto,mycorp,mygrade,chatroomsn,msg)
%><!--#include file="send/40.asp"--><%
case "对赌"
msg=gedu1(username,sendto,msg)
%><!--#include file="send/57.asp"--><%
case "暂离"
msg=leave(username)
%>
<!--#include file="send/54.asp"--><%
case "回来"
msg=comeback(username)
%>
<!--#include file="send/55.asp"--><%
case "神灵"
msg=shenlong(username,sendto,msg)
%>
<!--#include file="send/58.asp"--><%
case "轩辕"
msg=xuanyuan(username,sendto,msg)
%>
<!--#include file="send/59.asp"--><%
case "追杀"
msg=zuisha(username,sendto,mycorp,mygrade,msg)
%>
<!--#include file="send/61.asp"--><%
case "修炼"
msg=xw(username,sendto,msg)
%>
<!--#include file="send/72.asp"--><%
case "收尸"
msg=shoushi(username,mycorp,msg)
%>
<!--#include file="send/63.asp"--><%
case "掌令"
msg=zhangml(username,sendto,mycorp,msg)
%>
<!--#include file="send/65.asp"--><%
case "君令"
msg=junzhul(username,sendto,msg)
%>
<!--#include file="send/66.asp"--><%
case "求婚"
msg=qiuhun(username,sendto,msg)
%>
<!--#include file="send/67.asp"--><%
case "奖励"
msg=jiangli(username,sendto,mycorp,msg)
%>
<!--#include file="send/68.asp"--><%
case "基金"
msg=jijin(username,mycorp)
%>
<!--#include file="send/69.asp"--><%
case "加金"
msg=bujin(username,mycorp,msg)
%>
<!--#include file="send/70.asp"--><%
case else
			msg="<font color=FF0000>【系统】</font>##，对不起，不明命令，无法执行"	
	end select
	act=2
	set rst=nothing
	conn.Close
	set conn=nothing
	else
		Response.Write "<script language=javascript>alert('你不是本帮弟子，岂能在此撒野！');</script>"
		Response.End
	end if
%>