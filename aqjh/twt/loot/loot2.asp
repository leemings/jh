<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('刷钱是吗？喜欢是吗？点啊，刷啊！！');i=i+1;}top.location.href='../../exit.asp'}</script>
<!--#include file="../../chat/sjfunc/func.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
'if Instr(LCase(Application("aqjh_useronlinename"&nowinroom))," "&LCase(aqjh_name)&" ")=0 then
if session("aqjh_inthechat")<>1 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
id=request("id")
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End 
end if
if aqjh_jhdj<3 then
	Response.Write "<script Language=Javascript>alert（'你还是江湖小辈，就想来这种地方！!');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ID,性别,体力,等级,银两,内力,内力加,操作时间 from 用户 where 姓名='"& aqjh_name &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你不是江湖中人，不可以抢钱庄！');window.close();</script>"
	response.end
end if
if rs("内力")<10000 then
Response.Write "<script language=JavaScript>{alert('你的内力低于10000，没资格抢劫钱庄！');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("体力")<10000 then
Response.Write "<script language=JavaScript>{alert('你的体力低于10000，没资格抢劫钱庄！');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("银两")<1000000 then
Response.Write "<script language=JavaScript>{alert('你的银两低于一百万两，没资格抢劫钱庄！');window.close();}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
yl=rs("银两") 
sex=rs("银两")
nl=rs("内力")
%>
<%
select case hydj
	case 0
		bf=1.2
	case 1
		bf=1.25
	case 2
		bf=1.30
	case 3
		bf=1.35
	case 4
		bf=1.40
end select
if pdhy<>false then
	bf=1.3
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<120 then
	ss=120-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你是不是刚从钱庄出来呀？请等上"&ss&"秒再来吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="loot3.asp"-->
<%
rs.close
sql="select * from 抢劫 where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('你有没有搞错呀，那有这个钱庄!');window.close();</script>"
	response.end
end if
jiid=rs("ID")
mingji=rs("默默姓名")
meimao=rs("默默爱美")
yin=int(meimao*3)
if yl<meimao*3 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('没钱也想来抢这么高级的钱庄呵你想太多了，请止步！!');window.close();</script>"
	response.end
end if
yin1=int(meimao*3/2)
nlj=int(meimao/2)
randomize timer
r=int(rnd*3)+1
newnl=nl+nlj
if newnl>=nlsx then
	newnl=nlsx
end if
select case r
	case 1
	mess=aqjh_name &"手持一把AK47,冲进"&mingji&"，出来的时候手上多了一大袋钱,结果抢到了"&meimao&"两"
		conn.execute "update 用户 set 内力=内力-"&meimao*3&",体力=体力-"&meimao*3&",操作时间=now() where 姓名='"&aqjh_name&"' "
		conn.execute "update 用户 set 银两=银两+"&meimao*3&" where 姓名='"&aqjh_name&"' "
	case 2
	mess=aqjh_name &"手持一把AK47,冲进"&mingji&"，出来的时候手上多了一大袋钱,结果抢到了"&meimao&"两还得到法力500点"
		conn.execute "update 用户 set 法力=法力+500,内力=内力-"&meimao*3&",体力=体力-"&meimao*3&",操作时间=now() where 姓名='"&aqjh_name&"' "
		conn.execute "update 用户 set 银两=银两+"&meimao*3&" where 姓名='"&aqjh_name&"' "
	case 3
	mess=aqjh_name &"手持一把AK47,冲进"&mingji&"，出来的时候手上多了一大袋钱,结果抢到了"&meimao&"两还得到轻功500点"
		conn.execute "update 用户 set 轻功=轻功+500,内力=内力-"&meimao*3&",体力=体力-"&meimao*3&",操作时间=now() where 姓名='"&aqjh_name&"' "
		conn.execute "update 用户 set 银两=银两+"&meimao*3&" where 姓名='"&aqjh_name&"' "	
	case 4
	mess=aqjh_name &"太贪心，抢了"&mingji&"被官府的抓去坐牢罚款100万两!"
	conn.execute "update 用户 set 银两=银两-1000000,状态='牢',登录=now()+1/144,操作时间=now(),事件原因='"&aqjh_name&" 坐牢:"&fn1&"' where 姓名='" & aqjh_name &"'"
    call boot(aqjh_name,"坐牢，操作者："&aqjh_name&","&fn1)
end select
says="<font color=#ff0000><b>【中央新闻报道】</b></font>"&mess			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
Response.Write "<script Language=Javascript>alert('"&mess&"');window.close();</script>"
%>