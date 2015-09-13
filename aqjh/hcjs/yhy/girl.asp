<%@ LANGUAGE=VBScript codepage ="936" %>
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
rs.open "select ID,性别,银两,内力,内力加,等级,会员,会员等级,操作时间 from 用户 where 姓名='"& aqjh_name &"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你不是江湖中人，不可进入怡红院！');window.close();</script>"
	response.end
end if
yl=rs("银两") 
sex=rs("性别")
hydj=rs("会员等级")
pdhy=rs("会员")
nl=rs("内力")
%><!--#include file="../../config.asp"--><%
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
nlsx=int((rs("等级")*aqjh_nlsx+2000+rs("内力加"))*bf)
if rs("内力")>=rs("等级")*aqjh_nlsx+2000+rs("内力加") then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('对不起，你的内力已达到你的上限，不能再进红园了！');window.close();</script>"
	response.end
end if
if sex="女" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你是女的也来这种地方,小心抓你去做妓女！!');window.close();</script>"
	response.end
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<120 then
	ss=120-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你是不是刚从红园里出来呀？请等上"&ss&"秒再来吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="jiu.asp"-->
<%
rs.close
sql="select * from 妓女 where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('你有没有搞错呀，本园内哪有这个姑娘呀，是不是想姑娘想疯了!');window.close();</script>"
	response.end
end if
jiid=rs("妓女ID")
mingji=rs("姓名")
meimao=rs("美貌度")
yin=int(meimao*5)
if yl<meimao*5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script Language=Javascript>alert('没钱也想来这种高级的地方呀，请止步！!');window.close();</script>"
	response.end
end if
yin1=int(meimao*5/2)
nlj=int(meimao/2)
randomize timer
r=int(rnd*3)+1
newnl=nl+nlj
if newnl>=nlsx then
	newnl=nlsx
end if
select case r
	case 1
		mess=aqjh_name &"在快乐怡红园和小姐"&mingji&"促膝长谈人生哲理，直至深夜，"&aqjh_name&"体力消耗五百点，魅力增加100点，内力增加"& nlj &"点，付给红园小姐"&mingji&"银两"&int(yin)&"，红园从中收取一半银两，小姐"&mingji&"美貌度增加10点！"
		conn.execute "update 用户 set 银两=银两-" & yin & ",魅力=魅力+100,体力=体力-500,内力=" & newnl & ",操作时间=now() where 姓名='" & aqjh_name & "'"
		
		connt.execute "update 妓女 set 美貌度=美貌度+10 where id=" & id
	case 2
		mess=aqjh_name &"在快乐怡红园和小姐"&mingji&"聊天，没想到话不投机半句多，没说两句就吵了起来，"& aqjh_name &"道德下降100点，体力下降300点，还白白花了"& yin &"两银子。小姐"& mingji &"美貌度下降10点！"
		conn.execute "update 用户 set 银两=银两-" & yin & ",道德=道德-100,体力=体力-300,操作时间=now() where 姓名='" & aqjh_name & "'"
		
		connt.execute "update 妓女 set 美貌度=美貌度-10 where id=" & id
	case 3
		mess=aqjh_name &"在快乐怡红园和小姐"&mingji&"大谈人生，从中领悟一些真缔，道德上升100点，内力增加"& int(meimao/2) &"点，临走时放下"& yin &"两银子，小姐"& mingji &"得到基中的一半，美貌度上升20点！"
		conn.execute "update 用户 set 银两=银两-" & yin & ",道德=道德+100,体力=体力-500,内力=" & newnl & ",操作时间=now() where 姓名='" & aqjh_name & "'"
		
		connt.execute "update 妓女 set 美貌度=美貌度+20 where id=" & id
	case 4
		mess=aqjh_name &"在快乐怡红园对小姐"& mingji &"不怀好意，动手动脚，道德败坏，被红院保镖暴打一顿，体力下降100，内力下降"& int(meimao/3) &"，道德下降100，魅力下降100！"
		conn.execute "update 用户 set 体力=体力-100,内力=内力-int(meimao/3),道德=道德-100,魅力=魅力-100,银两=银两-" & yin & ",操作时间=now() where 姓名='" & aqjh_name & "'"
end select
says="<font color=#ff0000><b>【怡红院】</b></font>"&mess			'聊天数据
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
%>
<html><head>
<title>歌舞宴会</title>
<style type="text/css">
<!--
table { border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font { font-size: 12px}
.unnamed1 { font-size: 9pt}
-->
</style>
</head>
<body bgcolor="#FFB366">
<p>&nbsp;</p>
<table width="52%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td height="17" bgcolor="#996633" align="center">&nbsp;</td>
</tr>
<tr bgcolor="#66FF66"> 
<td align="center" height="378" bgcolor="#FFCC66"> 
<p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400">
<param name=movie value="girl.swf">
<param name=quality value=high>
<embed src="girl.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400">
</embed> 
</object><font> </font></p>
</td>
</tr>
<tr bgcolor="#0033CC"> 
<td align="center" height="26" class="unnamed1" bgcolor="#FFCC66"><b></b></td>
</tr>
</table>
<p>&nbsp;</p>
</body></html>