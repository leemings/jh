<%@ LANGUAGE=VBScript codepage ="936" %>
<SCRIPT LANGUAGE="JavaScript">
if(window.name!="d"){
var i=1;while(i<=50){
window.alert("刷钱是吗？喜欢是吗？点啊，刷啊！！");
i=i+1;}top.location.href="../exit.asp"
}
</script>
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
fromname=LCase(trim(request.querystring("fromname")))
toname=LCase(trim(request.querystring("toname")))
if toname<>aqjh_name then
	if fromname<>toname then
		Response.Write "<script language=JavaScript>{alert('你想作什么，人家"&fromname&"也不是在替你算命说！');}</script>"
		Response.End
	end if
		Response.Write "<script language=JavaScript>{alert('发神经吗"&fromname&" 到底是替自己算命还是替别人算命阿！');}</script>"
		Response.End
	else
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 性别 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close	
	set rs=nothing	
	conn.close	
	set conn=nothing	
	Response.Write "<script language=JavaScript>{alert('求签的人不存在，可能有作弊行为，请跟站长反映！');}</script>"	'测4
	Response.End	
end if
rs.close
rs.open "select w9,w6,性别 FROM 用户 WHERE 姓名='"&aqjh_name&"'",conn
tosex=rs("性别")
randomize 
r=int(rnd*9)+1
select case r
case 1
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【姻缘签】：<font color=red>※"& aqjh_name &"※</font>抽到一只姻缘签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支姻缘签.签文上说你今年将会有姻缘所以尽量留意身旁的人.看看是否就是你未来的另一半<font color=red>※"& aqjh_name &"※</font>顺便赠送你一只小鲤鱼让你姻缘更美好.."
	   	 duyao=add(rs("w6"),"小鲤鱼",1)
	     conn.execute "update 用户 set  银两=银两-20000,w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"

case 2
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【事业签】：<font color=red>※"& aqjh_name &"※</font>抽到一支事业签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支事业签.算是上上签.签文上说你今年官运亨通将会步步高升,凡事顺其自然将会一帆风顺<font color=red>※"& aqjh_name &"※</font>顺便赠送你一块木头祝你步步高升.."
	 	 duyao=add(rs("w6"),"木头",1)
	     conn.execute "update 用户 set  银两=银两-20000,w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"
	   
case 3
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【发财签】：<font color=red>※"& aqjh_name &"※</font>抽到一支发财签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支发财签相当的好.照签文上所说你不管做什么财运都会跟这来.真是恭喜你了这是上等签<font color=red>※"& aqjh_name &"※</font>顺便赠送你20万两.."
	   	conn.execute "update 用户 set  银两=银两+200000 where 姓名='"&aqjh_name&"'"
	    conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"


case 4
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【平安签】：<font color=red>※"& aqjh_name &"※</font>抽到一支平安签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支平安签.算是下下签.签文上说你今年犯小人.将会有血光之灾.本算命师帮你赶走那些小人<font color=red>※"& aqjh_name &"※</font>赠送你一条大鲨鱼来赶走那些小人.."
	 	 duyao=add(rs("w6"),"大鲨鱼",1)
	     conn.execute "update 用户 set  银两=银两-20000,w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"
	

case 5
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【健康签】：<font color=red>※"& aqjh_name &"※</font>抽到一支健康签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支健康签.算是下下签.签文上说你今年身体欠佳需要好好休养.不可过份操劳.要不然内力会慢慢消失<font color=red>※"& aqjh_name &"※</font>传授你2000点内力让你身体安康.."
	 	 conn.execute "update 用户 set  银两=银两-20000,内力=内力+2000 where 姓名='"&aqjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"
	
case 6
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【噩运签】：<font color=red>※"& aqjh_name &"※</font>抽到一支噩运签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支噩运签.非常非常的不好.看来你是噩运难逃了本算命师只能消耗你的体力来为你解除噩运<font color=red>※"& aqjh_name &"※</font>扣你体力2000.."
	 	 duyao=add(rs("w6"),"木头",1)
	     conn.execute "update 用户 set  银两=银两-20000,体力=体力-2000 where 姓名='"&aqjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"

case 7
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【好色签】：<font color=red>※"& aqjh_name &"※</font>抽到一支好色签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支好色签.签文上说道你这好色之徒也敢来求签.滚回家去<font color=red>※"& fromname &"※</font>顺便掌你一拳让你以后不在好色<font color=red>※"& aqjh_name &"※</font>武功立刻消失500"
	 	 conn.execute "update 用户 set  银两=银两-20000,武功=武功-500 where 姓名='"&aqjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"
case 8
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【好色签】：<font color=red>※"& aqjh_name &"※</font>抽到一支好色签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支好色签.签文上说道你这好色之徒也敢来求签.滚回家去<font color=red>※"& fromname &"※</font>顺便掌你一拳让你以后不在好色<font color=red>※"& aqjh_name &"※</font>内力立刻消失2000"
	 	 conn.execute "update 用户 set  银两=银两-20000,内力=内力-2000 where 姓名='"&aqjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"
case 9
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【姻缘签】：<font color=red>※"& aqjh_name &"※</font>抽了一只姻缘签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:你的姻缘还未到,必须等待.是你的就是你的不要过份强求.在等两年你心目中的女子将会出现.顺便增你体力2000让女子看见你觉得你更勇猛."
	      conn.execute "update 用户 set 银两=银两-20000 where 姓名='"&aqjh_name&"'"
	      conn.execute "update 用户 set 银两=银两+20000 where 姓名='"&fromname&"'"
case 10
Response.Write "<script language=JavaScript>{alert('希望这只签能够带来给你好运！');}</script>"
chice="【姻缘签】：<font color=red>※"& aqjh_name &"※</font>抽到一只姻缘签大家来看看阿，<font color=red>※"& fromname &"※</font>解签说道:这是一支姻缘签.签文上说你今年将会有姻缘所以尽量留意身旁的人.看看是否就是你未来的另一半<font color=red>※"& aqjh_name &"※</font>顺便赠送你一只小鲤鱼让你姻缘更美好.."
	   	 duyao=add(rs("w6"),"小鲤鱼",1)
	     conn.execute "update 用户 set  银两=银两-20000,w6='"&duyao&"' where 姓名='"&aqjh_name&"'"
	     conn.execute "update 用户 set  银两=银两+20000 where 姓名='"&fromname&"'"
	
end select
set rs=nothing
conn.close
set conn=nothing
Application.Lock
says="<font color=red><b>【算命消息】</b></font>"&chice	
says=replace(says,chr(39),"\'")
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