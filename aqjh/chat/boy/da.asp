<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select boy,boysex,w1,配偶,金币,喂养时间 from 用户 where 姓名='"&aqjh_name&"'",conn,3,3
myhua=rs("w1")
zz=rs("配偶")
if isnull(rs("boy")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您还没有小孩，赶快生育一个吧！');</script>"
	response.end
end if
sj=DateDiff("s",rs("喂养时间"),now())
if sj<30 then
s=30-sj
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('孩子工作要30秒一次，请等"& s &"秒,可别累着孩子！');parent.f3.location.href = 'boy.asp';</script>"
response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：小孩数据出错，请重新购买！');</script>"
	response.end
end if

mypic=rs("boysex")
if DateDiff("h",zt(7),now())<5 then
	zg=clng(zt(6))
else
	zg=0
end if
if clng(zt(3))<3000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：小孩体力少于3000，很虚弱，不能打工了！');parent.f3.location.href = 'boy.asp';</script>"
response.end
end if
if DateDiff("d",zt(2),now())<15 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：孩子还没长大，不能帮忙赚钱，15天才能打工赚钱！');parent.f3.location.href = 'boy.asp';</script>"
response.end
end if
if clng(zt(5))<2000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：小孩不开心，不能打工了，作为父母的就陪他去玩吧！');parent.f3.location.href = 'boy.asp';</script>"
response.end
end if
if zg>=(4+int(DateDiff("d",zt(2),now())/10)) then
	Response.Write "<script Language=Javascript>alert('提示：您已经照顾"&(4+int(DateDiff("d",zt(2),now())/10))&"次了，请等5小时再操作！');parent.f3.location.href = 'boy.asp';</script>"
	response.end
end if
if zt(1)="男" then
'名字|类别|生日|体力|攻击|心情|照顾数次|照顾时间
'金龍|金龍|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29|2002-6-28 11:04:29|2002-6-28 11:04:29
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-250)&"|"&(clng(zt(4))-10)&"|"&(clng(zt(5))-10)&"|"&(zg+1)&"|"&now()  
 Response.Write "<script Language=Javascript>alert('提示：小孩打工赚钱完成：体:-250 攻:-10 感:-10 ！你金币增加2枚!');parent.f3.location.href = 'boy.asp';</script>"
conn.execute "update 用户 set 金币=金币+2,boy='"&temp&"',喂养时间=now() where  姓名='"&aqjh_name&"'"
conn.execute "update 用户 set boy='"&temp&"' where  姓名='"&zz&"'" 
says="<font color=red><b>【孩子打工】</b></font><font color=blue>"&aqjh_name&"<font color=green>看着已经长大的小帅哥<img src=/chat/boy/IMAGES/boy.gif><font color=blue>"&zt(0)&"</font>已经能为父母打工赚钱了，好高兴，<font color=blue>"&zt(0)&"</font>帮父母<font color=blue>"&zz&"</font>跟<font color=blue>"&aqjh_name&"</font>打工赚到了<font color=red>2</font>枚金币，体力减少了250点，能为父母尽点心这点疲劳不算什么~~"
else
'名字|类别|生日|体力|攻击|心情|照顾数次|照顾时间
'金龍|金龍|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29|2002-6-28 11:04:29|2002-6-28 11:04:29
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-250)&"|"&(clng(zt(4))-10)&"|"&(clng(zt(5))-100)&"|"&(zg+1)&"|"&now()  
 Response.Write "<script Language=Javascript>alert('提示：小孩当家教赚钱：体:-250 攻:-10 感:-100 ！你金币增加3枚!');parent.f3.location.href = 'boy.asp';</script>"
conn.execute "update 用户 set 金币=金币+3,boy='"&temp&"',喂养时间=now() where  姓名='"&aqjh_name&"'"
conn.execute "update 用户 set boy='"&temp&"' where  姓名='"&zz&"'" 
says="<font color=red><b>【孩子教书】</b></font><font color=blue>"&aqjh_name&"<font color=0088FF>看着已经亭亭玉立的小丫头<img src=/chat/boy/IMAGES/GIRL.gif><font color=blue>"&zt(0)&"</font>虽然年纪还小，但学问已经特高了，<font color=blue>"&zt(0)&"</font>担任了小学学生的英语家教，教书赚到了<font color=red>3</font>枚金币，体力减少了250点，真是孝顺的孩子啊~~"
end if
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