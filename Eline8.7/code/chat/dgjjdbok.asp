<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../config.asp"-->
<!--#include file="sjfunc/chatconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"http://127.0.0.1/chat")=0 and InStr(http,"http://127.0.0.1/chat")=0then
	Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');}</script>" 
	Response.End 
end if 
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');</script>"
	Response.End
end if
nowinroom=session("nowinroom")
call roompd("夺宝独孤九剑")
to1=trim(request.form("to1"))
dgjjzs=trim(request.form("dgjjzs"))
tl=int(abs(clng(request.form("tl"))))
nl=int(abs(clng(request.form("nl"))))
maxtl=int(abs(clng(request.form("maxtl"))))
maxnl=int(abs(clng(request.form("maxnl"))))
money=int(abs(clng(request.form("money"))))
zsdj=int(abs(clng(request.form("dj"))))
if Instr(at,chr(39))<>0 or Instr(at,chr(34))<>0 then
	Response.Write "<script Language=Javascript>alert('你要作什么？滚远点！');location.href = 'javascript:history.go(-1)'</script>"
	response.end
end if
if tl<maxtl or nl<maxnl then
	Response.Write "<script Language=Javascript>alert('提示：要使用["&dgjjzs&"]，最少要使用"&maxtl&"体力，"&maxnl&"内力。');</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：所攻击的人不在江湖中！');</script>"
	Response.End
end if
if to1="大家" or to1=Application("sjjh_automanname") or to1=sjjh_name then
	Response.Write "<script Language=Javascript>alert('提示：你不可以对大家、自已或江湖机器人发招。');</script>"
	Response.End
end if
stunt=""
stunt1=""
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='" & sjjh_name & "'",conn,2,2
mp=rs("门派")
if rs("grade")>=6 and rs("grade")<10 then
	Response.Write "<script Language=Javascript>alert('提示：你是官府人员，不能使用["&dgjjzs&"]！');</script>"
	myclose()
end if
if rs("银两")<money then
	Response.Write "<script Language=Javascript>alert('提示：你没有"&money&"两银子，不能使用["&dgjjzs&"]！');</script>"
	myclose()
end if
if rs("内力")<nl or rs("体力")<tl then
	Response.Write "<script Language=Javascript>alert('提示：你内力或体力不足，不能发此招！');</script>"
	myclose()
end if
if rs("等级")<zsdj then
	Response.Write "<script Language=Javascript>alert('提示：你的等级不够["&zsdj&"],不可以使用此招！');</script>"
	myclose()
end if
sj=DateDiff("s",rs("操作时间"),now())
if sj<15 then
	s=15-sj
	Response.Write "<script Language=Javascript>alert('警告：请等"& s &"秒再发招,可别累着！');location.href = 'javascript:history.go(-1)';</script>"
	myclose()
end if
mynameid=rs("id")
lxjwg1=rs("武功")
lxjgj1=rs("攻击")
mydj=rs("等级")
nn=int(mydj/10)+1
tlsx=rs("等级")*sjjh_tlsx+5260+rs("体力加")
nlsx=rs("等级")*sjjh_nlsx+2000+rs("内力加")
wgsx=rs("等级")*sjjh_wgsx+3800+rs("武功加")
tlbf=tl/tlsx
nlbf=nl/nlsx
wgbf=lxjwg1/wgsx
bf=(nlbf+tlbf+wgbf)/3
rs.close
rs.open "select * from 用户 where 姓名='" & to1 & "'",conn,2,2
jhhy=rs("会员等级")
ntnt=rs("等级")
tjf=rs("通缉")
if rs("grade")>=6 then
	Response.Write "<script Language=Javascript>alert('提示：他是官府人员，对他有什么意见你可以向站长反应！');</script>"
	myclose()
end if
tomp=rs("门派")
lxjwg2=rs("武功")
lxjgj2=rs("防御")
yinliang2=rs("银两")
youid=rs("id")
select case dgjjzs
	case "总决式"
		ssl=5000
	case "破刀式"
		ssl=7000
	case "破剑式"
		ssl=9000
	case "破枪式"
		ssl=11000
	case "破鞭式"
		ssl=13000
	case "破索式"
		ssl=15000
	case "破箭式"
		ssl=17000
	case "破掌式"
		ssl=19000
	case "破气式"
		ssl=21000
end select
randomize timer
sss=int(rnd*99)+1
y=int(rnd*999)+1
qq1=lxjwg1+lxjgj1+ssl-(lxjwg2+lxjgj2)
qq2=(tl+nl)*sss
qq3=(qq1+qq2)*bf
qq4=sqr(qq3)
killer=int(qq4*nn+money/y)
if killer<=100 then
	randomize timer
	killer=int(rnd*99)+1
end if
shengtili=rs("体力")-killer
conn.execute "update 用户 set 体力=体力-"  & killer & " where 姓名='" & to1 & "'"
conn.execute "update 用户 set 道德=道德-100,银两=银两-"&money&",体力=体力-" & tl & ",内力=内力-" & nl & ",操作时间=now() where 姓名='" & sjjh_name & "'"
e=""
if shengtili<-100 then
	conn.execute "update 用户 set 杀人数=杀人数+1,总杀人=总杀人+1,银两=银两+" & yinliang2 & " where 姓名='" & sjjh_name & "'"
	conn.execute "update 用户 set 体力="&shengtili&",状态='死',银两=0,事件原因='"&sjjh_name&"|独孤九剑之"&dgjjzs&"' where 姓名='" & to1 & "'"
	e="点，<font color=blue>" & to1 & "</font><bgsound src=wav/daipu.wav loop=1>慢慢的倒了下去……从此江湖上又少了一只大虾," & sjjh_name & "得到" & to1 & "的所有现金<font color=blue><b>" & yinliang2 & "</b></font>两!"
	call boot(to1,"独孤九剑，操作者："&sjjh_name&",["&mp&"]"&dgjjzs)
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','独孤九剑之"&dgjjzs&"','人命')"
end if
if e="" then
	e="点。"
end if
stunt="<font color=blue>"& sjjh_name & "</font><bgsound src=wav/bsj.wav loop=1>运足" & nl & "点内力，" & tl & "点体力，对<img src='img/021.gif'><font color=blue>" & to1 & "</font>使用了江湖上久已传的独孤九剑之<font color=008000><b>【"&dgjjzs&"】</b></font>，杀伤" & killer & e
if shengtili<-100 then
	bbb=zxrsbd(sjjh_name,mynameid)
	stunt=stunt&"<BR><BR>"&bbb
	newstunt="为夺取宝物，"&stunt
	call dbxx(newstunt)
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('恭喜，您的"&dgjjzs&"已经发招完成！');</script>"
says="<font color=red><b>【独孤九剑之"&dgjjzs&"】</b>"& stunt &"</font>"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho=to1
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
function myclose()
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End
end function
%>