<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'同归于尽♀wWw.happyjh.com♀
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(nowinroom),"|")
if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房间不可以使用同归于尽功能！');}</script>"
	Response.End
end if
if sjjh_jhdj<20 then
	Response.Write "<script language=JavaScript>{alert('提示：同归于尽需要20级以上才可以操作！');}</script>"
	Response.End
end if
if sjjh_grade>=6 then
	Response.Write "<script language=JavaScript>{alert('身为官府人员还想和人打架？！');}</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if towho="大家" or towho=sjjh_name or towho=application("sjjh_automanname") or towho="" then
	Response.Write "<script language=JavaScript>{alert('提示：你不可以对大家、机器人或自已进行此项操作！');}</script>"
	Response.End
end if
'对暂离开、点哑穴判断
call dianzan(towho)
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
says="<font color=green>【同归于尽】<font color=" & saycolor & ">"+tgyj(towho)+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

function tgyj(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if Weekday(date())=6 and (Hour(time())=21) and chatinfo(0)="快乐江湖"  then
Response.Write "<script Language=Javascript>alert('提示：[快乐江湖]房间里现在是只给堂主和护法、长老、掌门等进行门派大战的，其他人等在场可让你门派加强，想打架到[高手房间]房间去吧！');parent.f2.document.af.mdsx.checked=true;parent.m.location.reload();</script>"
	Response.End
end if
rs.open "select 门派,体力,内力,武功,防御,等级,grade,保护 from 用户 where 姓名='" & to1 &"'",conn,2,2
if rs("保护")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对方正在练功保护，你不会是想白死吧！');}</script>"
	Response.End
end if
if rs("门派")="出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('他是出家人，早已看破红尘，世间一切恩怨情仇都已与他无关了！');}</script>"
	Response.End
end if
if rs("等级")<15 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不会是想和一个新手同归与尽吧！');}</script>"
	Response.End
end if
if rs("grade")>=6 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不可以对官府人员使用此功能！');}</script>"
	Response.End
end if
if sjjh_jhdj-rs("等级")>15 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的战斗等级高出对方15级以上了，你还打不过他吗？那需要同归于尽呀！');}</script>"
	Response.End
end if
totl=rs("体力")
tonl=rs("内力")
towg=rs("武功")
tofy=rs("防御")
rs.close
rs.open "select 杀人数,保护,门派,体力,内力,武功,攻击,防御,银两,存款,等级,通缉 from 用户 where 姓名='"&sjjh_name&"'",conn,2,2
if rs("门派")="出家" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是出家人，世间一切恩怨情仇都已与你无关了！');}</script>"
	Response.End
end if
if rs("通缉")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是通缉犯，不可以使用此功能！');}</script>"
	Response.End
end if
if rs("保护")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你正在练功，要想和别人同归于尽，必须先出关！');}</script>"
	Response.End
end if
if rs("杀人数")>=int(Application("sjjh_killman")) and sjjh_grade<10 then
	conn.execute "update 用户 set 保护=false where 姓名='" & sjjh_name &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你作恶太多，超过江湖杀人上限"& Application("sjjh_killman") &"不能使用此功能！');}</script>"
	response.end
end if
mytl=rs("体力")
mynl=rs("内力")
mywg=rs("武功")
if mytl<5000 or mynl<3000 or mywg<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：要想和别人同归于尽，体力必须大于5000，内力大于3000，武功大于1000！');}</script>"
	response.end
end if
myfy=rs("防御")
killer=mytl+mynl+mywg+myfy
yingliang=int(rs("银两")/500)

randomize timer
r=int(rnd*1999)+1
killer=killer+yingliang+r*sjjh_jhdj-towg-tofy
if killer<=4000 then
	killer=int(rnd*3999)+1
end if
tlkill=int(killer*2/3)
nlkill=killer-tlkill
shengtl=totl-tlkill
shengnl=tonl-nlkill
shengwg=towg-mywg
if shengwg<0 then shengwg=0
if shengtl<=-100 then
	conn.execute "update 用户 set 体力=-100,内力=-100,武功=0,防御=0,银两=0,状态='死',杀人数=杀人数+1,宝物='无',宝物修练=0,事件原因='"& sjjh_name & "|和"&to1&"同归与尽' where 姓名='"& sjjh_name &"'"
	conn.execute "update 用户 set 体力=" & shengtl & ",内力=" & shengnl & ",武功="& shengwg &",状态='死',宝物='无',宝物修练=0,事件原因='"& sjjh_name &"|与你同归与尽' where 姓名='" & to1 & "'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'" & sjjh_name & "','同归与尽','人命')"
	conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & sjjh_name & "','同归与尽','人命')"
	e="<br>%%也伴随着##去阎王殿报到去了。"
	call boot(to1,sjjh_name&"和你同归于尽")
else
	conn.execute "update 用户 set 体力=-100,内力=-100,武功=0,防御=0,银两=0,状态='死',宝物='无',宝物修练=0,事件原因='"&sjjh_name&"|和"&to1&"同归与尽' where 姓名='"& sjjh_name &"'"
	conn.execute "update 用户 set 体力=" & shengtl & ",内力=" & shengnl & ",武功="& shengwg &" where 姓名='" & to1 & "'"
	e=""
end if
tgyj="##<bgsound src=wav/tgyj.wav loop=1>仰天一声长啸，全身渐渐被红色气团所包围。突然间气团好似流星一般眨眼就飞向%%，江湖里立时闪起一道道刺眼红光，##全身的血肉都化做一道道利箭向%%射去。%%悴不及防，被杀伤体力：" & tlkill & "，内力："& nlkill &"，武功急速下降" & mywg & "点。##临死前的最后呐喊久久回荡在聊天室里……" & e
call boot(sjjh_name,"和"& to1 &"同归于尽")
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
%>