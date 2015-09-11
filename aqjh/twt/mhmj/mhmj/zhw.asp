<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then 
Response.Write "<script language=javascript>{alert('提示：您没有登陆或已经超时断开连接，请重新登陆！');parent.history.go(-1);}</script>" 
Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER")
if InStr(http,"mhmj/xxw.asp")=0 then 
Response.Write "<script language=javascript>{alert('提交数据非法，程序拒绝执行！');parent.history.go(-1);}</script>" 
Response.End 
end if

if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
gwm="变异吸血蚊"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb") 
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
nl=rs("内力")
fl=rs("法力")
hydj=rs("会员等级")
select case hydj
	case 0
		jgsj=320
	case 1
		jgsj=260
	case 2
		jgsj=220
	case 3
		jgsj=180
	case 4
		jgsj=130
	case 5
		jgsj=100
	case 6
		jgsj=60
end select
sj=DateDiff("s",rs("操作时间"),now())
if sj<jgsj then
	s=jgsj-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你是"&hydj&"级会员，"&jgsj&"秒可进行一次，请等"&s&"秒后再来猎场！');parent.history.go(-1);}</script>" 
	Response.End
end if
if fl<60000 then 
	Response.Write "<script Language=Javascript>alert('提示：召唤怪物需要60000点法力！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select * from 怪物 where 姓名='"&gwm&"'",conn
zt=rs("状态")
if zt="正常" then
	Response.Write "<script Language=Javascript>alert('提示：该怪物已被唤醒，你不需要重复操作！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
conn.execute "update 怪物 set 体力=70000000,杀气=2000000,攻击=1000000,防御=100000,等级=50,状态='正常'  where 姓名='"&gwm&"'"
conn.execute "update 用户 set 法力=法力-60000 where 姓名='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=RED><b>〖唤醒怪物〗</b></font><font color=green>[<font color=red><b>"& aqjh_name &"</b></font>]耗费[<font color=red><b>60000</b></font>]点法力成功的将已经死亡的[<font color=red><b>"&gwm&"</b></font>]唤醒！看来江湖又将是一场血雨腥风！不知道又将有多少江湖义士埋骨荒野。。。</font>"			'聊天数据
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
<script Language=Javascript>alert('提示：<%=aqjh_name%>口中念念有词，耗费60000点法力将<%=gwm%>成功唤醒！');location.href = 'javascript:history.go(-1)';</script>
