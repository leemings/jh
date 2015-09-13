<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../config.asp"-->
<!--#include file="../../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
pass= md5(trim(Request.form("pass")))
hygrade=int(abs(Request.form("hygrade")))
hymonth=int(abs(Request.form("hymonth")))
money=abs(int(Request.form("money")))
'if not isnumeric(hygrade) then 
if hygrade<>1 and hygrade<>2 and hygrade<>3 then
	Response.Write "<script language=JavaScript>{alert('提示：会员等级输入错！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if hymonth<1 then
	Response.Write "<script language=JavaScript>{alert('提示：快乐江湖免费会员只能申请一个月！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
userip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 姓名='"&session("aqjh_name")&"' and 密码='"&pass&"'",conn,1,2

If Rs.Bof OR Rs.Eof Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('提示：密码输入不正确！');location.href = 'javascript:history.go(-1)';</script>"
			Response.End
end if
oldhy=rs("会员等级")
if rs("会员等级")<>hygrade and oldhy<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你现在已经是会员，只能办理续时服务\n也就是说现在只能申请["&oldhy&"]会员\n要想升级请与站长联系！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("会员等级")>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你现在已经是会员了请会员到期后在申请！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

if rs("会员金卡")<600 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：600会费一个月，嘿嘿！');}</script>"
	Response.End
end if
if rs("会员等级")>1 then
	hydate=rs("会员日期")
	
else
	hydate=date()
end if
n=Year(hydate)
y=Month(hydate)
r=Day(hydate)
y=y+hymonth
if y>12 then 
	y=y-12
	n=n+1
end if
jy=0


if hygrade=1 then jy=0
if hygrade=2 then jy=0
if hygrade=3 then jy=0

if oldhy=1 then oldjy=0
if oldhy=2 then oldjy=0
if oldhy=3 then oldjy=0
if oldhy=4 then oldjy=0

sj=n & "-" & y & "-" & r
rs("会员等级")=hygrade
rs("会员日期")=cdate(sj)
rs("会员金卡")=rs("会员金卡")-600
rs.Update
conn.execute "insert into q(a,b,d,e,f) values ('"& session("aqjh_name") &"',date(),'已~付','会员等级："&hygrade&",免费会员时间：1个月!','"&userip&"')"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：恭喜，您的会员申请完成,会员截止日期为："&sj&"\n请您退出江湖重新进入！');location.href = '../chat/hy.asp';</script>"
Response.End 


says="<marquee width=100% scrollamount=4><font color=#ff0000>【会费换免费会员】</font>" & aqjh_name & "<font color=#ffFFFF>花了600元会费换得厦门江湖总站免费会员一个月，真是有点心疼呀……</font></marquee>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
says=replace(says,chr(13),"")
says=replace(says,chr(10),"")
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




