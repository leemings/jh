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
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">
<%my=aqjh_name
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 then 
	Response.Write "<script Language=Javascript>alert('你想作什么？！');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	coon.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
sj=DateDiff("n",rs("操作时间"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('对不起系统限制，请等["&s&"分钟]再操作！');}</script>"
	Response.End
end if	
if rs("银两")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你没有钱呀，你想作什么？！');window.close();</script>"
	response.end
end if
conn.execute "update 用户 set 道德=道德+"&int(money/500)&",银两=银两-"& money &",操作时间=now() where 姓名='"&aqjh_name&"'"
Response.Write "<script Language=Javascript>alert('您向孤儿院捐献了"& money &"两，道德上涨"&int(money/500)&"点！');window.close();</script>"
rs.close
conn.close
set rs=nothing
set conn=nothing
says="<font color=red>【孤儿院消息】["&aqjh_name&"]</font><font color=green>向孤儿院捐献了"&money&"两,道德上涨："&int(money/500)&"点,大家代表孩子们感谢你了!</font>"
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