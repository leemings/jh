<!--#include file="../config.asp"-->
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
co=session("yx8_mhjh_usercorp")
if un="" then Response.Redirect "../error.asp?id=016"
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
sqlstr="SELECT 最后领钱 FROM 用户 where 姓名='"&un&"' and 会员=True"
Set Rst=conn.Execute(sqlstr)
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=506"
rst.Close
chatroomsn=Session("yx8_mhjh_userchatroomsn")
if yx8_mhjh_fellow=false then
getmoney="现在会员功能关闭，到你的门派去领钱呀！"
else
nowdate=date()
nowdatetype="#"&month(nowdate)&"/"&day(nowdate)&"/"&year(nowdate)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select 500000 as 工资,最后领钱 from 用户 where 姓名='"&un&"'",conn
money=rst("工资")
lastdate=rst("最后领钱")
if money>500000 then money=500000
rst.Close
set rst=nothing
if datediff("d",lastdate,nowdate)>0 then
conn.Execute "update 用户 set 最后领钱="&nowdatetype&",银两=银两+"&money&" where 姓名='"&un&"'"
talkarr=Application("yx8_mhjh_talkarr")
talkpoint=clng(Application("yx8_mhjh_talkpoint"))
dim newtalkarr(600)
j=1
for i=11 to 600 step 10
newtalkarr(j)=talkarr(i)
newtalkarr(j+1)=talkarr(i+1)
newtalkarr(j+2)=talkarr(i+2)
newtalkarr(j+3)=talkarr(i+3)
newtalkarr(j+4)=talkarr(i+4)
newtalkarr(j+5)=talkarr(i+5)
newtalkarr(j+6)=talkarr(i+6)
newtalkarr(j+7)=talkarr(i+7)
newtalkarr(j+8)=talkarr(i+8)
newtalkarr(j+9)=talkarr(i+9)
j=j+10
next
newtalkarr(591)=talkpoint+1
newtalkarr(592)=2
newtalkarr(593)=1
newtalkarr(594)=un
newtalkarr(595)=un
newtalkarr(596)=""
newtalkarr(597)="#660099"
newtalkarr(598)="#660099"
newtalkarr(599)="<font color=FF0000>【领钱】</font>##从会员金库中领取了"&money&"两银子！<font class=timsty>（"&time()&"）</font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
getmoney="你从会员金库中领取了"&money&"两银子！"
else
getmoney="会员金库的管理人员对你说：'你今天刚来领过了银子，现在又想要呀，没有！'"
end if
conn.close
set conn=nothing
end if
%>
<html>
<head>
<title>领会费</title>
<link rel=stylesheet href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<div align=center>
<b><font size="4">领 会 费</font></b>
<p>　
</div>
<%=getmoney%>
</body>
</html>
