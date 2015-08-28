<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
co=session("yx8_mhjh_usercorp")
if un="" then Response.Redirect "../error.asp?id=016"
chatroomsn=Session("yx8_mhjh_userchatroomsn")
lastip=Request.ServerVariables("LOCAL_ADDR")
templastip=StrReverse(left(lastip&"9fhb6utime,./",15))
templenlastip=len(lastip)
mmcheck=""
for j=1 to 15
mmcheck=mmcheck+chr(asc(mid(templastip,j,1))-templenlastip+int(j*1.1))
next


set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rs=server.CreateObject ("adodb.recordset")
sql="SELECT 身份 FROM 用户 where 姓名='"&un&"'"
Set Rs=conn.Execute(sql)
shfen=rs("身份")
rs.Close
set rs=nothing
if co="无" or co="出家" or co="十八地狱" then
getmoney="你无门无派，到哪儿去领钱呀！"
elseif shfen="掌门" then
getmoney="你是本派掌门，不是弟子，不能来这里领钱，请到“门派事宜”的“掌门领钱”处领取！"
else
nowdate=date()
nowdatetype="#"&month(nowdate)&"/"&day(nowdate)&"/"&year(nowdate)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select tc.工资系数*("&gr&"+1)+tu.积分*2 as 工资,tu.最后领钱日期 from 用户 tu inner join 门派 tc on tu.门派=tc.门派 where tu.姓名='"&un&"'",conn
money=rst("工资")
lastdate=rst("最后领钱日期")
if money>1000000 then money=1000000
rst.Close
set rst=nothing
if datediff("d",lastdate,nowdate)>0 then
conn.Execute "update 用户 set 最后领钱日期="&nowdatetype&",银两=银两+"&money&" where 姓名='"&un&"'"
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
newtalkarr(599)="<font color=FF0000>【领钱】</font>##从"&co&"金库中领取了"&money&"两银子！<font class=timsty>（"&time()&"）</font>"
newtalkarr(600)=chatroomsn
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
getmoney="你从"&co&"金库中领取了"&money&"两银子！"
else
getmoney=co&"金库的管理人员对你说：'你今天刚来领过了银子，现在又想要呀，没有！'"
end if
conn.close
set conn=nothing
end if
%>
<html>
<head>
<title>领钱</title>
<link rel=stylesheet href='css3.css'>
<script language=javascript>
setTimeout("location.href='onlinelist.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false background='bg1.gif'>
<div align=center>
<font color=0000ff>领 工 资</font>
<hr noshade size="1" color=red>
<input type=button value='在线名单' onclick="javascript:location.href='onlinelist.asp'" id=button1 name=button1>
</div>
<%=getmoney%>
</body>
</html>
