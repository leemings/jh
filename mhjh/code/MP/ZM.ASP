<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("yx8_mhjh_username")
co=session("yx8_mhjh_usercorp")
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
sql="SELECT 身份 FROM 用户 where 姓名='"&Session("yx8_mhjh_username")&"'"
Set Rs=conn.Execute(sql)
shfen=rs("身份")
if un="" then Response.Redirect "../error.asp?id=016"
chatroomsn=Session("yx8_mhjh_userchatroomsn")
if co="无" then
	getmoney="你无门无派，到哪儿去领钱呀！"
elseif shfen<>"掌门" then
getmoney="你不是本派掌门，不能来这里领钱！"
else
	nowdate=date()
	nowdatetype="#"&month(nowdate)&"/"&day(nowdate)&"/"&year(nowdate)&"#"
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("yx8_mhjh_connstr")
	set rst=server.CreateObject("adodb.recordset")
	rst.Open "select count(*) as cnumber from 用户 where 门派='"&co&"'",conn
	cnumber=rst("cnumber")
	rst.Close
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("yx8_mhjh_connstr")
	set rst2=server.CreateObject("adodb.recordset")
	rst2.Open "select "&cnumber&"*1000+tu.积分*2 as 工资,tu.最后领钱日期 from 用户 tu inner join 门派 tc on tu.门派=tc.门派 where tu.姓名='"&un&"'",conn
	money=rst2("工资")
	lastdate=rst2("最后领钱日期")
	rst2.Close
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
		newtalkarr(599)="<font color=FF0000>【领钱】<\/font>##从"&co&"金库中领取了"&money&"两银子！<font class=timsty>（"&time()&"）<\/font>"
		newtalkarr(600)=chatroomsn
		Application.Lock
		Application("yx8_mhjh_talkpoint")=talkpoint+1
		Application("yx8_mhjh_talkarr")=newtalkarr
		Application.UnLock
		erase newtalkarr
		erase talkarr
		getmoney="你从"&co&"金库中领取了"&money&"两银子！银两数目为弟子数×10000+积分×2"
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
<link rel=stylesheet href='../style.css'>
</head>
<body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'>
<div align=center>
<font size="3">掌门领钱</font>
<hr>
</div>
<font><%=getmoney%></font>
</body>
</html>
