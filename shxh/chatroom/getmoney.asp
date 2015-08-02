<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=session("Ba_jxqy_username")
co=session("Ba_jxqy_usercorp")
if un="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
if co="无" then
	getmoney="你无门无派，到哪儿去领钱呀！"
else
	nowdate=date()
	nowdatetype="#"&month(nowdate)&"/"&day(nowdate)&"/"&year(nowdate)&"#"
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	rst.Open "select tc.工资系数*("&gr&"+1)+tu.积分\10 as 工资,tu.最后领钱日期 from 用户 tu inner join 门派 tc on tu.门派=tc.门派 where tu.姓名='"&un&"'",conn
	money=rst("工资")
	lastdate=rst("最后领钱日期")
	if money>100000 then money=100000
	rst.Close
	set rst=nothing
	if datediff("d",lastdate,nowdate)>0 then
		conn.Execute "update 用户 set 最后领钱日期="&nowdatetype&",银两=银两+"&money&" where 姓名='"&un&"'"
		talkarr=Application("Ba_jxqy_talkarr")
		talkpoint=clng(Application("Ba_jxqy_talkpoint"))
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
		Application("Ba_jxqy_talkpoint")=talkpoint+1
		Application("Ba_jxqy_talkarr")=newtalkarr
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
<link rel=stylesheet href='css.css'>
<script language=javascript>
setTimeout("location.href='onlinelist.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<div align=center>
<font color=0000ff size=5>领 工 资</font>
<hr>
<input type=button value='在线名单' onclick="javascript:location.href='onlinelist.asp'" id=button1 name=button1>
</div>
<%=getmoney%>
</body>
</html>