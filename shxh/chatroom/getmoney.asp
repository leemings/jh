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
if co="��" then
	getmoney="���������ɣ����Ķ�ȥ��Ǯѽ��"
else
	nowdate=date()
	nowdatetype="#"&month(nowdate)&"/"&day(nowdate)&"/"&year(nowdate)&"#"
	set conn=server.CreateObject("adodb.connection")
	conn.Open Application("Ba_jxqy_connstr")
	set rst=server.CreateObject("adodb.recordset")
	rst.Open "select tc.����ϵ��*("&gr&"+1)+tu.����\10 as ����,tu.�����Ǯ���� from �û� tu inner join ���� tc on tu.����=tc.���� where tu.����='"&un&"'",conn
	money=rst("����")
	lastdate=rst("�����Ǯ����")
	if money>100000 then money=100000
	rst.Close
	set rst=nothing
	if datediff("d",lastdate,nowdate)>0 then
		conn.Execute "update �û� set �����Ǯ����="&nowdatetype&",����=����+"&money&" where ����='"&un&"'"
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
		newtalkarr(599)="<font color=FF0000>����Ǯ��<\/font>##��"&co&"�������ȡ��"&money&"�����ӣ�<font class=timsty>��"&time()&"��<\/font>"
		newtalkarr(600)=chatroomsn
		Application.Lock
		Application("Ba_jxqy_talkpoint")=talkpoint+1
		Application("Ba_jxqy_talkarr")=newtalkarr
		Application.UnLock
		erase newtalkarr
		erase talkarr
		getmoney="���"&co&"�������ȡ��"&money&"�����ӣ�"
	else 
		getmoney=co&"���Ĺ�����Ա����˵��'����������������ӣ���������Ҫѽ��û�У�'"
	end if
	conn.close
	set conn=nothing
end if
%>
<html>
<head>
<title>��Ǯ</title>
<link rel=stylesheet href='css.css'>
<script language=javascript>
setTimeout("location.href='onlinelist.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<div align=center>
<font color=0000ff size=5>�� �� ��</font>
<hr>
<input type=button value='��������' onclick="javascript:location.href='onlinelist.asp'" id=button1 name=button1>
</div>
<%=getmoney%>
</body>
</html>