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
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from �û� where ����='"&un&"'",conn
nowcorp=rst("����")
if nowcorp="��" then
	leftcorp="��Ȼ�������ɼΣ�����ʵ���ް����"
elseif rst("����")<=1000 or rst("����")<=1000 or rst("�ȼ�")<2 then
	leftcorp="��Ȼ�������ɼ�,������������Ҫ���Ұ�����Ӵ"
else
	conn.Execute "update �û� set ����=����*0.8,����=����*0.8,����=����*0.8,����=����*0.8,�ȼ�=0,����=����*0.8,����=����*0.8,����='��',���='��' where ����='"&un&"'"
	onlinelist=Application("Ba_jxqy_onlinelist")
	onlinelistubd=ubound(onlinelist)
	for i=1 to onlinelistubd step 8
		if onlinelist(i)=un then
			onlinelist(i+2)="��"
			exit for
		end if
	next
	Application.Lock
	Application("Ba_jxqy_onlinelist")=onlinelist
	Application.UnLock
	erase onlinelist
	session("Ba_jxqy_usercorp")="��"
	session("Ba_jxqy_usergrade")=0
	leftcorp="<font color=FF0000>���Ѱ</font>##�뿪��"&nowcorp&",�ȼ���Ϊ0����������������½������ɣ�"
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
	newtalkarr(593)=0
	newtalkarr(594)=un
	newtalkarr(595)="���"
	newtalkarr(596)=""
	newtalkarr(597)="#660099"
	newtalkarr(598)="#660099"
	newtalkarr(599)="<font color=FF0000>���Ѱ</font>##�뿪��"&nowcorp&",�ȼ���Ϊ0����������������½������ɣ�<font class=timsty>��"&time()&"��<\/font>"
	newtalkarr(600)=chatroomsn
	Application.Lock
	Application("Ba_jxqy_talkpoint")=talkpoint+1
	Application("Ba_jxqy_talkarr")=newtalkarr
	Application.UnLock
	erase newtalkarr
	erase talkarr	
end if
rst.Close
set rst=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>ʹ��ҩƷ</title>
<link rel=stylesheet href='css.css'>
<script language=javascript>
setTimeout("location.href='onlinelist.asp'",3000);
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor='<%=bgcolor%>' background='<%=bgimage%>'>
<div align=center>
<font color=0000ff size=5>ʹ��ҩƷ</font>
<hr>
�������Զ�����<br>
<input type=button value='����' onclick="javascript:location.href='onlinelist.asp'" id=button1 name=button1>
</div>
<%=leftcorp%>
</body>
</html>