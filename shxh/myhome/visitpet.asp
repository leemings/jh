<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
pid=Request.QueryString("id")
if not isnumeric(pid) then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from pet where username='"&username&"' and option_T<"&nowtimetype&" and exist=false and id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=069"
biological=rst("biological")
sinew=rst("sinew")
maxsinew=rst("maxsinew")
cleanliness=rst("cleanliness")
jiankang=rst("health")
kuaile=rst("happy")
encourage_M=rst("encourage_M")
encourage_C=rst("encourage_C")
beftime=rst("option_T")
decrease=datediff("d",beftime,nowtime)
if decrease=0 then
	msg=biological&"��"&username&"˵:�ּ��������濪��,�����ںܿ���ϣ����Ҳ��"
elseif decrease=1 then
	select case encourage_M
		case "����","����"
			conn.Execute "update �û� set "&encourage_M&"="&encourage_M&"+"&encourage_C&" where ����='"&username&"'"
		case "��Ʒ"
			rst.Close 
			rst.Open "select * from ��Ʒ where ������='"&username&"' and ����='"&encourage_C&"'",conn
			if rst.EOF or rst.BOF then
				rst.close 
				rst.Open "select * from �̵� where ����='"&encourage_C&"'",conn
				conn.Execute "insert into ��Ʒ(����,����,����,����,����,����,��Ч,����,�۸�,������,����,װ��) values('"&rst("����")&"','"&rst("����")&"',"&rst("����")&","&rst("����")&","&rst("����")&","&rst("����")&",'"&rst("��Ч")&"',1,"&rst("�۸�")\2&",'"&username&"',false,false)"
			else
				conn.Execute "update ��Ʒ set ����=����+1 where ������='"&username&"' and ����='"&encourage_C&"'"
			end if				
	end select
	conn.Execute  "update pet set sinew="&maxsinew&",health=100,happy=100,cleanliness=100,option_T="&nowtimetype&" where id="&pid
	msg=biological&"��"&username&"˵:��ϣ�������������!�������͸��������"&encourage_M&encourage_C
else	
	conn.Execute "update pet set sinew=sinew-"&decrease&",health=health-"&decrease&",happy=happy-"&decrease&",cleanliness=cleanliness-"&decrease&",option_T="&nowtimetype&" where id="&pid
	msg=biological&"��"&username&"˵:ϣ�����컹�ܼ�����!"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<title>����֮��</title>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000ff size=5>����֮��</font><br>�����ĵĺ���,���ǰ��Ľ���</p><hr>
<p align=center>
<%=msg%><br>
<input type=button value='����' onclick=javascript:location.href='pet.asp';>
</p>
</body>
</html>