<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
pid=Request.QueryString("id")
if pid="" or not isnumeric(pid) then Response.Redirect "../error.asp?id=024"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from pet where username='"&username&"' and exist=true and id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=069"
sinew=rst("sinew")
maxsinew=rst("maxsinew")
cleanliness=rst("cleanliness")
jiankang=rst("health")
kuaile=rst("happy")
rst.Close
set rst=nothing
grow=sinew*100\maxsinew+cleanliness+jiankang+kuaile
if grow<320 then Response.Redirect "../error.asp?id=071"
conn.Execute "update pet set sinew="&maxsinew&",health=100,happy=100,cleanliness=100,exist=false,option_M='����',option_T="&nowtimetype&" where id="&pid
conn.Execute "update �û� set ����=����+"&grow&",����=����+"&maxsinew&" where ����='"&username&"'"
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
�����ɹ�.����������<%=maxsinew%>,���¼�<%=grow%>:)<br>
<input type=button value='����' onclick=javascript:location.href='pet.asp';>
</p>
</body>
</html>