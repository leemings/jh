<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
randomize()
rnd0=clng(rnd()*5)+1
rnd1=clng(rnd()*5)+1
rnd2=clng(rnd()*5)+1
if rnd0=rnd1 and rnd0=rnd2 then
	ratio=10
	point=rnd1*3
	grade=3
elseif rnd0=rnd1 then
	ratio=2
	point=rnd2
	grade=2
elseif rnd1=rnd2 then
	ratio=2
	point=rnd0
	grade=2
elseif rnd0=rnd2 then
	ratio=2
	point=rnd1
	grade=2
else
	ratio=0
	point=0
	grade=0	
end if
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from dice where username='"&username&"'",conn,1,2
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
zpoint=rst("zpoint")
if ratio<rst("ratio") then ratio=rst("ratio")
zgrade=rst("zgrade")
frequency=rst("frequency")
bet=rst("bet")
rst.Close
set rst=nothing
if point=0 and frequency<3 then
	conn.Execute "update dice set frequency=frequency+1 where username='"&username&"'"
	msg="ׯ��"&zpoint&"��;��"&point&"��(��"&frequency&"��).<input type=button value='����' onclick=javascript:location.href='pdcontinue.asp';>"
elseif 	zgrade>grade or (zgrade=grade and  zpoint>=point) or frequency>2 then
	conn.Execute "update dice set z0=6,z1=6,z2=6,zgrade=3,zpoint=18,frequency=0,ratio=3,bet=9999 where username='"&username&"'"
	if ratio=10 then 
		conn.Execute "update �û� set ����=����-"&bet*9&" where ����='"&username&"'"
		ratio=11
	end if	
	msg="ׯ��"&zpoint&"��;��"&point&"��(��"&frequency&"��);������˼������"&bet*(ratio-1)&"������<input type=button value='���¿�ʼ' onclick=javascript:location.href='pdbet.asp';  id=button1 name=button1>"	
		
else
	conn.Execute "update dice set z0=6,z1=6,z2=6,zgrade=3,zpoint=18,frequency=0,ratio=3,bet=9999 where username='"&username&"'"
	msg="ׯ��"&zpoint&"��;��"&point&"��(��"&frequency&"��);��ϲ��Ӯ��"&bet*(ratio-1)&"������<input type=button value='���¿�ʼ' onclick=javascript:location.href='pdbet.asp'; >"
	conn.Execute "update �û� set ����=����+"&bet*ratio&" where ����='"&username&"'"
end if  	
conn.Close
set conn=nothing
%>
<html>
<head>
<link rel=stylesheet href='../chatroom/css.css'>
<script language=javascript>
parent.imgfrm.showdice(3,<%=rnd0%>,<%=rnd1%>,<%=rnd2%>);
</script>
</head>
<body  bgcolor='<%=Application("Ba_jxqy_backgroundcolor")%>' background='<%=Application("Ba_jxqy_backgroundimage")%>'>
<p align=center><%=msg%></p>
</html>