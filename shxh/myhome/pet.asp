<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from pet where username='"&username&"' order by option_t desc",conn,1,2
if not( rst.EOF or rst.BOF) then
	rst.PageSize=20
	if acpage>rst.PageCount then acpage=rst.PageCount
	rst.AbsolutePage=acpage
	for i=1 to rst.PageSize
		if rst.EOF or rst.BOF then exit for
		if rst("option_T")>nowtime then
			opt=rst("option_M")&"��"&rst("option_T")
		elseif rst("exist")=True then
			opt="<a href='attend.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='�Ҷ���,������,�Ҿ���!';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">�չ�</a>"
		else
			opt="<a href='visitpet.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='�þò�����,��������ѽ!';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">̽��</a>"
		end if
		msg=msg&"<tr><td>"&rst("biological")&"</td><td>"&rst("sinew")&"/"&rst("maxsinew")&"</td><td>"&rst("health")&"</td><td>"&rst("cleanliness")&"</td><td>"&rst("happy")&"</td><td>"&opt&"</td></tr>"
		rst.MoveNext
	next
	msg=msg&"</table><table width=100% border=3 bgcolor=ffff00><form action='pet.asp' method=post id=form1 name=form1>"
	if acpage>1 then
		msg=msg&"<td><a href='pet.asp?acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='pet.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>ǰһҳ</td>"
	end if
	if acpage<rst.PageCount then
		msg=msg&"<td><a href='pet.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='pet.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
	else
		msg=msg&"<td>��һҳ</td><td>���һҳ</td>"
	end if
	msg=msg&"<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ/��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></form>"
else
	msg="<tr><td colspan=5>��û�г���ɹ��չ�</td></tr>"
end if
%>
<html>
<head>
<title>����֮��</title>
<link rel=stylesheet href='../chatroom/css.css'>
</head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=0000ff size=5>����֮��</font><br>�����ĵĺ���,���ǰ��Ľ���</p><hr>�������Խ�Ϊ���ϴο���֮ǰ���������,��֪������������ô����?ȥ�������ǰ�!
<table width=80% align=center border=3>
<tr align=center bgcolor=ffff00><td>����</td><td>����</td><td>����</td><td>���</td><td>����</td><td>����</td></tr>
<%=msg%>
</table>
<p align=center><input type=button value='����' onclick=javascript:location.href='myhome.asp';></p>
</head>
</html>