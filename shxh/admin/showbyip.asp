<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
usercorp=session("Ba_jxqy_usercorp")
usergrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="�ٸ�") then Response.Redirect "../error.asp?id=046"
ip=Request("ip")
if ip="" then Response.Redirect "../error.asp?id=060"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=cint(acpage)
if acpage<1 then acpage=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open application("Ba_jxqy_connstr")
Set rst=Server.CreateObject("ADODB.RecordSet")
rst.Open "SELECT * FROM �û� where ����¼ip='"&ip&"'",conn,1,3
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
Response.Write "<head><link rel='stylesheet' href='../chatroom/css.css'></head><body oncontextmenu=self.event.returnValue=false topmargin=0 background='"&bgimage&"' bgcolor='"&bgcolor&"' text='FF0000'><p align=center><h4>��IP��ѯ:<font color=FF0000>"&ip&"</font><h4></p><hr><table width=100% align=center border=3><tr bgcolor=FFFF00 align=center><td>����</td><td>����</td><td>����</td><td>�ȼ�</td><td>����</td><td>����</td><td>����</td><td>����</td><td>����</td><td>���</td></tr>"
rst.PageSize=20
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
for j=1 to rst.Pagesize
	if rst.EOF or rst.BOF then
		exit for
	else
		Response.Write "<tr><td><a href='showuser.asp?search="&rst("����")&"'>"&rst("����")&"</a></td><td>"&rst("����")&"</td><td>"&rst("����")&"&nbsp;</td><td>"&rst("�ȼ�")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("����")&"</td><td>"&rst("���")&"</td></tr>"
		rst.MoveNext
	end if
next	
Response.Write "</table><form action='showbyip.asp?ip="&ip&"' method=post id=form1 name=form1><table border=1 width=100% bgcolor=cccccc><tr align=center>"
if acpage>1 then
	response.write "<td><a href='showbyip.asp?ip="&ip&"&acpage=1' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='showbyip.asp?ip="&ip&"&acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='ǰһҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">ǰһҳ</a></td>"
else
	response.write "<td>��һҳ</td><td>ǰһҳ</td>"
end if
if acpage<rst.PageCount then
	response.write "<td><a href='showbyip.asp?ip="&ip&"&acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='��һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">��һҳ</a></td><td><a href='showbyip.asp?ip="&ip&"&acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='���һҳ';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">���һҳ</a></td>"
else
	response.write "<td>��һҳ</td><td>���һҳ</td>"
end if
response.write "<td>��<input type=text name=acpage size=4 value='"&acpage&"'>ҳ/��"&rst.PageCount&"ҳ<input type=submit value='GO' id=submit1 name=submit1></td></tr></table></form></body>"
%>