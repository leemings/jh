<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"�ٸ�" then Response.Redirect "../exit.asp"
username=Request.Form("search")
cz=Request.form("sjcz")
name=request.querystring("username")
if name<>"" then
username=name
cz="����"
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
if cz="ID" then
username=int(username)
sqlstr="SELECT * FROM �û� where "&cz&"="&username
rst.Open sqlstr,conn
else
sqlstr="SELECT * FROM �û� where "&cz&"='"&username&"'"
rst.Open sqlstr,conn
end if
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('��Ǹ����Ҫ���ҵ��������Ҳ�������鿴�Ƿ���ȷ��');history.back();</script>"
else
Response.Write "<form method=post action=updateuser.asp><table><tr><td>id(readonly)</td><td><input type=text readonly name=id value="&rst("id")&"></td></tr>"
for i=1 to rst.Fields.Count-1
Response.Write "<tr><td>"&rst.Fields(i).Name&"("&rst.fields(i).Type&")</td><td><input type=text name="&rst.Fields(i).Name&" value='"&rst.Fields(i).value&"'></td></tr>"
next
%>
 
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=10 background="../chatroom/bg1.gif">
<tr align=center><td colspan=2 align=center width="426"><input type=button onclick="javascript:location.href='laren.asp?uname=<%=username%>';" value=' �� �� '> <input type=button onclick="javascript:location.href='showwubyname.asp?uname=<%=username%>&ucur=����';" value=' �� �� '> <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=ҩƷ';" value=' ҩ Ʒ '> <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=����';" value=' �� �� '>  <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=����';" value=' �� �� ' > <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=ͷ��';" value=' ͷ �� ' > <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=��Ʒ';" value=' �� Ʒ ' > <input type=submit value=' �� �� ' name=submit> <input type=submit value=' ɾ �� ' name=submit> <input type=reset value=' �� �� '> <input type="button" name="ok" value=" �� �� " onclick=javascript:history.go(-1)></td></tr>
<%
Response.Write "</table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
 


