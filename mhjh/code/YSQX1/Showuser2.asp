<%
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
username=Request.Form("search")
cz=Request.form("sjcz")
name=request.querystring("username")
if name<>"" then
username=name
cz="姓名"
end if
set conn=server.createobject("adodb.connection") 
conn.open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
if cz="ID" then
username=int(username)
sqlstr="SELECT * FROM 用户 where "&cz&"="&username
rst.Open sqlstr,conn
else
sqlstr="SELECT * FROM 用户 where "&cz&"='"&username&"'"
rst.Open sqlstr,conn
end if
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要查找的人我们找不到！请查看是否正确！');history.back();</script>"
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
<tr align=center><td colspan=2 align=center width="426"><input type=button onclick="javascript:location.href='laren.asp?uname=<%=username%>';" value=' 拉 人 '> <input type=button onclick="javascript:location.href='showwubyname.asp?uname=<%=username%>&ucur=攻击';" value=' 武 功 '> <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=药品';" value=' 药 品 '> <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=武器';" value=' 武 器 '>  <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=防具';" value=' 防 具 ' > <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=头盔';" value=' 头 盔 ' > <input type=button onclick="javascript:location.href='showcurbyname.asp?uname=<%=username%>&ucur=饰品';" value=' 饰 品 ' > <input type=submit value=' 更 新 ' name=submit> <input type=submit value=' 删 除 ' name=submit> <input type=reset value=' 重 置 '> <input type="button" name="ok" value=" 返 回 " onclick=javascript:history.go(-1)></td></tr>
<%
Response.Write "</table></form>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
 


