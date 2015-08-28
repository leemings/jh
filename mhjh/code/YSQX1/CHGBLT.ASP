<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.QueryString("id")
pid=Request.QueryString("pid")
opt=Request.QueryString("opt")
if not (isnumeric(id) and isnumeric(pid)) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ballot where pid="&pid&" and id="&id ,conn
dim valuearr(3)
if opt="新增" then
valuearr(0)=-1
valuearr(1)=pid
valuearr(2)=""
valuearr(3)=0
else
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
for i=0 to rst.Fields.Count-1
valuearr(i)=rst.Fields(i).Value
next
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif'>
<p align=center>投票系统</p><hr>
<form action='upblt.asp' method=post id=form1 name=form1>
<table width='80%' align=center border=3>
<tr><td>ID(只读)</td><td><input type=text name=id readonly value='<%=valuearr(0)%>'></td></tr>
<tr><td>投票箱(只读)</td><td><input type=text name=pid readonly value='<%=valuearr(1)%>'></td></tr>
<tr><td>选项(文本)</td><td><input type=text name=parti value='<%=Trim(valuearr(2))%>' size=20 maxlength=125></td></tr>
<tr><td>得票(长整)</td><td><input type=text name=vote value='<%=valuearr(3)%>'></textarea></td></tr>
<tr align=center><td colspan=2><input type=submit name=submit value='<%=opt%>'> <input type=button value='返回' onclick=javascript:location.href='showblt.asp?pid=<%=pid%>'; ></td></tr>
</table>
</form>



