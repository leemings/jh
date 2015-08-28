<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
cid=server.HTMLEncode(trim(Request.QueryString("cid")))
opt=server.HTMLEncode(trim(Request.QueryString("opt")))
if not isnumeric(cid) then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from cardshop where id="&cid,conn
if opt="新增" then
cid=-1
cname=""
cespecial="疯狂"
ctime=60
cprice=10000
else
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=100"
cid=rst("id")
cname=Trim(rst("cname"))
cespecial=Trim(rst("cespecial"))
ctime=rst("ctime")
cprice=rst("cprice")
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
esopt=array("疯狂","封咒","中毒","麻痹","驱逐","逮捕","入狱","炸弹","死亡")
for i=0 to ubound(esopt)
if esopt(i)=cespecial then
msg=msg&"<option value='"&esopt(i)&"' selected>"&esopt(i)&"</option>"
else
msg=msg&"<option value='"&esopt(i)&"'>"&esopt(i)&"</option>"
end if
next
%>
<html>
<head>
<link rel='stylesheet' href='../chatroom/css.css'>
</head>
<body oncontextmenu=self.event.returnValue=false  background='../chatroom/bg1.gif'>
'
<p align=center><font color=0000ff size=4>道具管理</font></p><hr>
<form action=upcard2.asp method=post>
<table width=80% align=center border=3>
<tr><td><input type=hidden value='<%=cid%>' name=id>道具名称</td><td><input type=text name=cname value='<%=cname%>' size=14 maxlength=7></td></tr>
<tr><td>作用</td><td><select name=especial><%=msg%></select></td></tr>
<tr><td>时限</td><td><input type=text name=ctime value='<%=ctime%>' size=3 maxlength=3></td></tr>
<tr><td>价格</td><td><input type=text name=cprice value='<%=cprice%>' size=8 maxlength=8></td></tr>
<tr align=center><td colspan=2><input type=submit name=opt value='<%=opt%>'> <input type=button value='返回' onclick='javascript:history.back();'></td></tr>
</table>
</form>
</body>
</html>

