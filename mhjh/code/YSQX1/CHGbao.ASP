<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=0
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
id=Request.QueryString("id")
opt=Request.QueryString("opt")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
set conn=server.CreateObject("adodb.connection")
set rst=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/jihe.asp")
rst.Open "select * from 宝物 where id="&id ,conn
dim valuearr(7)
if opt="新增" then
valuearr(0)=-1
valuearr(1)=""
valuearr(2)=""
valuearr(3)=""
valuearr(4)=""
valuearr(5)=""
valuearr(6)=""
valuearr(7)=""
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
<p align=center>　</p><hr>
<form action='upbao.asp' method=post id=form1 name=form1>
<table width='80%' align=center border=3>
<tr><td>ID(只读)</td><td>
  <input type=text name=id readonly value=<%=valuearr(0)%> size="20"></td></tr>
<tr><td>宝物名(文本)</td><td><input type=text name=宝物名 value="<%=trim(valuearr(1))%>" size=20 maxlength=20></td></tr>
<tr><td>攻击力(数字)</td><td><input type=text name=攻击力 value="<%=trim(valuearr(2))%>" size=20 maxlength=20></td></tr>
<tr><td>拿走(文本)</td><td><input type=text name=拿走 value="<%=trim(valuearr(3))%>" size=20 maxlength=20></td></tr>
<tr><td>获宝者(文本)</td><td><input type=text name=获宝者 value="<%=trim(valuearr(4))%>" size=20 maxlength=20></td></tr>
<tr><td>时间(文本)</td><td><input type=text name=时间 value="<%=trim(valuearr(5))%>" size=20 maxlength=20></td></tr>
<tr><td>防御力(数字)</td><td><input type=text name=防御力 value="<%=trim(valuearr(6))%>" size=20 maxlength=20></td></tr>
<tr><td>价格(数字)</td><td><input type=text name=价格 value="<%=trim(valuearr(7))%>" size=20 maxlength=20></td></tr>
<tr align=center><td colspan=2><input type=submit name=submit value='<%=opt%>'> <input type=button value='返回' onclick=javascript:location.href='showbao.asp';></td></tr>
</table>
</form>