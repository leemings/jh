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
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../21ex/mj.asp")
rst.Open "select * from 商店 where id="&id ,conn
dim valuearr(11)
if opt="新增" then
valuearr(0)=-1
valuearr(1)=""
valuearr(2)=""
valuearr(3)=""
valuearr(4)=""
valuearr(5)=""
valuearr(6)=""
valuearr(7)=""
valuearr(8)=""
valuearr(9)=""
valuearr(10)=""
valuearr(11)=""
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
<form action='upsl.asp' method=post id=form1 name=form1>
<table width='80%' align=center border=3>
<tr><td>ID(只读)</td><td>
  <input type=text name=id readonly value=<%=valuearr(0)%> size="20"></td></tr>
<tr><td>物品名(文本)</td><td><input type=text name=物品名 value="<%=trim(valuearr(1))%>" size=20 maxlength=20></td></tr>
<tr><td>类型(文本)</td><td><input type=text name=类型 value="<%=trim(valuearr(2))%>" size=20 maxlength=20></td></tr>
<tr><td>说明(文本)</td><td><input type=text name=说明 value="<%=trim(valuearr(3))%>" size=20 maxlength=20></td></tr>
<tr><td>功能(文本)</td><td><input type=text name=功能 value="<%=trim(valuearr(4))%>" size=20 maxlength=20></td></tr>
<tr><td>精力(文本)</td><td><input type=text name=精力 value="<%=trim(valuearr(5))%>" size=20 maxlength=20></td></tr>
<tr><td>攻击(文本)</td><td><input type=text name=攻击 value="<%=trim(valuearr(6))%>" size=20 maxlength=20></td></tr>
<tr><td>防御(文本)</td><td><input type=text name=防御 value="<%=trim(valuearr(7))%>" size=20 maxlength=20></td></tr>
<tr><td>特效(文本)</td><td><input type=text name=特效 value="<%=trim(valuearr(8))%>" size=20 maxlength=20></td></tr>
<tr><td>会员(文本)</td><td><input type=text name=会员 value="<%=trim(valuearr(9))%>" size=20 maxlength=20></td></tr>
<tr><td>银两(数字)</td><td><input type=text name=银两 value="<%=trim(valuearr(10))%>" size=20 maxlength=20></td></tr>
<tr align=center><td colspan=2><input type=submit name=submit value='<%=opt%>'> <input type=button value='返回' onclick=javascript:location.href='showsl.asp';></td></tr>
</table>
</form>