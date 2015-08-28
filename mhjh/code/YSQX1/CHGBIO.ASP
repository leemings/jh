<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.QueryString("id")
opt=Request.QueryString("opt")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
adminname=Application("yx8_mhjh_admin")
if username<>adminname then Response.Redirect "../error.asp?id=046"
if session("jxqy_adminbok")="" or session("yx8_mhjh_usercorp")<>"官府" then Response.Redirect "../exit.asp"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from biological where id="&id ,conn
dim valuearr(6)
if opt="新增" then
valuearr(0)=-1
valuearr(1)=""
valuearr(2)=""
valuearr(3)="10000"
valuearr(4)="10000"
valuearr(5)="10000"
valuearr(6)="银两 10"
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
<p align=center>怪物管理</p><hr>
奖励可使用以下格式：<br>
银两 数量 例：银两 10 击败该怪物得到银两10<br>
积分 数量 例：积分 10 击败该怪物得到积分10<br>
物品 物品名 例：物品 天意 击败该怪物行到物品天意，请确认物品名存在于武器店或药店中<br>
<form action='upbio.asp' method=post id=form1 name=form1>
<table width='80%' align=center border=3>
<tr><td>ID(只读)</td><td><input type=text name=id readonly value=<%=valuearr(0)%>></td></tr>
<tr><td>怪物(文本)</td><td><input type=text name=biological value="<%=trim(valuearr(1))%>" size=40 maxlength=80></td></tr>
<tr><td>图片(文本)</td><td><input type=text name=picture value="<%=trim(valuearr(2))%>" size=40 maxlength=80></td></tr>
<tr><td>体力(长整)</td><td><input type=text name=hp value="<%=trim(valuearr(3))%>" size=40 maxlength=80></td></tr>
<tr><td>攻击(长整)</td><td><input type=text name=attack value="<%=trim(valuearr(4))%>" size=40 maxlength=80></td></tr>
<tr><td>防御(长整)</td><td><input type=text name=defence value="<%=trim(valuearr(5))%>" size=40 maxlength=80></td></tr>
<tr><td>奖励(文本)</td><td><input type=text name=encourage value="<%=trim(valuearr(6))%>" size=40 maxlength=80></td></tr>
<tr align=center><td colspan=2><input type=submit name=submit value='<%=opt%>'> <input type=button value='返回' onclick=javascript:location.href='showbio.asp';></td></tr>
</table>
</form>
