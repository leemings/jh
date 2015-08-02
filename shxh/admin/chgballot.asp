<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
id=Request.QueryString("id")
opt=Request.QueryString("opt")
if not isnumeric(id) then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
usergrade=session("Ba_jxqy_usergrade")
usercorp=session("BA_jxqy_usercorp")
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
if username="" then Response.Redirect "../error.asp?id=016"
if not(usergrade=Application("Ba_jxqy_allright") and usercorp="官府") then Response.Redirect "../error.asp?id=046"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from ballotsystem where id="&id ,conn
dim valuearr(6)
if opt="新增" then
	valuearr(0)=-1
	valuearr(1)=""
	valuearr(2)=""
	valuearr(3)=dateadd("d",3,now())
	valuearr(4)="true"
	valuearr(5)="false"
	valuearr(6)=";"
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
<body oncontextmenu=self.event.returnValue=false  background='<%=bgimage%>' bgcolor='<%=bgcolor%>'>
<p align=center>投票系统</p><hr>
<form action='updateballot.asp' method=post id=form1 name=form1>
<ol>关于格式的说明<li>请勿在以下的文本中便用半角的双引号<li>请使用正确的完整的时间格式如：01-4-10 00:00:00<li>条件：true代表所有人都可以参加投票。相反false就没有人能参加投票了，如果你想使用特别的条件，可以使用类sql语句，例：门派='少林寺' and (攻击>1000 or 防御>1000)<li>活跃可设为true或false之一</ul>
<table width='80%' align=center border=3>
<tr><td>ID(只读)</td><td><input type=text name=id readonly value=<%=valuearr(0)%>></td></tr>
<tr><td>标题(文本)</td><td><input type=text name=title value="<%=trim(valuearr(1))%>" size=20 maxlength=20></td></tr>
<tr><td>说明(备注)</td><td><textarea name=explain cols=50 rows=6><%=trim(valuearr(2))%></textarea></td></tr>
<tr><td>截止(时间)</td><td><input type=text name=deadline value=<%=valuearr(3)%> size=20 maxlength=20></td></tr>
<tr><td>条件(文本)</td><td><input type=text name=condition value="<%=valuearr(4)%>" size=50></td></tr>
<tr><td>活跃(逻辑)</td><td><input type=text name=active  value=<%=valuearr(5)%> size=5 maxlength=5></td></tr>
<tr><td>已投票人(备注)</td><td><textarea name=voter cols=50 rows=6><%=trim(valuearr(6))%></textarea></td></tr>
<tr align=center><td colspan=2><input type=submit name=submit value='<%=opt%>'> <input type=button value='返回' onclick=javascript:location.href='showballot.asp';></td></tr>
</table>
</form>