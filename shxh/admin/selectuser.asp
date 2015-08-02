<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if not(session("Ba_jxqy_usercorp")="官府" and Session("Ba_jxqy_usergrade")=Application("Ba_jxqy_allright")) then Response.Redirect "../error.asp?id=500"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
%>
<head>
<link rel="stylesheet" href="../chatroom/css.css">
</head>
<body oncontextmenu=self.event.returnValue=false topmargin=50 background="<%=bgimage%>" bgcolor="<%=bgcolor%>">
<div align=center>
按姓名查询<form action="showuser.asp" method=post name=form1><input type=text name=search maxlength=14 size=14 value=""><INPUT type=submit value=" 查 找 " ></form>
按IP查询<form action="showbyip.asp" method=post><input type=text name=ip maxlength=15 size=15 value=''><input type=submit value=' 查 找 '></form>
按等级查询<form action=showbygr.asp method=post>
<select name=grade>
	<option value=1 selected>1级</option>
	<option value=2>2级</option>
	<option value=3>3级</option>
	<option value=4>4级</option>
	<option value=5>5级</option>
	<option value=6>6级</option>
	<option value=7>7级</option>
	<option value=8>8级</option>
	<option value=9>9级</option>
	<option value=10>10级</option>
</select><input type=submit value='查找'></form>
按门派查询<form action=showbyco.asp method=post><select name=corp>
<option value='无' selected>无门无派</option>
<%
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "门派",conn
do while not rst.EOF
	Response.Write "<option value='"&rst("门派")&"'>"&rst("门派")&"</option>"
	rst.MoveNext
loop
rst.Close
set rst=nothing
conn.Close 
set conn=nothing
%>
</select><input type=submit value='查找'></form>
删除过期帐号<form action=deluser.asp method=post>
<select name=howday>
<option value=7>7天</option>
<option value=15>15天</option>
<option value=30 selected>30天</option>
<option value=90>90天</option>
</select>
<input type=submit value='确认删除'>
</form>
</div>
</body></html>
