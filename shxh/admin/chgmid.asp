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
rst.Open "select * from mid where id="&id ,conn
dim valuearr(2)
if opt="新增" then
	valuearr(0)=-1
	valuearr(1)=""
	valuearr(2)=""
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
<p align=center>点歌系统</p><hr>
<form action='upmid.asp' method=post id=form1 name=form1>
<table width='80%' align=center border=3>
<tr><td>ID(只读)</td><td><input type=text name=id readonly value=<%=valuearr(0)%>></td></tr>
<tr><td>文件名(文本)</td><td><input type=text name=filename value="<%=trim(valuearr(1))%>" size=20 maxlength=20></td></tr>
<tr><td>歌曲名(文本)</td><td><input type=text name=songname value="<%=trim(valuearr(2))%>"></td></tr>
<tr align=center><td colspan=2><input type=submit name=submit value='<%=opt%>'> <input type=button value='返回' onclick=javascript:location.href='showmid.asp';></td></tr>
</table>
</form>