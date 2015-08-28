<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=Request.QueryString("un")
if instr(un,"'")<>0 or instr(un," ")<>0 or un="" then Response.Redirect "../error.asp?id=024"
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject ("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select 银两 from 用户 where 姓名='"&username&"' and 银两>499"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=023"
rst.Close
sqlstr="select 状态 from 用户 where 姓名='"&un&"' and 状态='逮捕'"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=025"
rst.Close
set rst=nothing
conn.Execute "update 用户 set 银两=银两-500 where 姓名='"&username&"'"
conn.Execute "update 用户 set 状态='正常',最后登录时间=now() where 姓名='"&un&"'"
conn.Close
set conn=nothing
%><head><title>监狱</title><LINK href="../style.css" rel=stylesheet></head>
<body background='../chatroom/bg1.gif'  oncontextmenu=self.event.returnValue=false>
<p align=center><font color="#CC0000" size="4" face="幼圆"><b>监狱</b></font></p>
<p align=center>我们已经释放了<%=un%>,也谢谢你的500两白花花的银子</p>
<p align=center>
<input type="button" value=" 关 闭 " onClick="javascript:top.window.close();" name="button">
</p>
</body>
