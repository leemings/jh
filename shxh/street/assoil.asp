<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
un=Request.QueryString("un")
if instr(un,"'")<>0 or instr(un," ")<>0 or un="" then Response.Redirect "../error.asp?id=024"
username=session("Ba_jxqy_username")
mygrade=session("Ba_jxqy_usergrade")
if username="" then Response.Redirect "../error.asp?id=016"
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject ("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select * from 用户 where 姓名='"&un&"' and 状态 in('逮捕','入狱')"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=025"
state=rst("状态")
if (state="逮捕" and mygrade>=Application("Ba_jxqy_arrestright")) or (state="入狱" and mygrade>=Application("Ba_jxqy_gaolright")) then
	conn.Execute "update 用户 set 状态='正常',最后登录时间=now() where 姓名='"&un&"'"
else
	Response.Redirect "../error.asp?id=026"
end if
rst.Close
set rst=nothing
conn.Close 
set conn=nothing
aberrantlist=Application("Ba_jxqy_aberrantlist")
aberrantlistubd=ubound(aberrantlist)
dim newaberrantlist()
newaberrantname=";"
j=1
for i=1 to aberrantlistubd step 4
	if not (onlinelist(i)=un and instr("逮捕;入狱",onlinelist(i+2))<>0) then
		redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
		newaberrantlist(j)=aberrantlist(i)
		newaberrantlist(j+1)=aberrantlist(i+1)
		newaberrantlist(j+2)=aberrantlist(i+2)
		newaberrantlist(j+3)=aberrantlist(i+3)
		j=j+4
		newaberrantname=newaberrantname&aberrantlist(i)&";"
	end if
next
Application.Lock
if j=1 then
	dim index(0)
	Application("Ba_jxqy_aberrantlist")=index
else	
	Application("Ba_jxqy_aberrantlist")=newaberrantlist
end if	
Application("Ba_jxqy_aberrantname")=newaberrantname
Application.UnLock
%>
<head><title><%=Application("Ba_jxqy_systemname")%></title><link rel='stylesheet' href='../chatroom/css.css'></head>
<body bgcolor='<%=bgcolor%>' background='<%=bgimage%>' oncontextmenu=self.event.returnValue=false>
<p align=center><font color=CC0000 size=4 face=幼圆><b><%=Application("Ba_jxqy_systemname")%>监狱</b></font></p>
<p align=center>你放心我们马上释放<%=un%>！</p>
<p align=center>
  <input type="button" value=" 关 闭 " onClick="javascript:top.window.close();" name="button">
</p>
</body>