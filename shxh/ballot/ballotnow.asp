<%
Response.Expires=-1
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=016"
id=Request.Form("id")
pid=Request.Form("pid")
if id="" or pid="" then Response.Redirect "../error.asp?id=048"
if not (isnumeric(id) and isnumeric(pid)) then	Response.Redirect "../error.asp?id=024"
nowtime=now()
set conn=server.CreateObject("adodb.connection")
conn.Open Application("Ba_jxqy_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select deadline,condition,voter from ballotsystem where id="&pid,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=024"
deadline=cdate(rst("deadline"))
condition=Trim(rst("condition"))
voter=trim(rst("voter"))
rst.Close
rst.Open "select 体力 from 用户 where 姓名='"&username&"' and "&condition,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=049"
rst.Close 
set rst=nothing
if instr(voter,";"&username&";") then Response.Redirect "../error.asp?id=072"
if deadline<nowtime then Response.Redirect "../error.asp?id=073"	
conn.Execute "update ballot set vote=vote+1 where id="&id&" and pid="&pid
conn.Execute "update ballotsystem set voter='"&voter&username&";' where id="&pid
conn.Close
set conn=nothing
Response.Redirect "showballot.asp?pid="&pid
%>