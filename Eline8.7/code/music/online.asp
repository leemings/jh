<!--#include file="Config.asp" -->
<%
if UserName="" then
   UserName="客人"
end if

call sqlonline()
Response.Write "document.write(" & chr(34) & "<a href=ShowOnline.asp title=查看在线用户列表><font color=red>在线<strong>"& online() &"</strong>人</font></a>"& chr(34) & ")"

sub sqlonline()
dim statuserid
     statuserid=replace(Request.ServerVariables("REMOTE_HOST"),".","")	
	 Response.Cookies("mesky")("onlineid")=statuserid
	sql="select id from SoftDown_online where id="&cstr(request.cookies("mesky")("onlineid"))
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		sql="insert into "&CategoryName&"_online(id,UserName,ip,startime,lastimebk,browser,actforip) values ("&statuserid&",'"&UserName&"','"&Request.ServerVariables("REMOTE_HOST")&"',now(),now(),'"&Request.ServerVariables("HTTP_USER_AGENT")&"','"&Request.ServerVariables("HTTP_X_FORWARDED_FOR")&"')"
	else
		sql="update "&CategoryName&"_online set lastimebk=now(),UserName='"&UserName&"' where id="&cstr(request.cookies("mesky")("onlineid"))
	end if
	conn.execute(sql)
set rs=nothing
Rem 删除超时用户
sql="Delete FROM "&CategoryName&"_online WHERE DATEDIFF('s', lastimebk, now()) > "&kicktime&"*60"
Conn.Execute sql
end sub



function online()
dim tmprs
	sql="Select count(id) from "&CategoryName&"_online"
set tmprs=conn.execute(sql) 
online=tmprs(0) 
set tmprs=nothing 
if isnull(online) then online=0
end function 

CloseDatabase
%>