<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("yx8_mhjh_username")="" then Response.Redirect "../../error.asp?id=016"
username=Request.Form("search")
cz=Request.form("sjcz")
name=request.querystring("username")
if name<>"" then
username=name
cz="姓名"
end if
%>
<!--#include file="data.asp"-->
<%
set rst=server.CreateObject ("adodb.recordset")
if cz="歌名" then
sqlstr="SELECT * FROM 歌曲 where 歌名='"&username&"'"
rst.Open sqlstr,conn
else
sqlstr="SELECT * FROM 歌曲 where 歌手='"&username&"'"
rst.Open sqlstr,conn
end if
if rst.EOF or rst.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要查找的资料我们找不到！');history.back();</script>"
else
Response.Redirect "show.asp?corp="&cz&"&corpa="&username&""
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
