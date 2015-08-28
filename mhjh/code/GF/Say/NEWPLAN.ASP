<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username=""  then
Response.Redirect "../../error.asp?id=016"
else
if request.form("topic")="" then%>
<script language=vbscript>
MsgBox "请写上状纸的题目！"
location.href = "javascript:history.back()"
</script>
<%
else
%>
<!--#include file="data.asp"-->
<%
name=request.form("name")
topic=request.form("topic")
play=request.form("play")
text=request.form("text")

set rs=server.createobject("adodb.recordset")
sql="select * from bbs where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("被告")=name
rs("标题")=topic
rs("要求")=play
rs("状词")=text
rs("原告")=username
rs.update
rs.close
conn.close
set conn=nothing
response.redirect("nao.asp")
rs.close
set rs=nothing
end if
end if
%>
