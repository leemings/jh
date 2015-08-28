<%@ LANGUAGE=VBScript%>
<!--#include file="fun.inc"-->
<%Response.Expires=0
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if session("yx8_mhjh_inchat")<>"1" then %>
<script language="vbscript">
alert "请进入聊天室再进入股票！"
window.close()
</script>
<%Response.End
end if
name=trim(request.form("name"))
jjr=request.form("jjr")
%>
<!--#include file="jhb.asp"-->
<%
sql="select * from 客户 where 帐号='"&name&"'"
set rs=conn.execute(sql)
if rs.eof then
	sql="insert into 客户(帐号,经纪人)values('" & name & "','"&jjr&"')"
	conn.execute sql
	mess="股票帐户申请成功！我们安排了金牌经纪人"&jjr&"为您服务."
else
	sql="update 客户 set 经纪人='"&jjr&"' where 帐号='"&name&"'"
	conn.execute sql
	mess="我们安排了新的经纪人"&jjr&"为您服务."
end if
conn.close
%>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=jingji.asp>返回</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>