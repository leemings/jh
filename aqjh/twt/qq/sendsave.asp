<%
if session("aqjh_name")="" then response.redirect"../../error.asp?id=210"
if request("oicq")="" then response.redirect"send.asp?id=2"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select * from QQ where oicq="&request("oicq")
rs.open sql,conn,3,2
if rs.eof or rs.bof then
rs.addnew
rs("oicq")=request("oicq")
rs("日期")=now()
rs("用户名")=session("aqjh_name")
rs("奖励")=0
rs("处罚")=0
rs.update
rs.close
response.redirect "qqlist.asp"
else
%>
<html>
<head>
<title>奖励申请</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head><body bgcolor="#FFFFFF">
<div align="center">
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>您填写的OICQ号码已经存在</p>
  <p><a href="send.asp">返回</a></p>
</div>
</body>
</html>
<%end if%>