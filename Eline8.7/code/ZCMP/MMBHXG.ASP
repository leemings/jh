<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select id,门派,性别 from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你不是江湖中人，不能申请密码保护？！');window.close();</script>"
	response.end
end if
sjjh_id=rs("id")
menpai=rs("门派")
sex=rs("性别")
rs.close
conn.close
set conn=nothing
if sex="男" then
	sex="大侠"
else
	sex="女侠"
end if
if menpai="游侠" or menpai="无" or menpai="" then
	menpai=""
end if
%><!--#include file="data.asp"--><%
sql="select * from 密码保护 where 用户="&sjjh_id
set rs=conn1.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('你还未申请过密码保护，为了您的数据安全，请先申请密码保护！');location.href = 'mmbh1.asp';</script>"
	response.end
end if
rs.close
set rs=nothing
conn1.close
set conn1=nothing
%>
<html>
<head>
<title>修改密码保护</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<bgsound src="pic/bj.mid" loop="-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="pic/WALLPAPER256.JPG">
<form name="form1" method="post" action="mmbhxg1.asp">
  <p align="center"><font size="5"><b><font color="#0000FF">修改密码保护资料</font></b></font></p>
  <p align="center"><font color="#FF0000" size="2">第一步：</font><font color="#0000FF" size="2">验证身份</font></p>
  <p align="center"><font size="2" color="#800040"><%=menpai%></font><font size="2" color="#FF0000"><%=sex%></font><font size="2" color="#8000FF"><%=sjjh_name%></font><font size="2">，您的ID是：<font color="#FF0000"><%=sjjh_id%></font></font></p>
  <p align="center"><font size="2" color="#0000FF">您的密码提示问题：</font><font size="2"> 
    <input type="text" name="wt" size="12" maxlength="10">
    </font></p>
  <p align="center"><font size="2" color="#0000FF">您的提示问题答案：</font><font size="2"> 
    <input type="text" name="da" size="12" maxlength="10">
    </font></p>
  <p align="center"> 
    <input type="submit" name="Submit" value="下一步">
    <input type="submit" name="Submit2" value="清除">
  </p>
</form>
<p>　</p>
<p align="center"><font color="#FF0000" size="2"><b>『快乐江湖』之密码保护</b></font></p>
<p align="center"><font color="#FF0000" size="2"><b>程序修改:回首当年</b></font></p>
</body>
</html>