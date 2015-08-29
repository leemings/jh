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
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你不是江湖中人，不能申请密码保护？！');window.close();</script>"
	response.end
end if
sjjh_id=rs("id")
rs.close
conn.close
set conn=nothing
%><!--#include file="data.asp"--><%
sql="select * from 密码保护 where 用户="&sjjh_id
set rs=conn1.execute(sql)
if not(rs.eof or rs.bof) then
	rs.close
	set rs=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('你已申请过密码保护了，不可以再申请！');window.close();</script>"
	response.end
end if
rs.close
set rs=nothing
conn1.close
set conn1=nothing
%>
<html>
<head>
<title>密码保护申请</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<bgsound src="pic/bj.mid" loop="-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="pic/WALLPAPER256.JPG">
<table width="48%" border="0" cellspacing="0" cellpadding="0" height="113">
  <tr>
    <td height="55"> 
      <div align="left"></div>
      <form name="form1" method="post" width=300 height=109 action="mmbh2.asp">
        <div align="left" style="width: 661; height: 214">
        <p align="left"><font color="#0000FF" size="4"><b><font color="#FF0000">&nbsp;申请密码保护</font></b></font></p>
        <p align="left"><font color="#0000FF" size="2">您的<b><font color="#FF0000">ID</font></b>号为</font><font size="2">：<%=sjjh_id%>，<font color="#0000FF">姓名</font>：<%=sjjh_name%></font></p>
        <p align="left"><font color="#800040" size="2">请认真填写以下内容，并请牢记，</font></p>
        <p align="left"><font color="#800040" size="2">这是你找回丢失密码的唯一办法。</font></p>
          <p align="left"><font color="#0000FF" size="2">提示问题：</font> 
            <input type="text" name="wt" maxlength="10" size="12">
          </p>
          <p align="left"><font color="#0000FF" size="2">问题答案：</font> 
            <input type="text" name="da" maxlength="10" size="12">
          </p>
          <p align="left"> <font color="#FF0000" size="2"><b>(以上内容填写只可以使用汉字，最长不得超过五个字)</b></font></p>  
          <p align="left"> 
            <input type="submit" name="Submit" value="申  请">
          </p>
  </div>
</form>
</td>
  </tr>
</table>
</body>
</html>