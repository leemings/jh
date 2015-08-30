<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then
	Response.Write "<script Language=Javascript>alert('你不是站长？！');window.close();</script>"
	response.end
end if
id=lcase(trim(request("id")))
%>
<!--#include file="data.asp"-->
<%
sql="select * from 密码保护 where 用户="&id
set rs=conn1.execute(sql)
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn1.close
	set conn1=nothing
	Response.Write "<script Language=Javascript>alert('没有找到你要修改的密码保护资料！');window.close();</script>"
	response.end
end if
wt=rs("问题")
da=rs("答案")
rs.close
set rs=nothing
conn1.close
set conn1=nothing
%>
<html>
<head>
<title>密码保护资料修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<bgsound src="pic/bj.mid" loop="-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="pic/WALLPAPER256.JPG">
<div align="center">
  <p><font size="5" color="#0000FF">密码保护资料修改</font></p>
  <form name="form1" method="post" action="mmbhmodi.asp?id=<%=id%>">
    <p>用户ID：<%=ID%> </p>
    <p>提示问题： 
      <input type="text" name="wt" size="10" maxlength="10" value="<%=wt%>">
    </p>
    <p>问题答案： 
      <input type="text" name="da" size="10" maxlength="10" value="<%=da%>">
    </p>
    <p> 
      <input type="submit" name="modify" value="修改">
      <input type=button value=删除 onClick="delmmbh.asp?id=<%=id%>" name="button">
      </a>
    </p>
  </form>
  <p>　</p>
</div>
<p align="center"><font color="#FF0000" size="2"><b>『快乐江湖』之密码保护</b></font></p>
<p align="center"><font color="#FF0000" size="2"><b>程序:回首当年</b></font></p>
</body>
</html>