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
%><!--#include file="data.asp"--><%
sql="select * from 密码保护 order by 用户"
set rs=conn1.execute(sql)
%>
<html>
<head>
<title>『快乐江湖』密码保护资料查询</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="pic/WALLPAPER256.JPG">
<div align="center">
  <p><font size="5" color="#FF0000"><b>密码保护资料查询</b></font></p>
  <p>&nbsp; </p>
  <table width="56%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#800000">
    <tr align="center" valign="middle"> 
      <td height="27" width="18%"> 
        <div align="center"><font color="#0000FF" size="2">用户ID号</font></div>
      </td>
      <td height="27" width="28%"> 
        <div align="center"><font color="#0000FF" size="2">提示问题</font></div>
      </td>
      <td height="27" width="26%"> 
        <div align="center"><font color="#0000FF" size="2">答案</font></div>
      </td>
      <td height="27" width="28%"> 
        <div align="center"><font color="#0000FF" size="2">操作</font></div>
      </td>
    </tr>
    <%do while not rs.bof and not rs.eof
%>
    <tr align="center" valign="middle"> 
      <td height="22" width="18%"><font size="2" color="#0000FF"><%=rs("用户")%></font></td>
      <td height="22" width="28%"><font size="2" color="#0000FF"><%=rs("问题")%></font></td>
      <td height="22" width="26%"><font size="2" color="#0000FF"><%=rs("答案")%></font></td>
      <td height="22" width="28%"> 
        <div align="center"><font color="#0000FF" size="2"><a href="delmmbh.asp?id=<%=rs("用户")%>">删除</a> 
          | <a href="modmmbh.asp?id=<%=rs("用户")%>">修改</a>   
          </font></div>
      </td>
    </tr>
    <%rs.movenext
  loop
  rs.close
  set rs=nothing
  conn1.close
  set conn1=nothing
  %>
  </table>
</div>
<p align="center"><font color="#FF0000" size="2"><b>『快乐江湖』之密码保护</b></font></p>
<p align="center"><font color="#FF0000" size="2"><b>程序修改:回首当年</b></font></p>
</body>
</html>