<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!-- #include file="conn.asp" -->
<%
Session.TimeOut=30
'判断用户填交的表单
Dim username,username_mail,password,list,strSQL
If request.form("username") = "" Then
%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>资料修改</title>
<style>
<!--
td           { font-family: 宋体; font-size: 9pt; color: #000000 }
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style></head><body topmargin="0" leftmargin="0">
<p align="center"><p align="center"><p><a href=./>返回交友列表</a>　</p>
<form method="POST" action="userre.asp"><div align="center"><center>  <table border="1" width="600" cellspacing="0" bordercolor="#988CD8">
<tr><td colspan="5" bgcolor="#988CD0">
<p align="center"><font color="#FFFF80" size="3"><b>修 改 用 户 信 息</b></font></p> 
</td> </tr> </center> <tr> <td>
          <p align="right">你的姓名：</p>
</td><center>  <td><p align="center"><input maxLength="20" name="username" size="15" tabIndex="1"></p>
</td></center><td><p align="right">你的密码：</p></td><center>  
<td><p align="center"><font color="#0080ff"><input maxLength="20" name="password" size="20" tabIndex="2" type="password" value></font></p>
</td><td><p align="center"><input tabIndex="10" type="submit" value="   进入   "></p>
</td></tr></table></center></div></form><p>　</p><p>　</p><p>　</p>
<form method="POST" action="userre.asp"><div align="center">
<center><table border="1" width="600" cellspacing="0" bordercolor="#988CD8"><tr><td colspan="4" bgcolor="#988CD0"><p align="center"><font color="#FFFF80" size="3"><b>密 码 查 询</b></font></td> 
</tr> <tr> <td> 
            <p align="center">你的姓名：</p>
          </td><td><p align="center"><input maxLength="20" name="username_mail" size="14" tabIndex="3"></p></td>
<td><p align="center">你的密码会自动发送到你申请时使用的信箱中</p></td><td><p align="center"><input tabIndex="10" type="submit" value="   确定   "></p></td>
</tr></table></center></div></form><center></center>    <p>　</p><p>　</p>
<p> 
  <center>
    <font style="color: #000000; font-family: 宋体; font-size: 9pt"> Copyright &copy; 
    2001-2003 cidu.net All rights reserved<br>
    版权所有：<a href="http://www.21ex.com" target="_blank">游戏吧</a> 　网站维护：<a href="mailto:21ex@21cn.com" target="_blank">剑侠情缘</a> 
    </font> 
  </center>
</p>
</body>   </html>      

<%Else
username=Request.form("username")
password=request.form("password")
If request.form("username_mail") <> "" Then
'发邮件给用户，先需要索检用户的名称和MAIL，然后将密码发送给用户

End If
'检测用户名称和密码
If request.form("username")="" or request.form("password")="" Then
 response.write "<script language='javascript'>" & chr(13)
Response.Write "错误！请重新正确输入！"
 response.write "window.document.location.href='userre.asp';"&Chr(13)
 response.write "</script>" & Chr(13)
Response.End
Else	
		strSQL="select * from list where name='"+Request.form("username")+"'"
		'索检用户的名称，然后对照密码是否正确
		  set list=server.createobject("adodb.recordset")

    list.open strSQL,conn,adOpenKeySet,adLockPessimistic
         if list.eof and list.bof then
         %>
            <div align="center"><center>
            <table border="1" width="650" cellspacing="0" cellpadding="3" bordercolor="#988CD0" bgcolor="#E0E0E0" style="font-family: 宋体; font-size: 9pt; color: #000000">
            <tr><td width="100%">
            <p align="center"><font color="#FF0000">
            Sorry! 没有你想找的资料！</font></p> 
            </td></tr></table></center></div> 
          <%
        response.write "<script language='javascript'>" & chr(13)
        response.write "alert('没有该用户！请注册.');" & Chr(13)
        response.write "window.document.location.href='add.asp';"&Chr(13)
        response.write "</script>" & Chr(13)
          response.end
          ENd If


        If request.form("password")<>list("password") or list("password")="" or request.form("password")="" Then
        Response.Write "密码错误！<a href = javascript:history.back()>返回重输</a>"
        Else
        Session("username") =list("name")
        Session("password") =list("password")
        response.write "<script language='javascript'>" & chr(13)
        response.write "window.document.location.href='moduser.asp?UserID="&list("id")&"';"&Chr(13)
        response.write "</script>" & Chr(13)
        End If
list.close
End if
ENd if
%>