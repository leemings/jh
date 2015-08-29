<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->

<%
ID=request("ID")
set rs1=server.createobject("adodb.recordset")
sql1="select * from Flash where id="&request("id")
rs1.open sql1,conn,1,1
%>
<html>
<head>
<title>修改动画</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="author" content="Thirdsnow;Email:qb169@sohu.com;QQ:34608582">
<link rel="stylesheet" href="../style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="bkg.gif">
<form method="post" action="save_flash.asp" name="flash">
  <table width="500" border="0" cellspacing="1" cellpadding="0" align="center">
    <tr> 
      <td width="120" height="40" align="center">动画标题：</td>
      <td height="40"> 　 
        <input type="text" name="title" size="50" class="button" maxlength="100" value="<%=rs1("title")%>">
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画类别：</td>
      <td height="40"> 　 
        <select name="class" class="button">
          <option value="" selected>请选择动画类别</option><%
	dim rs,sql,sel
	set rs=server.createobject("adodb.recordset")
	sql="select * from FlashClass"
	rs.open sql,conn,1,1
	do while not rs.eof
%>
<option value=<%=rs("ClassID")%> <% if rs1("classid")=rs("classid") then %> selected <%end if%> name=class><%=rs("classname")%></option>
<%	          
	rs.movenext
    	loop
	rs.close
	set rs=nothing
%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画推荐：</td>
      <td height="40">　 
        <select name="commend" class="button">
          <option selected>请选择推荐等级</option>
          <option value="★" <% if rs1("commend")="★" then %>selected <%end if%>>★</option>
          <option value="★★" <% if rs1("commend")="★★" then %>selected <%end if%>>★★</option>
          <option value="★★★" <% if rs1("commend")="★★★" then %>selected <%end if%>>★★★</option>
          <option value="★★★★" <% if rs1("commend")="★★★★" then %>selected <%end if%>>★★★★</option>
          <option value="★★★★★" <% if rs1("commend")="★★★★★" then %>selected <%end if%>>★★★★★</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画简介：</td>
      <td height="40"> 　 
        <textarea name="content" cols="48" rows="10" class="text"><%=rs1("content")%></textarea>
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画大小：</td>
      <td height="40"> 　 
        <input type="text" name="KB" size="50" class="button" value="<%=rs1("kb")%>">
        KB </td>
    </tr>
    <tr>
      <td width="120" height="40" align="center">作者信箱：</td>
      <td height="40"> 　
        <input type="text" name="email" size="50" class="button" value="<%=rs1("email")%>">
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画图片：</td>
      <td height="40"> 　 
        <input type="text" name="img" size="50" class="button" value="<%=rs1("img")%>">
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画地址：</td>
      <td height="40">　 
        <input type="text" name="url" size="50" class="button" value="<%=rs1("url")%>">
      </td>
    </tr>
    <tr align="center"> 
      <td colspan="2" height="40"> 
	<input type="hidden" name="id" value="<%=rs1("id")%>">
	<input type="hidden" name="action" value="edit">
        <input type="submit" name="Submit" value="提交动画" class="button">
        　　 
        <input type="reset" name="reset" value="重新填写" class="button">
      </td>
    </tr>
  </table>
</form>
<%
rs1.close
set rs1=nothing
conn.close
set conn=nothing
%>
</body>
</html>
