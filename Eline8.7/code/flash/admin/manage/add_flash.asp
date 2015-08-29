<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<html>
<head>
<title>添加动画</title>
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
        <input type="text" name="title" size="50" class="button" maxlength="100">
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画类别：</td>
      <td height="40"> 　 
        <select name="class" class="button">
          <option value="" selected>请选择动画类别</option>
<%
dim rs,sql,sel
set rs=server.createobject("adodb.recordset")
sql="select * from Flashclass"
rs.open sql,conn,1,1
			  do while not rs.eof
		          response.write "<option value='"+CStr(rs("ClassID"))+"' name=class>"+rs("classname")+"</option>"+chr(13)+chr(10)
		          rs.movenext
    		          loop
			  rs.close
			  set rs=nothing
			  conn.close
			  set conn=nothing
			  %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画推荐：</td>
      <td height="40">　 
        <select name="commend" class="button">
          <option selected>请选择推荐等级</option>
          <option value="★">★</option>
          <option value="★★">★★</option>
          <option value="★★★">★★★</option>
          <option value="★★★★">★★★★</option>
          <option value="★★★★★">★★★★★</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画简介：</td>
      <td height="40"> 　 
        <textarea name="content" cols="48" rows="10"></textarea>
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画大小：</td>
      <td height="40"> 　 
        <input type="text" name="KB" size="50" class="button">
        KB </td>
    </tr>
    <tr>
      <td width="120" height="40" align="center">作者信箱：</td>
      <td height="40"> 　
        <input type="text" name="email" size="50" class="button">
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画图片：</td>
      <td height="40"> 　 
        <input type="text" name="img" size="50" class="button">
      </td>
    </tr>
    <tr> 
      <td width="120" height="40" align="center">动画地址：</td>
      <td height="40">　 
        <input type="text" name="url" size="50" class="button">
      </td>
    </tr>
    <tr align="center"> 
      <td colspan="2" height="40"> 
        <input type="submit" name="Submit" value="提交动画" class="button">
        　　 
        <input type="reset" name="reset" value="重新填写" class="button">
      </td>
    </tr>
  </table>
</form>
</body>
</html>
