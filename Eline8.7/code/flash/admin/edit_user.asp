<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="level.asp"-->
<%
dim ID
ID=request("ID")
set rs=server.createobject("adodb.recordset")
sql="select * from admin where ID="&request("ID")
rs.open sql,conn,1,1
%>
<html>
<head>
<title>修改用户</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/style.css" type="text/css">
<Script Language=Javascript>
function check(){
var f=document.add
if(f.username.value.length==0){alert("用户名不能为空");f.username.focus();return false}
if(f.password.value.length==0){alert("用户密码不能为空");f.password.focus();return false}
}
function over(locate){
locate.style.border='1px solid black';locate.style.backgroundColor='#FFFFFF';locate.style.cursor='hand'
}
function out(locate){
locate.style.border='';locate.style.backgroundColor='#F7F7F7'
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" background="bkg.gif">
<form name="add" method="post" action="save_user.asp" onsubmit="return check()">
        
  <table width="40%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
            <td height="27" width="120" align="center">用户姓名：</td>
            
      <td height="27" width="209"> 
        <input type="text" name="username" size="20" maxlength="20" class="form" value="<%=rs("username")%>">
            </td>
          </tr>
          <tr> 
            <td height="27" width="120" align="center">用户密码：</td>
            
      <td height="27" width="209"> 
        <input type="password" name="password" size="20" maxlength="20" class="form" value="<%=rs("password")%>">
            </td>
          </tr>
          <tr> 
            <td height="27" width="120" align="center">用户等级：</td>
            
      <td height="27" width="209"> 
        <select name="level" class="form">
                <option value="">请选择用户等级</option>
                <option value="全部" <% if rs("Level")="全部" then%>selected<%end if%>>全部</option>
                <option value="新闻" <% if rs("Level")="新闻" then%>selected<%end if%>>新闻</option>
                <option value="影视" <% if rs("Level")="影视" then%>selected<%end if%>>影视</option>
                <option value="音乐" <% if rs("Level")="音乐" then%>selected<%end if%>>音乐</option>
		<option value="动画" <% if rs("Level")="动画" then%>selected<%end if%>>动画</option>
		<option value="游戏" <% if rs("Level")="游戏" then%>selected<%end if%>>游戏</option>
		<option value="生活" <% if rs("Level")="生活" then%>selected<%end if%>>生活</option>
		<option value="女性" <% if rs("Level")="女性" then%>selected<%end if%>>女性</option>
		<option value="网吧" <% if rs("Level")="网吧" then%>selected<%end if%>>网吧</option>
		<option value="旅游" <% if rs("Level")="旅游" then%>selected<%end if%>>旅游</option>
		<option value="软件" <% if rs("Level")="软件" then%>selected<%end if%>>软件</option>
		<option value="人才" <% if rs("Level")="人才" then%>selected<%end if%>>人才</option>
		<option value="二手" <% if rs("Level")="二手" then%>selected<%end if%>>二手</option>
		<option value="壁纸" <% if rs("Level")="壁纸" then%>selected<%end if%>>壁纸</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td height="37" colspan="2" align="center">
		<input type="hidden" name="ID" value="<%=rs("id")%>"> 
              <input type="hidden" name="action" value="edit">
              <input type="submit" name="Submit" value="提交修改">
              　 
              <input type="reset" name="reset" value="重设表单">
            </td>
          </tr>
        </table>
      </form>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</body>
</html>