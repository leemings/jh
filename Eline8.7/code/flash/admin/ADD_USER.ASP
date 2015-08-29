<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="session.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="level.asp"-->
<html>
<head>
<title>添加用户</title>
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
        
  <table width="36%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
            <td height="27" width="120" align="center">用户姓名：</td>
            
      <td height="27" width="159"> 
        <input type="text" name="username" size="20" maxlength="20" class="form">
            </td>
          </tr>
          <tr> 
            <td height="27" width="120" align="center">用户密码：</td>
            
      <td height="27" width="159"> 
        <input type="password" name="password" size="20" maxlength="20" class="form">
            </td>
          </tr>
          <tr> 
            <td height="27" width="120" align="center">用户等级：</td>
            
      <td height="27" width="159"> 
        <select name="level" class="form">
                <option value="" selected>请选择用户等级</option>
		<option value="全部">全部</option>
                <option value="新闻">新闻</option>
                <option value="影视">影视</option>
                <option value="音乐">音乐</option>
		<option value="动画">动画</option>
		<option value="游戏">游戏</option>
		<option value="生活">生活</option>
		<option value="女性">女性</option>
		<option value="网吧">网吧</option>
		<option value="旅游">旅游</option>
		<option value="软件">软件</option>
		<option value="人才">人才</option>
		<option value="二手">二手</option>
		<option value="壁纸">壁纸</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td height="37" colspan="2" align="center"> 
              <input type="hidden" name="action" value="add">
              <input type="submit" name="Submit" value="提交用户">
              　 
              <input type="reset" name="reset" value="重设表单">
            </td>
          </tr>
        </table>
      </form>
</body>
</html>