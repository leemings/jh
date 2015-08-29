<%@ LANGUAGE=VBScript codepage ="936" %><html>
<head>
<title>江湖怨偶</title>
<style></style>
<script LANGUAGE="javascript">
<!--
function FrmAddLink_onsubmit() {
if (document.FrmAddLink.name.value=="")
{
alert("用户名称没有填！")
document.FrmAddLink.name1.focus()
return false
}
else if(document.FrmAddLink.pass.value=="")
{
alert("密码没有填！")
document.FrmAddLink.pass.focus()
return false
}
else if(document.FrmAddLink.name2.value=="")
{
alert("离婚对象没有填！")
document.FrmAddLink.name2.focus()
return false
}
else if(document.FrmAddLink.liyou.value=="")
{
alert("离婚理由没有填！")
document.FrmAddLink.liyou.focus()
return false
}
}

//-->
</script>
<link rel="stylesheet" href="../setup.css">
</head>
<body bgcolor="#000000" text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../jhimg/bk_hc3w.gif">
<center>
<h2> <font size="5" color="#FF00FF">请大侠仔细填写以下信息 </font></h2>
<table cellspacing=0 cellpadding=0 width=548
border=0 align="center" height="142">
<tbody>
<tr>
<td valign=top width=18 height=3>&nbsp;</td>
<td valign=top width=616 height=3>&nbsp;</td>
<td valign=top width=10 height=3>&nbsp;</td>
</tr>
<tr>
<td valign=top width=18 height=59>&nbsp;</td>
<td valign=top width=616 height=59>
<form method=POST action='lihunok.asp'>
<table width="300" border="0" cellspacing="2" cellpadding="2">
<tr>
<td><img src="../jhimg/yl3.gif" width="250" height="200"></td>
              <td>
                <table width="340" border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td width="74"><font color="#000000" size="-1">离 婚 人：</font></td>
                    <td width="132"><font color="#000000" size="-1"> 
                      <input type="text" name="name" size="10" maxlength="10">
                      </font></td>
                  </tr>
                  <tr> 
                    <td width="74"><font color="#000000" size="-1">分 手 费：</font></td>
                    <td width="132"><font color="#000000" size="-1"> 
                      <input type="text" name="money" size="9" maxlength="9">
                      两</font></td>
                  </tr>
                  <tr> 
                    <td width="74"><font color="#000000" size="-1">离婚理由：</font></td>
                    <td width="132"><font color="#000000" size="-1"> 
                      <input type="text" name="mess" size="30" maxlength="50">
                      </font></td>
                  </tr>
                  <tr> 
                    <td width="74"><font color="#000000" size="-1"></font></td>
                    <td width="132"><font color="#000000" size="-1"> 
                      <input type="submit" value="加入" name="submit">
                      <input type="reset" value="取消" name="reset">
                      </font></td>
                  </tr>
                </table>
              </td>
</tr>
</table>
          <p align="center"><font color="#000000" size="2">你离婚登记登记一次需要银两10000!注意，<b><font color="#0000FF">如果结婚不到一个月是不可能离婚的！</font></b></font></p>
</form>
</td>
<td valign=top width=10 height=59>&nbsp;</td>
</tr>
<tr>
<td valign=top width=18 height=8>&nbsp;</td>
<td valign=top width=616 height=8>&nbsp;</td>
<td valign=top width=10 height=8>&nbsp;</td>
</tr>
</tbody>
</table>
</center>
</body>
</html>