<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
if username<>Application("yx8_mhjh_admin") then Response.Redirect "../error.asp?id=046"%>
<html>
<head>
<title>站长发放操作</title>
<link rel=stylesheet href='css.css'>
</head>
<body background='bg1.gif' topmargin=30>
<div align="center">
  <font face="宋体"><b>站长发放物品</b><br>
  </font>
  <table align="center"  border="1" cellPadding="0" cellSpacing="0" width="70%" bordercolorlight="#000000" bordercolordark="#FFFFFF" height="254">
    <tbody>
      <tr>
        <td id="sm" height="21">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=银两&value=500000"  title="发放银两，最大值为10万两!">银两</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="21">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=体力&value=1000000"  title="发放体力，最大值为1万!">体力</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="21">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=积分&value=1000"  title="发放积分，最大值为1000点!">积分</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="21">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=内力&value=1000000"  title="发放豆点，最大值为1万!">内力</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="21">
          <div align="center">
            <a class=blue href="#" onClick="window.open('dalie/zj.html','fangliwu','scrollbars=no,resizable=no,width=180,height=180')" title="点击后，请不要关闭小窗口！">老虎</a>  
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="21">
          <div align="center">
<a class=blue href="#" onClick="window.open('hand.asp','fangjs','scrollbars=yes,width=200,height=200,left=20,top=10')">僵尸</a>                                                                                                  

          </div>
        </td>
      </tr>
      <tr>
        <td height="21" id="sm">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=衰哥&value=10000"  title="发放衰哥，很恐怖，最大体力1万点!">衰哥</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td height="21" id="sm">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=元宝&value=100000"  title="发放元宝，最大银两1万!">元宝</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="21">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=精力&value=10000" title="发放精力，最高发放5千点!">精力</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="21">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=攻击&value=10000" title="发放攻击，最高发放5千点!">攻击</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="22">
          <div align="center">
            <font size="-1"><a href="fafang2.asp?cz=防御&value=20000" title="发放防御，最高发放5千点!">防御</a></font>
          </div>
        </td>
      </tr>
      <tr>
        <td id="sm" height="22">
          <div align="center">
            <font size="-1">选择要发放的操作，点击即可，不要放的太多哟(:</font>
          </div>
        </td>
      </tr>
    </tbody>
  </table>
</div>
</body>
</html>
