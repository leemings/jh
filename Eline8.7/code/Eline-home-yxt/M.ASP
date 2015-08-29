<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>添加新的帮派！</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 宋体}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: 宋体; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#000000" text="#000000" link="#000080" alink="#800000" vlink="#000080" background="../jhimg/bk_hc3w.gif">
<form action="newzt.asp" method=POST >
  <ul>
    <table border=1 cellspacing=0 cellpadding=0 align="center" width="560" height="104">
      <tr> 
        <td height="13"> 
          <div align="center"><font size="+1"><b>门派添加</b></font></div>
        </td>
      </tr>
      <tr> 
        <td height="27"> 
          <div align="center"><font color="#FFFFFF"><font color="#000000">门派名字:</font></font> 
            <input name="mp" value="" size=10 maxlength=10>
            掌门： 
            <input name="zm" value="" size=10 maxlength=10>
            限制： 
            <input name="xz" value="" size=3 maxlength=1>
          </div>
          <div align="center"></div>
          <div align="center"><font color="#FFFFFF"><font color="#000000">门派简介:</font></font> 
            <input name="sm" value="" size=27 maxlength=30>
            <font color="#FFFFFF"><font color="#000000"><br>
            </font><font color="#FFFFFF"><font color="#000000">限制说明:</font></font> 
            <input name="xzsm" value="" size=27 maxlength=30>
            <br>
            <font color="#000000">表 达 式:</font></font> 
            <input name="bds" value="" size=27 maxlength=30>
            <font color="#FFFFFF"> <font color="#000000"> </font></font><br>
            注：门派简介最多为30个字符，且不可为空！</div>
        </td>
      </tr>
    </table>
    <div align="center"> <font size="3" class="c" color="#000000"><br>
      <br>
      <input type="HIDDEN" name="action" value="RegSubmit">
      <input type="SUBMIT" name="Submit" value="注册">
      <input type="RESET" name="Reset" value="清除">
      </font> </div>
  </ul>
</form>
<p align="center">门派名字：1-10个字符 掌门：1-10个字符，自动设置成管理：5级 身份掌门<br>
  限制：0 没有 1限制 ，当为1时限制说明，表达式将会起作用 <br>
  门派简介：有关门派的说明，写什么都可以<br>
  限制说明：就是加入的条件如：战斗等级&gt;10级<br>
  表达式：等级&gt;10<br>
  <br>
  所对应的表达式一定要按格式书写，书写正确，否则将不能正确加入门派！</p>
</body>
</html>
