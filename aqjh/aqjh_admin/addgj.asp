<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>添加新的国家</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<form action="addgjok.asp" method=POST >
  <ul>
    <table border="0" cellspacing="1" cellpadding="4" bgcolor="#B8AF86" cellspacing=0 cellpadding=0 align="center" width="560" height="104">
      <tr> 
        <td height="13" bgcolor=#336633> 
          <div align="center"><b><font color=ffffff>国家添加</b></div>
        </td>
      </tr>
      <tr bgcolor="f2f2ea"> 
        <td height="27"> 
          <div align="center"><font color="#FFFFFF"><font color="#000000">国家名字:</font></font> 
            <input name="gj" value="" size=10 maxlength=10 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
            君主： 
            <input name="jz" value="" size=10 maxlength=10 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
          </div>
          <div align="center"></div>
          <div align="center"><font color="#FFFFFF"><font color="#000000">国家简介:</font></font> 
            <input name="sm" value="这里欢迎您的加入！" size=27 maxlength=30 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
            <font color="#FFFFFF"><font color="#000000"><br>
            </font><font color="#FFFFFF"><font color="#000000">限制说明:</font></font> 
            <input name="xzsm" value="加入此国家无限制" size=27 maxlength=30 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
            <br>
            <font color="#000000">表 达 式:</font></font> 
            <input name="bds" value="True" size=27 maxlength=30 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
            <font color="#FFFFFF"> <font color="#000000"> </font></font><br>
            注：国家简介最多为30个字符，且不可为空！</div>
        </td>
      </tr>
    </table>
    <div align="center"> <font size="3" class="c" color="#000000"><br>
      <br>
      <input type="HIDDEN" name="action" value="RegSubmit">
      <input type="SUBMIT" name="Submit" value="注册" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
      <input type="RESET" name="Reset" value="清除" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
      </font> </div>
  </ul>
</form>
<p align="center">国家名字：1-10个字符 君主：1-10个字符<br>
  国家简介：有关国家的说明，写什么都可以<br>
  限制说明：就是加入的条件如：战斗等级&gt;10级<br>
  表达式为：等级&gt;10<br>
  如只允许2级以上女玩家加入 <br>
  表大式为：等级&gt;=2 and 性别='女'<br>
  如任意思加入请输入：True<br>
  <br>
  <br>
  所对应的表达式一定要按格式书写，书写正确，否则将不能正确加入国家！</p>
</body>
</html>
