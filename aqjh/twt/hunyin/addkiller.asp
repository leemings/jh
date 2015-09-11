<%@ LANGUAGE=VBScript codepage ="936" %><html>
<head>
<title>雇佣杀手</title>
<style></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../css.css">
</head>
<body bgcolor="#000000" text="#ffffff" link="#0000FF" alink="#0000FF"
vlink="#0000FF" leftmargin="0" topmargin="0" background="../../bg.gif">
<center>
  <h2 style="color:#F4A460"><font color="#000000">想要找杀手要写上以下信息</font></h2>
<form method="POST" action="addkillerok.asp">
    <table width="78%" border="0" cellspacing="2" cellpadding="2">
      <tr>
        <td width="32%"><img src="killer.jpg" width="256" height="210"></td>
        <td width="68%"> 
          <table width="338" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="#339999">
            <tr> 
              <td width="67"> 
                <div align="center"><font color="#000000" size="-1">将要杀谁：</font></div>
              </td>
              <td width="247"><font color="#000000" size="-1"> 
                <input type="text" name="name" size="10" maxlength="10">
                (<font color="#0000FF">仇人名字</font>)</font></td>
            </tr>
            <tr> 
              <td width="67" height="28"> 
                <div align="center"><font color="#000000" size="-1">佣 金：</font></div>
              </td>
              <td width="247" height="28"><font color="#000000" size="-1"> 
                <input type="text" name="money" size="9" maxlength="9">
                两(<font color="#0000FF">1万-1亿</font>)</font></td>
            </tr>
            <tr> 
              <td width="67"> 
                <div align="center"><font color="#000000" size="-1">杀人理由：</font></div>
              </td>
              <td width="247"><font color="#000000" size="-1"> 
                <input type="text" name="mess" size="25" maxlength="50">
                50汉字</font></td>
            </tr>
            <tr> 
              <td width="67"> 
                <div align="center"><font color="#000000" size="-1"></font></div>
              </td>
              <td width="247"><font color="#000000" size="-1"> 
                <input type="submit" value="加入" name="submit">
                <input type="reset" value="取消" name="reset">
                </font></td>
            </tr>
          </table>
</td></tr></table>
    <br>
    <font color="#0000FF">在这里你可以使用钱杀你的仇人！</font> 
  </form></center></body></html>