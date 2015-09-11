<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>孤儿院</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center">
<tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=aqjh_name%></font>光临爱情孤儿院</b></font></div>
<form method=POST action='gryok.asp'>
<table width="300" align="center">
<tr>
<td>
<tr>
<td align=center>我的爱心： 
   <select name=money size=1>
<option value="1000" selected>1000</option>
<option value="10000">10000</option>
<option value="100000">100000</option>
<option value="1000000">1000000</option>
</select>
</td>
</tr>
<tr>
<td  align=center>
<input type=submit value=祝福这些孩子 name="submit">
</td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
                孤儿院简介</div>
</td>
</tr>
<tr>
<td valign="top" >
              <p>战争、贫穷使用他们这些孩子失去了父母亲，无助他的他们多么希望得到我们最真诚的帮助，请献出你的爱心，为世界所的孤儿捐献一片爱心！(捐献爱心可以使您德高望众的1点道德=500两！)</p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><b><font color="#0000FF">版权所有『爱情江湖总站』</font></b></font>
</div>
</body>
</html>