<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>门派基金♀wWw.happyjh.com♀</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center">
<tr>
<td height="81" valign="top">
<div align="center"><font color="#000000"><b>欢迎<font color=blue><%=sjjh_name%></font>光临门派基金库</b></font></div>
<form method=POST action='putmpjj.asp'>
<table width="300" align="center">
<tr>
<td>
<tr>
<td align=center>请选择所以存入的钱数：
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
<input type=submit value=好了就这样 name="submit">
</td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
门派基金</div>
</td>
</tr>
<tr>
<td valign="top" >
<p>门派基金，是一项取之于民，用之于民的政策，所取得金额存放于本门派金库中，由执法长老分配给所需要的人。您对本派的付出我会有应有的回报的。对于本门支持最大或从本门取钱最多的人我们都会有记录排行的，请到江湖排行中查询！</p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><b><font color="#0000FF"><%=Application("sjjh_chatroomname")%></font></b></font>
</div>
</body>
</html>



