<%@ codepage ="936" %>
<%
username=session("sjjh_name")
if username="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
const MaxPerPage=10
sql="select * from 用户  where 姓名='"&username&"'"
set rs=conn.execute(sql)
%>
<html>
<head>
<title>快乐江湖募捐</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="chat/READONLY/Style.css">
</head>

<body bgcolor="#000000" text="#FFFFFF" background="../chat/img/Bk_hc3w.gif">
<p align="center"><img border="0" src="../chat/img/menoy.gif"><b><font size="4" color="#008080">江湖募捐<img border="0" src="../chat/img/1414.jpg"></font></b></p>
<form method="post" action="mj.asp">
  <table width="500" border="1" cellspacing="0" cellpadding="5" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr> 
      <td bgcolor="#008080"><font color="#FFFFFF">募捐人：</font></td>
      <td bgcolor="#008080"><font color="#FFFFFF"><%=session("sjjh_name")%>　</font></td>
      <td bgcolor="#008080"><font color="#FFFFFF">魅力:<%=rs("魅力")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#008080"><font color="#FFFFFF">现在银两：</font></td>
      <td bgcolor="#008080"><font color="#FFFFFF"><%=rs("银两")%>&nbsp;元</font></td>
      <td bgcolor="#008080"><font color="#FFFFFF">攻击:<%=rs("攻击")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#008080"><font color="#FFFFFF">募捐金额:</font></td>
      <td bgcolor="#008080"> 
        <font color="#FFFFFF"> 
        <input type="text" name="mj" size="10" maxlength="10">
        最小金额10000</font></td> 
      <td bgcolor="#008080"><font color="#FFFFFF">防御:<%=rs("防御")%></font></td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#008080" > 
        <div align="center"> 
          <font color="#FFFFFF"> 
          <input type="submit" name="Submit" value="募 捐">  　   
          <input type="reset" name="Submit2" value="清 空"></font>
        </div>
      </td>
      <td bgcolor="#008080"><font color="#FFFFFF">道德:<%=rs("道德")%></font></td>
    </tr>
  </table>
</form>
<div align="center"><font color="#000000"><br>
  </font> 
  <font size="2" color="#339966">江湖募捐可对你有大大的好处哟，钱多了没地方花可以捐给社会大众嘛！<br>
  捐100000两白银你的魅力会增加10,防御会加10,道德会提升20哟<br>
  没有魅力可不能结婚呀，哪有人愿意找个大恶贼做伴侣?? 
  </font><font color="#000000"> <br>
  <a href="../desktop/"><img src="back.GIF" width="75" height="40" border="0"></a> 
  </font> </div>
<div align="center"><font color="#800080" size="3">『快乐江湖』修改制作</font> </div>
</body>
</html>