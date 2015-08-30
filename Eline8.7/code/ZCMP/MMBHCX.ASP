<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then
	Response.Write "<script Language=Javascript>alert('你不是站长？！');window.close();</script>"
	response.end
end if
%>
<html>
<head>
<title>密码保护查询修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<bgsound src="pic/bj.mid" loop="-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="pic/WALLPAPER256.JPG">
<div align="center">用户密码保护资料查询 </div>
<form name="form1" method="post" action="mmbhcx1.asp">
  <div align="center">请输入要查询的用户ID号:  
    <input type="text" name="id" size="5" maxlength="5">
    <input type="submit" name="locat" value="查询"> 
  </div>
</form>
<p>　</p>
<p align="center"><a href="mmbhall.asp">查看所有人</a> </p>
<p align="center"><font color="#FF0000" size="2"><b>『快乐江湖』之密码保护</b></font></p>
<p align="center"><font color="#FF0000" size="2"><b>程序修改:回首当年</b></font></p>
</body>
</html>