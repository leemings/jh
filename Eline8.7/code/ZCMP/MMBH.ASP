<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from 用户 where 姓名='"&sjjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你不是江湖中人？！');window.close();</script>"
	response.end
end if
rs.close
conn.close
set conn=nothing
%>
<html>
<head>
<title>密码保护</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<bgsound src="pic/bj.mid" loop="-1">
</head>

<body background="pic/WALLPAPER256.JPG" text="#000000">
<p><font color="#FF0000">敬告：</font></p>
<p><font color="#000000">如果你还没有申请过</font><font color="#FF00FF">密码保护</font><font color="#000000">，请尽快申请，因为</font><font color="#FF00FF">密码保护</font><font color="#000000">是在你密码丢失后</font><b><font color="#FF00FF">唯一</font></b><font color="#000000">可以找回的办法。</font></p>
<p>请点击这里：<a href="mmbh1.asp">申请密码保护</a></p>
<p>请点击这里： <a href="mmbhxg.asp">修改密码保护资料 </a></p>
<%if sjjh_grade>=10 then%>
<p><a href="mmbhcx.asp">查询密码保护资料（站长操作）</a></p>
<%end if%>
<p>　</p>
<p align="center"><font color="#FF0000" size="2"><b>&nbsp;『快乐江湖』之密码保护</b></font></p>
<p align="center"><font color="#FF0000" size="2"><b>程序修改:回首当年</b></font></p>
</body>
</html>