<%@ LANGUAGE=VBScript%>
<!--#include file="fun.inc"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "请进入聊天室再进入股票！"
window.close()
</script>
<%Response.End
end if

if sjjh_name="" then
%>
<script language="vbscript">
  alert "不能进入，你还没有登录江湖"
  location.href = "../index.asp"
</script>
<%
else
name=sjjh_name
%>
<!--#include file="jhb.asp"-->
<%
	sql="select * from 客户 where 帐号='" & name & "'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
	%>
<script language="vbscript">
  alert "你还没有开户呢！"
  location.href = "jingji.asp"
</script>
<%	   
    else
    jjr=rs("经纪人")
%>
<html>
<head>
<title>经纪人办公室</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style>
td{color:#000000}
p{font-size:16;color:red}
</style>
</head>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p> 
</td></tr>  

<tr><td style="color:red;font-size:9pt">  
<br><p align=center>您的股票经济人<%=jjr%>在此为您服务</p><br>
<a href=cun.asp>存钱进股票帐户</a><br><a href=qu.asp>从股票帐户提钱</a><br><a href=cha.asp>查看最近股票买卖情况</a><br><a href=huan.asp>换经纪人</a>  
<br>  
<p align=center><a href=jingji.asp>离开</a></p>
</table></table>  

		
</body>
</html>
<%
end if
end if
end if
%>