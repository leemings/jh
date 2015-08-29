<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
%>
<html>
<head>
<title><%=Application("sjjh_chatroomname")%>医院</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style type="text/css">td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { color: #FFC106; font-family: 宋体; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<br>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="318">&nbsp;</td>
<td width="460"><font color="#FF3300"><b>患病人员名单</b></font> </td>
</tr>
</table>
<br>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
<td bgcolor="#000000">&nbsp;</td>
</tr>
<tr bgcolor="#990000">
    <td><img src="eline.jpg" width="778" height="80"></td>
</tr>
<tr>
<td bgcolor="#000000">&nbsp;</td>
</tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="778">
<tr>
<td>
<div align="center">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
align="center" >
<tr>
<td colspan="2" valign="top" width="100%" height="110">
<div align="center"> <font color="#FF0000" size="5"> <br>
                <font size="4">道德-5000点的时候记得及时来治疗噢，否则后果会很严重 ^_^</font></font> 
                <table border="1" cellspacing="1" cellpadding="0" width="100%"
bordercolor="#000000" align="center">
<tr>
<td width="20%" nowrap>
<div align="center">精神病患者</div>
<td width="20%" nowrap>
<div align="center">门派</div>
</td>
<td width="20%">
<div align="center">身份</div>
</td>
<td width="20%">
<div align="center">道德</div>
</td>
<td width="20%">
<div align="center">请求</div>
</td>
</tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 where 道德<-1000",conn
do while not rs.bof and not rs.eof
%>
<tr>
<td width="20%">
<p align="center"><%=rs("姓名")%></td>
<td width="20%">
<p align="center"><%=rs("门派")%></td>
<td width="20%">
<p align="center"><%=rs("身份")%></td>
<td width="20%">
<p align="center"><%=rs("道德")%></td>
<%if rs("道德")<-5000 then %>
<% else
if rs("登录")>now() then%><%else%><% conn.execute("update 用户 set 状态='正常',登录=now() where 姓名='"&rs("姓名")&"'")
end if%>
<td width="20%">
<p align="center"> <a href="jt.asp?name=<%=rs("姓名")%>"><font color="#008000">看病</font><%end if%></a>   
</td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
                <p align="center">朋友你罪孽深重，我们来让你重生吧...来感化你的灵魂：天灵灵、地灵灵，各路大仙听我令，让这位大位从恶梦中醒来吧！啊门！<br>
                  <font color="#CC3300">您的道德：+100，银两全部捐献，存款花了1/2 用于作法费用，积分减100点你就可以重生了！</font> 
                  <br>
                </p>
</div>
</td>
</tr>
</table>
</div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>

