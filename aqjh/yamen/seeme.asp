<%@ LANGUAGE=VBScript codepage ="936" %>
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if aqjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="Select  * from 用户 where 姓名='"&aqjh_name&"'"
set rs=conn.Execute(sql)
%>
<html>
<head>
<title>江湖档案</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<table width="428" border="0" cellspacing="15" cellpadding="0" background="../images/b2.gif" height="323">
<tr>
<td width="415" height="310">
<table border="1"
width="400" cellspacing="1" cellpadding="1" background="../bg.gif">
<tr>
<td colspan="5" height="33">
<div align="center"><%=rs("姓名")%>大侠的江湖背景 </div>
</td>
</tr>
<tr>
<td width="86">姓 名</td>
<td colspan="2"><%=rs("姓名")%></td>
<td colspan="2"> 江湖名号： <font color="#0000FF" size="2">
<%if rs("等级")<5  then response.write("初来乍道")
if rs("等级")>=5 and rs("等级")<10  then response.write("江湖初行")
if rs("等级")>=10 and rs("等级")<15  then response.write("小有成就")
if rs("等级")>=15 and rs("等级")<20  then response.write("声名显赫")
if rs("等级")>=20 and rs("等级")<35  then response.write("行闯江湖")
if rs("等级")>=35 and rs("等级")<45  then response.write("一代大侠")
if rs("等级")>=45 and rs("等级")<55  then response.write("江湖剑客")
if rs("等级")>=55 and rs("等级")<65  then response.write("闻名天下")
if rs("等级")>=65 and rs("等级")<75  then response.write("云游仙胜")
if rs("等级")>=100 then response.write("剑仙")
%>
</font></td>
</tr>
<tr>
<td width="86" height="23">性 别</td>
<td height="23" width="84"><%=rs("性别")%> </td>
<td height="23" width="35"><font size="2">状 态</font></td>
<td height="23" width="57"><%=rs("状态")%></td>
<td height="23" width="110">会员费：<font color="#0000FF" size="2"><%=rs("会员金卡")%>元</font></td>
</tr>
<tr>
<td width="86">Email地址:</td>
<td width="84"><%=rs("信箱")%> </td>
<td width="35">OIcq:</td>
<td colspan="2"><%=rs("oicq")%></td>
</tr>
</table>
<div align="left"></div>
<table border="1"
width="400" cellspacing="1" cellpadding="1" background="../bgcheetah.gif">
<tr>
<td colspan="5" height="20">
<div align="center"> <font size="2">江 湖 档 案</font> </div>
</td>
</tr>
<tr>
<td width="66" height="25">现 金：</td>
<td width="110" height="25"><%=rs("银两")%> 两</td>
<td width="63" height="25">介绍人：</td>
<td height="25" colspan="2"><%=rs("介绍人")%></td>
</tr>
<tr>
<td width="66" height="20">存 款：</td>
<td width="110" height="24"><%=rs("存款")%> 两</td>
<td width="63" height="24">道 德：</td>
<td height="24" colspan="2"><%=rs("道德")%></td>
</tr>
<tr>
<td width="66" height="20">武 功：</td>
<td width="110" height="24"><%=rs("武功")%></td>
<td width="63" height="24">攻 击：</td>
<td height="24" colspan="2"><%=rs("攻击")%></td>
</tr>
<tr>
<td width="66" height="20">内 力：</td>
<td width="110"><%=rs("内力")%></td>
<td width="63">防 御：</td>
<td colspan="2"><%=rs("防御")%></td>
</tr>
<tr>
<td width="66" height="20">魅 力：</td>
<td width="110"><%=rs("魅力")%></td>
<td width="63">门 派：</td>
<td colspan="2"><%=rs("门派")%> </td>
</tr>
<tr>
<td width="66" height="20">体 力：</td>
<td width="110"><%=rs("体力")%></td>
<td width="63">身 份：</td>
<td colspan="2"><%=rs("身份")%></td>
</tr>
<tr>
<td width="66" height="20">配 偶：</td>
<td width="110"><%=rs("配偶")%></td>
<td width="63">赢 / 赌：</td>
<td colspan="2"><%=rs("赢次数")%> / <%=rs("赌次数")%></td>
</tr>
<tr>
<td width="66" height="20">赌场等级：</td>
<td width="110">
<%if rs("赢钱")<100  then response.write("赌场菜鸟")
if rs("赢钱")>=2000 and rs("赢钱")<10000  then response.write("赌场游民")
if rs("赢钱")>=10000 and rs("赢钱")<50000  then response.write("初级赌徒")
if rs("赢钱")>=50000 and rs("赢钱")<100000  then response.write("赌场游侠")
if rs("赢钱")>=100000 and rs("赢钱")<300000  then response.write("职业高手")
if rs("赢钱")>=300000 and rs("赢钱")<800000  then response.write("赌侠")
if rs("赢钱")>=800000 and rs("赢钱")<1500000  then response.write("赌王")
if rs("赢钱")>=1000000 then response.write("赌神")
%>
</td>
<td width="63">战斗级：</td>
<td width="61"><%=rs("等级")%>级</td>
<td width="84">管理级:<%=rs("grade")%>级</td>
</tr>
</table>
</td>
</tr>
</table>
</body>

</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>