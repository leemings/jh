<%@ LANGUAGE=VBScript%>
<%
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();}</script>"
	Response.End 
end if
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if sjjh_name="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim page
page=request.querystring("page")
myname=trim(request.querystring("myname"))
PageSize = 15
rs.open "delete * from s where m<now()-10 and c<>'保险柜'",conn,3,3
if myname="" then
	rs.open "Select * From s where c='交易' Order by m DESC",conn,3,3
else
	rs.open "Select * From s where c='交易' and (d='"& myname &"' or b="& sjjh_id &") Order by m DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title><%=Application("sjjh_chatroomname")%>-商贸交易</title>
<style type="text/css">td           { font-family: 宋体; font-size: 9pt }
body         { font-family: 宋体; font-size: 9pt }
select       { font-family: 宋体; font-size: 9pt }
a            { color: #FFC106; font-family: 宋体; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 宋体; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#66CCCC">
<table border="0" height="24" width="91%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="15" width="79%" bgcolor="#669999"><font color="#669966"> <font color="#FFFFFF"><b>商贸交易(</b><font color="#000000">注意10天交易不成功的过期物品我们将送到垃圾场处理!</font><b>）<a                    href="smjy.asp?myname=<%=sjjh_name%>">查找我的交易</a>
</b></font><font color="#669966"><font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新此页面</a></font></font><font color="#CC0000" face="幼圆"></font><font color="#FFFFFF"><b>
</b></font></font></td>
<td height="15" width="21%" bgcolor="#669999">
<div align="right"><font color="#669966"><font color="#FFFFFF"><b><a href="wupin.asp" target="_blank">处理自己物品点这里</a></b></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
<table width="92%" align="center" cellspacing="0" border="0"
cellpadding="0">
<tr>
<td width="100%">
<table border="1" cellspacing="1" cellpadding="0" width="732"
bordercolorlight="#EFEFEF">
<tr bgcolor="#FFFFFF">
<td width="7%" height="16">
<div align="center"><font color="#FF6600">物品名</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600">拥有者</font></div>
</td>
<td width="4%" height="16">
<div align="center"><font color="#FF6600"> 数量</font></div>
</td>
<td width="9%" height="16">
<div align="center"><font color="#FF6600">对象</font></div>
</td>
<td width="15%" height="16">
<div align="center"><font color="#FF6600"> 时 间 </font></div>
</td>
<td width="34%" height="16">
<div align="center"><font color="#FF6600">说 明</font></div>
</td>
<td width="5%" height="16">
<div align="center"><font color="#FF6600"> 原价</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600">现价</font></div>
</td>
<td width="12%" height="16">
<div align="center"><font color="#FF6600">购买</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';">
<td width="7%" height="21">
<div align="center"> <font color="#0000FF"><%=rs("a")%></font>
</div>
</td>
<td width="7%" height="21">
<div align="center"> <%=rs("b")%> </div>
</td>
<td width="4%" height="21">
<div align="center"> <%=rs("j")%> </div>
</td>
<td width="9%" height="21">
<div align="center"><%=rs("d")%></div>
</td>
<td width="15%" height="21">
<div align="center"> <%=rs("m")%> </div>
</td>
<td width="34%" height="21">
<div align="left"></div>
<%=rs("k")%></td>
<td width="5%" height="21">
<div align="center"> <%=rs("n")%> </div>
</td>
<td width="7%" height="21">
<div align="center"> <%=rs("l")%> </div>
</td>
<td width="12%" height="21">
<div align="center">
<% if sjjh_grade>=9 then%><a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">删除</font></a><%end if%>
<%if rs("d")=sjjh_name then%>
<a href="smjy1.asp?id=<%=rs("id")%>"><font color="#0000FF"><b>要了</b></font></a>
<%else%>
<font color="#0000FF">no交易</font>
<%end if%>
</div></td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
mu=musers()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
<td align="left" width="37%" height="2">[共<font color="red"><b><%=pa%></b></font>页][<font
color="red"><b><%=mu%></b></font>次交易]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="smjy.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a                    href="smjy.asp?page=<%=page+1%>">下一页</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 数量 from s where c='交易'")
musers=tmprs("数量")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>

</body>