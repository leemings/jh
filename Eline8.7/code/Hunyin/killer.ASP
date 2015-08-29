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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from e where d<now()-5",conn,3,3
rs.open "SELECT * FROM e order by d DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>雇佣杀手</title>
<style>td           { font-size: 14; color: #000000 }
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body text="#000000" background="../images/bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font size="2"><b><font size="+2" color="#000000" face="楷体_GB2312">雇佣杀手</font></b></font> 
  <a href="addkiller.asp">我要雇佣</a> <br>
  <font color="#FF00FF">说明：江湖上的恩怨事情，如果你自己解决不了有钱就行，<br>
  当[被杀人]被杀死后复活开始计算，酬劳将付给杀手！</font><font color="#00FF33"><br>
  </font></div>
<table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="#000000" bordercolorlight="#EFEFEF">
  <tr bgcolor="#DFEDFD"> 
    <td width="6%" height="10"> 
      <div align="center"><font size="2">申请人</font></div>
    </td>
    <td width="8%" height="10"> 
      <div align="center"><font size="2">被杀人</font></div>
    </td>
    <td width="29%" height="10"> 
      <div align="center"><font size="2">有什么江湖怨恨</font></div>
    </td>
    <td width="9%" height="10"> 
      <div align="center"><font size="2">酬劳</font></div>
    </td>
    <td width="15%" height="10"> 
      <div align="center"><font size="2">申请时间</font></div>
    </td>
    <td width="33%" height="10"> 
      <div align="center"><font size="2">结果</font></div>
    </td>
  </tr>
  <%
count=0
do while not rs.eof and count<rs.PageSize
%>
  <tr bgcolor="#f7f7f7" onmouseout="this.bgColor='#f7f7f7';"onmouseover="this.bgColor='#3399CC';">
    <td width="6%" height="8"> 
      <div align="center"><font size="2"><%=rs("a")%></font></div>
    </td>
    <td width="8%" height="8"> 
      <div align="center"><font size="2"><%=rs("b")%></font></div>
    </td>
    <td width="29%" height="8"> 
      <div align="center"><font size="2"><%=rs("c")%></font></div>
    </td>
    <td width="9%" height="8"> 
      <div align="center"><font size="-2"><%=rs("e")%> 两</font> </div>
    <td width="15%" height="8"> 
      <div align="center"><font size="-2"><%=rs("d")%> </font> </div>
    <td width="33%" height="8"> 
      <div align="center"><font size="2"><%=rs("f")%></font></div>
    
  </tr>
  <%rs.movenext%>
  <%count=count+1%>
  <%loop%>
</table>
<br>
<font size="2">[共<font color="#990000"><b><%=rs.pagecount%></b></font>页][<font
color="#990000"><b><%=musers()%></b></font>人申请] [<a
href="killer.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a
href="killer.asp?page=<%=page+1%>">下一页</a>]</font> 
</body>

</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As a from h")
musers=tmprs("a")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%>