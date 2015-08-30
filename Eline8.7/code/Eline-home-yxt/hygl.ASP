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
rs.open "SELECT * FROM q where d='未付' and b<now()-7 order by b DESC",conn,3,3
do while not rs.eof 
	conn.Execute "update 用户 set 状态='监禁',事件原因='监禁:\n您在我们这里办理会员："&rs("e")&"，\n因为你不付款，我们监禁你的账号！!' where 姓名='"&rs("a")&"'"
	conn.Execute "update q set d='监禁' where a='"&rs("a")&"'"

rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim page
page=request.querystring("page")
id=0
id=request.querystring("id")
if isnull(id) or id="" then id=0
PageSize = 15
if id=1 then
	rs.open "SELECT * FROM q where d='监禁' order by b DESC",conn,3,3
else
	rs.open "SELECT * FROM q where d='未付' order by b DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>会员在线申请♀wWw.happyjh.com♀</title>
<link rel="stylesheet" type="text/css" href="../style.css">
<style>td           { font-size: 14; color: #000000 }
</style>
</head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table width="95%" align="center" cellspacing="2" border="2" cellpadding="5"
bgcolor="#90c088">
  <tr bgcolor="#f7f7f7"> 
    <td align="left" width="56%"><font size="2">[共<font color="#990000"><b><%=rs.pagecount%></b></font>页][<font
color="#990000"><b><%=musers()%></b></font>人申请] [<a
href="hygl.asp?page=<%=page-1%>&id=<%=id%>">上一页</a>][第<%=page%>页][<a
href="hygl.asp?page=<%=page+1%>&id=<%=id%>">下一页</a>]</font></td>
    <td align="left" width="44%"> 
      <div align="right"><font size="2">(<a href="hyglsy.asp">警告所有未付</a> <a href="hygl.asp">查看未付款</a> 
        <a href="hygl.asp?id=1">查看监禁用户</a>)</font></div>
    </td>
  </tr>
  <tr> 
    <td style="font-size:21;color:#FF1493" align="center" colspan="2"><font size="2"><b><font size="+2">会</font><font size="2"><b><font size="+2">员在线申请</font><font size="2"><b><font size="2">(</font></b><%if id=1 then%></font>监禁用户查看</b><%else%>未付款用户查看<%end if%><b><font size="2">)</font></b></font></b></font></td>
  </tr>
  <tr> 
    <td height="10" colspan="2"> 
      <table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="#000000" bordercolorlight="#EFEFEF">
        <tr bgcolor="#DFEDFD"> 
          <td width="11%" height="10"> 
            <div align="center"><font size="2">申请者</font></div>
          </td>
          <td width="13%" height="10"> 
            <div align="center"><font size="2">申请时间</font></div>
          </td>
          <td width="47%" height="10"> 
            <div align="center"><font size="2">申请信息</font></div>
          </td>
          <td width="10%" height="10"> 
            <div align="center"><font size="2">状态</font></div>
          </td>
          <td width="19%" height="10"> 
            <div align="center"><font size="2">操作</font></div>
          </td>
        </tr>
        <%
count=0
do while not rs.eof and count<rs.PageSize
%>
        <tr bgcolor="#f7f7f7"> 
          <td width="11%" height="8"> 
            <div align="center"><font size="2"><%=rs("a")%></font></div>
          </td>
          <td width="13%" height="8"> 
            <div align="center"><font size="2"><%=rs("b")%></font></div>
          </td>
          <td width="47%" height="8"> 
            <div align="center"><font size="2"><%=rs("e")%></font>ip:<font size="2"><%=rs("f")%></font></div>
          </td>
          <td width="10%" height="8"> 
            <div align="center"><font size="2"><%=rs("d")%></font> </div>
          <td width="19%" height="8"> 
            <div align="center"><font size="2"> <a href="hygljg.asp?name=<%=rs("a")%>&mess=<%=rs("e")%>">警告</a> 
              <a href="hygljj.asp?name=<%=rs("a")%>&mess=<%=rs("e")%>">监禁</a> 
              <a href="hygljs.asp?name=<%=rs("a")%>">结算</a> </font></div>
          </td>
        </tr>
        <%rs.movenext%>
        <%count=count+1%>
        <%loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
      </table>
    </td>
  </tr>
</table>
</body>

</html>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As a from q")
musers=tmprs("a")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%>