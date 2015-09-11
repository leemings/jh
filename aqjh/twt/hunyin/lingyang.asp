<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "SELECT * FROM gry order by id DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>领养小孩</title>
<link rel="stylesheet" type="text/css" href="../../css.css">
</head>
<body text="#000000" background="../../bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table width="98%" align="center" cellspacing="2" border="2" cellpadding="5"
bgcolor="#90c088">
  <tr bgcolor="#f7f7f7">
    <td align="left">[共<font color="#990000"><b><%=rs.pagecount%></b></font>页][<font
color="#990000"><b><%=musers()%></b></font>个孤儿] [<a
href="lingyang.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a
href="lingyang.asp?page=<%=page+1%>">下一页</a>]</font></td>
</tr>
<tr>
    <td style="font-size:21;color:#FF1493" align="center"><b>孤儿院</b></font></td>
</tr>
<tr>
    <td height="2"> 
      <table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="#000000" bordercolorlight="#EFEFEF">
        <tr bgcolor="#DFEDFD"> 
          <td width="110" height="10"> 
            <div align="center">姓名</div>
          </td>
          <td width="50" height="10"> 
            <div align="center">性别</div>
          </td>
          <td width="150" height="10"> 
            <div align="center">生日</div>
          </td>
          <td width="90" height="10"> 
            <div align="center">体力</div>
          </td>
          <td width="150" height="10"> 
            <div align="center">父母</font></div>
          </td>
          <td width="50"> 
            <div align="center">操作</font> </div>
          </td><%if aqjh_grade=10 then%>
          <td width="47">
            <div align="center">管理</font> </div>
          </td>   <%end if%>       
        </tr>
        <%
count=0
do while not rs.eof and count<rs.PageSize
boy=rs("boy")
fm1=rs("fm1")
fm2=rs("fm2")
zt=split(boy,"|")
if boy="" or UBound(zt)<>7 then
 sx="出错"
 boy_name=sx
 boy_sex=sx
 boy_sr=sx
 boy_tl=sx
 boy_fm1=sx
 boy_fm2=sx
else
 boy_name=zt(0)
 boy_sex=zt(1)
 boy_sr=zt(2)
 boy_tl=zt(4)
 boy_fm1=fm1
 boy_fm2=fm2
end if
%>
        <tr bgcolor="#f7f7f7"> 
          <td> 
            <div align="center"><%=boy_name%></font></div>
          </td>
          <td> 
            <div align="center"><%=boy_sex%></font></div>
          </td>
          <td> 
            <div align="center"><%=boy_sr%></font></div>
          </td>
          <td> 
            <div align="center"><%=boy_tl%></font> </div>
          <td> 
            <div align="center"><%=boy_fm1%>|<%=boy_fm1%></font> </div>
          <td><div align="center"><a href=lingyangok.asp?id=<%=rs("id")%>>领养</a>
          </td><%if aqjh_grade=10 then%>
          <td>
            <div align="center"><a href=lingyangok.asp?w=del&id=<%=rs("id")%>>删除</a></div>
          </td><%end if%>
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
</td></tr></table>
</body></html>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As id from gry")
musers=tmprs("id")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>