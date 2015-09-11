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
rs.open "delete * from h where e<now()-2",conn,3,3
rs.open "SELECT * FROM h where c='结婚' order by e DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html><head>
<title>月老</title>
<link rel="stylesheet" type="text/css" href="../style.css">
<style>td           { font-size: 14; color: #000000 }
</style>
</head>
<body text="#000000" background="../../bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table width="98%" align="center" cellspacing="2" border="2" cellpadding="5"
bgcolor="#90c088">
  <tr bgcolor="#f7f7f7">
<td align="left"><font size="2">[共<font color="#990000"><b><%=rs.pagecount%></b></font>页][<font
color="#990000"><b><%=musers()%></b></font>人求婚] [<a
href="yuelao.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a
href="yuelao.asp?page=<%=page+1%>">下一页</a>]</font></td>
</tr>
<tr>
<td style="font-size:21;color:#FF1493" align="center"><font size="2"><b><font size="+2">求婚信息</font></b></font></td>
</tr>
<tr>
    <td height="10"> 
      <table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="#000000" bordercolorlight="#EFEFEF">
        <tr bgcolor="#DFEDFD"> 
          <td width="7%" height="10"> 
            <div align="center"><font size="2">求婚者</font></div>
          </td>
          <td width="9%" height="10"> 
            <div align="center"><font size="2">求婚对象</font></div>
          </td>
          <td width="43%" height="10"> 
            <div align="center"><font size="2">真情表白</font></div>
          </td>
          <td width="11%" height="10"> 
            <div align="center"><font size="2">聘礼</font></div>
          </td>
          <td width="22%" height="10"> 
            <div align="center"><font size="2">时间</font></div>
          </td>
          <td width="8%" height="10"> 
            <div align="center"><font size="2">回复</font></div>
          </td>
        </tr>
        <%
count=0
do while not rs.eof and count<rs.PageSize
%>
        <tr bgcolor="#f7f7f7"> 
          <td width="7%" height="8"> 
            <div align="center"><font size="2"><%=rs("a")%></font></div>
          </td>
          <td width="9%" height="8"> 
            <div align="center"><font size="2"><%=rs("b")%></font></div>
          </td>
          <td width="43%" height="8"> 
            <div align="center"><font size="2"><%=rs("f")%></font></div>
          </td>
          <td width="11%" height="8"> 
            <div align="center"><font size="2"><%=rs("d")%> 两</font> </div>
          
          <td width="22%" height="8"> 
            <div align="center"><font size="2"><%=rs("e")%> </font> </div>
          <td width="8%" height="8"> 
            <div align="center"><font size="2"> 
             <%if trim(rs("b"))=aqjh_name then %>
              <a href="yuelaook.asp?name=<%=rs("a")%>">同意</a> 
               <%end if%>
              </font></div>
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
</table></td></tr>
</table></body></html>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As a from h")
musers=tmprs("a")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>