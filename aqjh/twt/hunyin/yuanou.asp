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
rs.open "SELECT * FROM h where c='离婚' order by e DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>离婚</title>
<link rel="stylesheet" type="text/css" href="../../css.css">
</head>
<body text="#000000" background="../../bg.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table width="98%" align="center" cellspacing="2" border="2" cellpadding="5"
bgcolor="#90c088">
  <tr bgcolor="#f7f7f7">
    <td align="left">[共<font color="#990000"><b><%=rs.pagecount%></b></font>页][<font
color="#990000"><b><%=musers()%></b></font>人离婚] [<a
href="yuanou.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a
href="yuanou.asp?page=<%=page+1%>">下一页</a>]</font></td>
</tr>
<tr>
    <td style="font-size:21;color:#FF1493" align="center"><b>离婚信息</b></font></td>
</tr>
<tr>
    <td height="2"> 
      <table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="#000000" bordercolorlight="#EFEFEF">
        <tr bgcolor="#DFEDFD"> 
          <td width="45" height="10"> 
            <div align="center">申请人</div>
          </td>
          <td width="59" height="10"> 
            <div align="center">离婚对象</div>
          </td>
          <td width="304" height="10"> 
            <div align="center">离 婚 理 由</div>
          </td>
          <td width="74" height="10"> 
            <div align="center">分手费</div>
          </td>
          <td width="129" height="10"> 
            <div align="center">时间</font></div>
          </td>
          <td width="26"> 
            <div align="center"> 回复</font> </div>
          </td>
          <%if aqjh_grade=10 then%>
          <td width="47"> 
            <div align="center"> 判决</font> </div>
          </td>
          <%end if%>
        </tr>
        <%
count=0
do while not rs.eof and count<rs.PageSize
%>
        <tr bgcolor="#f7f7f7"> 
          <td width="45" height="8"> 
            <div align="center"><%=rs("a")%></font></div>
          </td>
          <td width="59" height="8"> 
            <div align="center"><%=rs("b")%></font></div>
          </td>
          <td width="304" height="8"> 
            <div align="center"><%=rs("f")%></font></div>
          </td>
          <td width="74" height="8"> 
            <div align="center"><%=rs("d")%> 两</font> </div>
          <td width="129" height="8"> 
            <div align="center"><%=rs("e")%> </font> </div>
          <td width="26"> 
            <%if aqjh_name=rs("b") then %>
            <a href="yuanouok.asp?name=<%=rs("a")%>&amp;my=<%=rs("b")%>"><font
            size="2">同意</font></a> 
            <%end if%>
          </td>
           <%if aqjh_grade=10 then%>
          <td width="47"> 
            <div align="center">  
              <a href="panjue.asp?name=<%=rs("a")%>&amp;my=<%=rs("b")%>">同意</a> 
              </font> </div>
          </td>
          <%end if%>
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
tmprs=conn.execute("Select count(*) As a from h")
musers=tmprs("a")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>