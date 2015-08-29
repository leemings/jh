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
<link rel="stylesheet" type="text/css" href="../style.css">
<style>td           { font-size: 14; color: #000000 }
</style>
</head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<a href="lihun.asp"><font color="#0000FF">我要登记</font></a><font color="#0000FF"><br>
<a href="yuanou.asp" title="离婚要损失生命100">江湖怨偶</a> </font> <br>
<font color="#0000FF"><a href="Stunt.asp" title="离婚要损失生命100"> 合体特技</a> </font>
<table width="95%" align="center" cellspacing="2" border="2" cellpadding="5"
bgcolor="#90c088">
  <tr bgcolor="#f7f7f7">
    <td align="left"><font size="2">[共<font color="#990000"><b><%=rs.pagecount%></b></font>页][<font
color="#990000"><b><%=musers()%></b></font>人离婚] [<a
href="yuelao.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a
href="yuelao.asp?page=<%=page+1%>">下一页</a>]</font></td>
</tr>
<tr>
    <td style="font-size:21;color:#FF1493" align="center"><font size="2"><b><font size="+2">离婚信息</font></b></font></td>
</tr>
<tr>
    <td height="2"> 
      <table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="#000000" bordercolorlight="#EFEFEF">
        <tr bgcolor="#DFEDFD"> 
          <td width="45" height="10"> 
            <div align="center"><font size="2">申请人</font></div>
          </td>
          <td width="59" height="10"> 
            <div align="center"><font size="2">离婚对象</font></div>
          </td>
          <td width="304" height="10"> 
            <div align="center"><font size="2">离 婚 理 由</font></div>
          </td>
          <td width="74" height="10"> 
            <div align="center"><font size="2">分手费</font></div>
          </td>
          <td width="129" height="10"> 
            <div align="center"><font size="2">时间</font></div>
          </td>
          <td width="26"> 
            <div align="center"> <font size="2">回复</font> </div>
          </td>
          <%if sjjh_grade=10 then%>
          <td width="47"> 
            <div align="center"> <font size="2">判决</font> </div>
          </td>
          <%end if%>
        </tr>
        <%
count=0
do while not rs.eof and count<rs.PageSize
%>
        <tr bgcolor="#f7f7f7"> 
          <td width="45" height="8"> 
            <div align="center"><font size="2"><%=rs("a")%></font></div>
          </td>
          <td width="59" height="8"> 
            <div align="center"><font size="2"><%=rs("b")%></font></div>
          </td>
          <td width="304" height="8"> 
            <div align="center"><font size="2"><%=rs("f")%></font></div>
          </td>
          <td width="74" height="8"> 
            <div align="center"><font size="2"><%=rs("d")%> 两</font> </div>
          <td width="129" height="8"> 
            <div align="center"><font size="2"><%=rs("e")%> </font> </div>
          <td width="26"> 
            <%if sjjh_name=rs("b") then %>
            <a href="yuanouok.asp?name=<%=rs("a")%>&amp;my=<%=rs("b")%>"><font
            size="2">同意</font></a> 
            <%end if%>
          </td>
           <%if sjjh_grade=10 then%>
          <td width="47"> 
            <div align="center"> <font size="2"> 
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
</td>
</tr>
</table>
</body>

</html>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As a from h")
musers=tmprs("a")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%>