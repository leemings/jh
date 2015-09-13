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
%>
<html>
<head>
<title>快乐江湖鲜花店</title>
<link rel="stylesheet" href="../../css.css">
</head>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="bk2.gif" oncontextmenu=self.event.returnValue=false>
<center>
<br><div align="center"><font size="5" face="黑体" color=green>江 湖 花 店</font><br><br>
<table width="530">
<tr>
    <td colspan="2" valign="top" align="center">    
      <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%">
        <tr> 
          <td width="80" height="13"> 
            <div align="center">物品名</div>
          </td>
          <td width="80" height="13"> 
            <div align="center">经营者</div>
          </td>
          <td width="234" height="13"> 
            <div align="center">功 能</div>
          </td>
          <td width="59" height="13"> 
            <div align="center">售 价</div>
          </td>
          <td width="55" height="13"> 
            <div align="center">数 量</div>
          </td>
          <td width="51" height="13"> 
            <div align="center">操 作</div>
          </td>
        </tr>
<%
rs.open "SELECT * FROM b where b='鲜花' order by h",conn,1,3
if rs.recordcount<>0 then
  rs.pagesize = 30
  page = cint(request("pageno"))
     if page <= 1 then 
        page = 1
     end if
     if page >= rs.pagecount then
        page = rs.pagecount
     end if
     rs.absolutepage = page
  for i=1 to rs.pagesize
%>
        <tr> 
          <form method=POST action='buyhua.asp'><input type=hidden name=huaname value='<%=rs("a")%>'>
            <td width="80"> 
              <div align="center"><img src="images/<%=rs("i")%>" alt="<%=rs("a")%>"></div></td>
            <td width=80> 
              <div align="center"><%=rs("n")%></div>
            </td>
            <td width="234"> 
              <div align="left">　<%=rs("a")%>：<%=rs("c")%></div>
            </td>
            <td width="59"> 
              <div align="center"><%=rs("h")%>两<br></div></td>
            <td width="55"> 
              <div align="center">
                <select name="huasl">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="5">5</option>
                  <option value="10">10</option>
                  <option value="20">20</option>
                  <option value="50">50</option>
                  <option value="99">99</option>
                  <option value="999">999</option>
                </select>
              </div>
            </td>
            <td width="51"> 
              <div align="center"> 
                <input type="submit" name="submit" value="进行">
              </div>
            </td>
          </form>
        </tr>
<% 
  rs.movenext
  if rs.eof then exit for
  next
 %>  
</table>
共<font color=2020dd><% =rs.pagecount %></font>页<font color=2020dd><% =rs.recordcount %></font>张 当前:<font color=2020dd><%=page%></font>/<%=rs.pagecount%>&nbsp;&nbsp;<a href=hua.asp?pageno=1>首页</a>&nbsp;<a href=hua.asp?pageno=<% =page-1 %>>上页</a>&nbsp;<a href=hua.asp?pageno=<% =page+1 %>>下页</a>&nbsp;<a href=hua.asp?pageno=<% =rs.pagecount %>>末页</a>
<%end if%>
</td>
</tr>
</table>
<p align=center><FONT color=#0000ff>&copy; 版权所有 2015-2015 </FONT><A href="http://www.happyjh.com/" target=_blank><FONT color=#0000ff>快乐江湖网</FONT></A></p>
</body>