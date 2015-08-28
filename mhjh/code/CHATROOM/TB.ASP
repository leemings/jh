<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from 用户 where 姓名='"&username&"'",conn
if rs.EOF or rs.BOF then Response.Redirect "../error.asp?id=016"            
%>    
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel='stylesheet' href='css.css'>
<title>投奔国家</title>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFDDF2" background='bg1.gif'>
<br><br>
<p align="center">选择要投奔的国家</p>
<div align="center">
  <center>
  <table border="1" width="125">
    <tr>
      <td width="60" align="center">国家</td>
      <td width="61" align="center">国力</td>
    </tr>
    <tr>
      <td width="60" align="center"><a href="tb2.asp?mg=秦">秦</a></td>
      <%set rst=server.CreateObject("adodb.recordset")
      rst.Open "select count(*) as cnumber from 用户 where 国家='秦'",conn
      if rst.EOF or rst.BOF then
         rst("cnumber")="0"
      else
         cnumber=rst("cnumber")
      end if
      rst.Close           
      %>
      <td width="61" align="center"><%=cnumber%></td>                
    </tr>
    <tr>
      <td width="60" align="center"><a href="tb2.asp?mg=楚">楚</a></td>
            <%
      rst.Open "select count(*) as cnumber from 用户 where 国家='楚'",conn
      if rst.EOF or rst.BOF then
         rst("cnumber")="0"
      else
         cnumber=rst("cnumber")
      end if
      rst.Close           
      %>
      <td width="61" align="center"><%=cnumber%>　</td> 
    </tr>
    <tr>
      <td width="60" align="center"><a href="tb2.asp?mg=赵">赵</a></td>
            <%
      rst.Open "select count(*) as cnumber from 用户 where 国家='赵'",conn
      if rst.EOF or rst.BOF then
         rst("cnumber")="0"
      else
         cnumber=rst("cnumber")
      end if
      rst.Close           
      %>
      <td width="61" align="center"><%=cnumber%>　</td>
    </tr>
    <tr>
      <td width="60" align="center"><a href="tb2.asp?mg=魏">魏</a></td>
<%rst.Open "select count(*) as cnumber from 用户 where 国家='魏'",conn
      if rst.EOF or rst.BOF then
         rst("cnumber")="&nbsp;"
      else
      cnumber=rst("cnumber")
      end if
      rst.Close           
      %>
      <td width="61" align="center"><%=cnumber%>　</td>
    </tr>
    <tr>
      <td width="60" align="center"><a href="tb2.asp?mg=齐">齐</a></td>
 <%rst.Open "select count(*) as cnumber from 用户 where 国家='齐'",conn
      if rst.EOF or rst.BOF then
         rst("cnumber")="&nbsp;"
      else
         cnumber=rst("cnumber")
      end if
      rst.Close           
%>
      <td width="61" align="center"><%=cnumber%>　</td>
    </tr>
    <tr>
      <td width="60" align="center"><a href="tb2.asp?mg=韩">韩</a></td>
     <% rst.Open "select count(*) as cnumber from 用户 where 国家='韩'",conn
        if rst.EOF or rst.BOF then
           rst("cnumber")="&nbsp;"
        else
           cnumber=rst("cnumber")
        end if
        rst.Close           
%>
      <td width="61" align="center"><%=cnumber%>　</td>
    </tr>
    <tr>
      <td width="60" align="center"><a href="tb2.asp?mg=燕">燕</a></td>
<%    rst.Open "select count(*) as cnumber from 用户 where 国家='燕'",conn
      if rst.EOF or rst.BOF then
         rst("cnumber")="&nbsp;"
      else
         cnumber=rst("cnumber")
      end if
      rst.Close           
      %>
      <td width="61" align="center"><%=cnumber%>　</td>
    </tr>
  </table>
  </center>
</div>
<p align="center">积分-4000<br>体力-10万<br>银两-10万</p>
</body>                             
</html>                  
<%
rs.Close           
set rs=nothing
set rst=nothing             
conn.Close           
set conn=nothing
%>
