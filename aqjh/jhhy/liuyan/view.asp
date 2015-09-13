<!--#include file="conn.asp"-->
<!--#include file="ubb.asp"-->
<%
user=session("aqjh_name")
if user="" then
response.write"<script>alert('对不起，只有江湖玩家才可以在这里发布信息给站长，请您去登陆江湖之后再来！'); history.back();</script>"
response.end()
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" rel=stylesheet  content="text/css">
<title>快乐江湖</title>
</head>

<!--#include file="top.asp"-->
<table width="760" border="0" cellpadding="0" cellspacing="0" align=center>
  <tr>
    <td align="center" valign=top height=100>
<!--=========================-->
<%sql="select * from book where id ="&request("id")
set rs=server.createobject("adodb.recordset")                     
rs.open sql,conn,1,1                     
if rs.eof and rs.bof then                     
response.write"此贴不存在"
else
conn.execute("update book Set htc=htc+1 where id="&request("id"))
%>
<%set rs1=server.createobject("adodb.recordset")
sql1="select * from re where id="&rs("id")
rs1.open sql1,conn,1,1
rehtc=rs1.recordcount
adid=rs("id")
rs1.close%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor=#efefef>
    <td height=20 bgcolor="#0080C0">&nbsp;<img src="img/minus.gif"> <font color=#ffffff><strong><%=rs("title")%></strong></font></td>
  </tr>
  <tr>
    <td bgcolor="#F7FCFF" style="border:1px #0080C0 solid">

<table width="100%" border="0" cellspacing="0" cellpadding="0" align=center>
<tr><td height=10 colspan=2></td></tr>

<tr><td width="200" height=80 align=center style="border-right:1px #0080C0 solid" valign=top>

<table width="140" border="0" cellspacing="0" cellpadding="0" align=center>
<tr><td height=10><p style="line-height:150%"><strong>发贴：<font color=red><%=rs("user")%></font><br>查看：<font color=red><%=rs("htc")%></font><br>回复：<font color=red><%=rehtc%></font></strong></td></tr>
</table>

</td>
<td width="560" valign=top>

<table width="500" border="0" cellspacing="0" cellpadding="0" align=center>
<tr><td height=80 valign=top><p style="line-height:150%"><%=code_jk(rs("book"))%></td></tr>
<tr><td height=10></td></tr>
</table>

</td></tr>

<tr><td height=30 style="border-right:1px #0080C0 solid">
<table width="150" border="0" cellspacing="0" cellpadding="0" align=center style="border-top:1px #0080C0 solid">
<tr><td height=30 align="center"><strong><%=rs("date")%></strong></td></tr></table>
</td>
<td>
<table width="500" border="0" cellspacing="0" cellpadding="0" align=center style="border-top: 1px solid #0080C0">
<tr><td height=30 align=right width="498"></td></tr></table>
</td></tr>
<tr><td height=10 colspan=2></td></tr>
</table>

	</td>
  </tr>
</table>
<%end if
rs.close%>
<!--=========================-->
<%sql="select * from re where id="&request("id")
set rs=server.createobject("adodb.recordset")                     
rs.open sql,conn,1,1                     
if rs.eof and rs.bof then                     
response.write""
else  
'分页的实现 
listnum=10
Rs.pagesize=listnum
page=Request("page")
if (page-Rs.pagecount) > 0 then
page=Rs.pagecount
elseif page = "" or page < 1 then
page = 1
end if
Rs.absolutepage=page
'编号的实现
j=rs.recordcount
j=j-(page-1)*listnum%>
<%i=0
nn=request("page")
if Rs.recordcount<(nn-1)*listnum then
response.redirect"index.asp?page="&nn-1&""
end if
if nn="" then
n=0
else
nn=nn-1
n=listnum*nn
end if%>
      <table width="760" border="0" cellspacing="0" cellpadding="0">
        <%do while not rs.eof and i<listnum
n=n+1
if right(n,1)=1 or right(n,1)=3 or right(n,1)=5 or right(n,1)=7 or right(n,1)=9 then
aaa="#ffffff" 
else
aaa="#F7FCFF"
end if
%>
        <tr> 
          <td bgcolor="<%=aaa%>" style="border-left:1px #0080c0 solid;border-right:1px #0080c0 solid;border-bottom:1px #0080c0 solid"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0" align=center>
              <tr>
                <td height=10 colspan=2></td>
              </tr>
              <tr>
                <td width="200" height=80 align=center style="border-right:1px #0080C0 solid" valign=top> 
                  <table width="140" border="0" cellspacing="0" cellpadding="0" align=center>
                    <tr>
                      <td height=10><p style="line-height:150%"><strong>回复：<font color=red><%=rs("user")%></font></strong></td>
                    </tr>
                  </table></td>
                <td width="560" valign=top> <table width="500" border="0" cellspacing="0" cellpadding="0" align=center>
                    <tr>
                      <td height=80 valign=top><p style="line-height:150%"><%=code_jk(rs("re"))%></td>
                    </tr>
                    <tr>
                      <td height=10></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td height=30 style="border-right:1px #0080C0 solid"> <table width="150" border="0" cellspacing="0" cellpadding="0" align=center style="border-top:1px #0080C0 solid">
                    <tr>
                      <td height=30 align="center"><strong><%=rs("date")%></strong></td>
                    </tr>
                  </table></td>
                <td> <table width="510" border="0" cellspacing="0" cellpadding="0" align=center style="border-top:1px #0080C0 solid">
                    <tr>
                      <td height=30 align=right><strong>第 <font color=red><%=n%></font> 
                        楼</strong></td>
                    </tr>
                  </table></td>
              </tr>
              <tr>
                <td height=10 colspan=2></td>
              </tr>
            </table></td>
        </tr>
        <%rs.movenext 
i=i+1 
l=l+1
j=j-1
loop%>
        <%filename="view.asp?id="&request("id")&""%>
        <%response.write "<form method=post action="&filename&">"%>
        <tr> 
          <td align=right bgcolor=#0080C0><font color=#ffffff><%=Rs.recordcount%> 
            篇回复　<%=listnum%> 条/页　共 <%=Rs.pagecount%> 页     
            <% if page=1 then    
else%>    
            <a class=1 href=<%=filename%>><strong>|<<</strong></a> <a class=1 href=<%=filename%>&page=<%=page-1%>><strong><<</strong></a>     
            <a class=1 href=<%=filename%>&page=<%=page-1%>><b>[<%=page-1%>]</b></a>     
            <%end if%>    
            <% if Rs.pagecount=1 then %>    
            <%else%>    
            <b>[<%=page%>]</b>     
            <%end if%>    
            <% if Rs.pagecount-page <> 0 then %>    
            <a class=1 href=<%=filename%>&page=<%=page+1%>><b>[<%=page+1%>]</b></a>     
            <a class=1 href=<%=filename%>&page=<%=page+1%>><strong>>></strong></a>     
            <a class=1 href=<%=filename%>&page=<%=Rs.pagecount%>><strong>>>|</strong></a>     
            <%end if%>    
            　     
            <%response.write "转到：<input type='text' name='page' style='height:18' size=4 maxlength=10 value="&page&">"                   
response.write " <input type='submit' value='GO' style='height:18' name='cndok'>"%>    
            &nbsp;</font></td>   
          <%end if%>   
        </tr>   
        <tr>   
          <td align=center bgcolor=#0080C0>   
            <table border="0" cellspacing="1" width="50%">   
              <tr>   
                <td width="50%" align="center">  
                  <p align="center"><a href=chakan.asp><font color="#FFFF00"><b>【返回】</b></font></a></p>
                </td> 
                <%if session("aqjh_name")=Application("aqjh_adminok") then%><td width="50%" align="center"> 
                  <p align="center"><a href=guest_re.asp?id=<%=adid%>><font color="#FFFF00"><b>【回复此贴】</b></font></a></td><%end if%> 
              </tr> 
            </table> 
          </td> 
        </tr></form> 
      </table> 
<!--=========================--> 
	</td> 
  </tr> 
</table> 
</html>