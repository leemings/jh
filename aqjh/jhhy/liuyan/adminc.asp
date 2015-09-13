<!--#include file="conn.asp"-->
<!--#include file="ubb.asp"-->
<%
if session("aqjh_username")<>Application("aqjh_adminok") then
response.write"<script>alert('对不起，只有江湖主站长才有权查看这里的内容！'); history.back();</script>"
response.end()
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" rel=stylesheet  content="text/css">
<title>快乐江湖</title>
<script language="JavaScript">
function dis(t,s,y){
  if(t.style.display=="none"){
    for(g=1;g<=y;g++){
	  document.all("book"+g).style.display="none";
	  document.all("plus"+g).src="img/plus.gif";
	}
    s.src="img/minus.gif";
	s.title="点击关闭此贴";
	t.style.display="";
	}
	else{
	s.src="img/plus.gif";
	s.title="点击查看此贴";
	t.style.display="none";
	}
}
</script>
</head>

<!--#include file="top.asp"-->
<table border="0" cellpadding="0" cellspacing="0" align=center>
  <tr>
      <td width="600" align="center" valign=top height=100 style="border-left:1px #cccccc solid">
<!--=========================-->
<%sql="select * from book order by lastdate desc"
set rs=server.createobject("adodb.recordset")                     
rs.open sql,conn,1,1                     
if rs.eof and rs.bof then                     
response.write"还没有任何留言"
else  
'分页的实现 
listnum=20
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
<table width="590" border="0" cellspacing="1" cellpadding="0" align=right bgcolor=0080C0>
<%
l=1
do while not rs.eof and i<listnum
n=n+1
aa=rs("lastdate")
a=right(year(aa),2)
b=month(aa)
c=day(aa)%>
  <tr bgcolor=#efefef>
    <td width="590" height=20 align=center colspan="3"><font color="#FF0000">有什么事在这里留言给站长哦！他会尽力帮助你的</font></td>
  </tr>  
  <tr bgcolor=#efefef>
    <td width="40" height=20 align=center><%=n%></td>
    <td width="470">&nbsp;<img src="img/plus.gif" id="plus<%=l%>" title="点击查看此贴" style="cursor:hand;" onclick="dis(book<%=l%>,plus<%=l%>,<%if Rs.recordcount>listnum then%><%=listnum%><%else%><%=Rs.recordcount%><%end if%>)"> <a href=view.asp?id=<%=rs("id")%>><%if len(rs("title"))>31 then response.write left(rs("title"),30)&"..." else response.write rs("title")%></a></td>
    <td width="80" align=center><%=rs("user")%></td>   
    <td width="80" align=center>[ <%=a&"-"&b&"-"&c%> ]</td>   
  </tr>   
  <tr style="display:none" id="book<%=l%>" bgcolor=#ffffff>   
    <td colspan=3>   
   
<table width="550" border="0" cellspacing="0" cellpadding="0" align=center>   
<tr><td height=10 colspan=2></td></tr>   
<tr><td height=30 colspan=2><p style="line-height:150%"><%=code_jk(rs("book"))%></td></tr>   
<tr><td height=10 colspan=2></td></tr>   
<%set rs1=server.createobject("adodb.recordset")   
sql1="select * from re where id="&rs("id")   
rs1.open sql1,conn,1,1   
rehtc=rs1.recordcount   
rs1.close%>   
<tr><td width="354" height=27 valign=bottom style="border-top:1px #0080c0 solid"><strong>发贴人：<font color=red><%=rs("user")%></font>　查看/回复：<font color=red><%=rs("htc")%>/<%=rehtc%></font></strong></td> 
<td align=right valign=bottom style="border-top:1px #0080c0 solid"><strong>[<a href=view.asp?id=<%=rs("id")%>>查看所有</a>]　[<a href=guest_re.asp?id=<%=rs("id")%>>回复此贴</a>]</strong></td></tr> 
<tr><td height=10 colspan=2></td></tr> 
</table> 
 
	</td> 
  </tr> 
 
<%rs.movenext  
i=i+1  
l=l+1 
j=j-1 
loop%> 
<%filename="index.asp"%> 
<%response.write "<form method=post action="&filename&">"%> 
<tr> 
<td colspan=3 align=right bgcolor=#cccccc><%=Rs.recordcount%> 篇贴子　<%=listnum%> 条/页　共 <%=Rs.pagecount%> 页    
<% if page=1 then   
else%><a href=<%=filename%>><strong>|<<</strong></a>    
<a href=<%=filename%>?page=<%=page-1%>><strong><<</strong></a>    
<a href=<%=filename%>?page=<%=page-1%>><b>[<%=page-1%>]</b></a>    
<%end if%>   
<% if Rs.pagecount=1 then %><%else%> <b>[<%=page%>]</b> <%end if%>   
<% if Rs.pagecount-page <> 0 then %>   
<a href=<%=filename%>?page=<%=page+1%>><b>[<%=page+1%>]</b></a>    
<a href=<%=filename%>?page=<%=page+1%>><strong>>></strong></a>    
<a href=<%=filename%>?page=<%=Rs.pagecount%>><strong>>>|</strong></a>    
<%end if%>　   
<%response.write "转到：<input type='text' name='page' style='height:18' size=4 maxlength=10 value="&page&">"                  
response.write " <input type='submit' value='GO' style='height:18' name='cndok'>"%> &nbsp;</td>   
<%end if%></tr>   
</form>   
</table>   
<!--=========================-->   
	</td>   
  </tr>   
</table>   
</html>



