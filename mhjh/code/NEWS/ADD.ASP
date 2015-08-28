<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
if session("yx8_mhjh_username")<>Application("yx8_mhjh_admin") then Response.Redirect "../error.asp?id=046"
%>
<html>

<head>
<title>cctony首页更新</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet>
</head>
<body background='../chatroom/bg1.gif'>
<table width="81%" border="0" cellspacing="0" cellpadding="3" align="center">
<tr><td><table border="1" width="100%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" align="center">
<tr align=center>
<td bgcolor="#33CCFF"><B>添加新公告</B></td></tr><tr><td>
<form method="POST" action="save.asp" >
<div align="center">标题     
<input type="text" name="name" size="60" value="">     
<br>     
内容 <br>     
<textarea name="neirong" rows="3" cols="60"></textarea>     
<br>     
<input type="submit" name="Submit" value="提交">     
<input type="reset" name="Submit2" value="清除">     
</div>     
</form></td></tr></table></td></tr></table>     
<p align="center">    
<%    
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then    
CurPage = 1    
Else    
CurPage = CINT(Request.QueryString("CurPage"))    
End If    
%>    
<!-- #include file="connect.asp" -->    
<%rs.Open "Select * From data Order By time DESC", conn, 1,1%>    
<%if rs.eof and rs.bof then%>    
暂时没有任何记录！！     
<%else     
RS.PageSize=10'设置每页记录数     
Dim TotalPages     
TotalPages = RS.PageCount     
     
If CurPage>RS.Pagecount Then     
CurPage=RS.Pagecount     
end if     
     
RS.AbsolutePage=CurPage     
rs.CacheSize = RS.PageSize'设置最大记录数     
Dim Totalcount     
Totalcount =INT(RS.recordcount)     
     
StartPageNum=1     
do while StartPageNum+10<=CurPage     
StartPageNum=StartPageNum+10     
Loop     
     
EndPageNum=StartPageNum+9     
     
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount       %>     
<table border="1" width="80%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#999999" align="center">     
<tr>     
<td bgcolor="#33CCFF" height="20"><b>&nbsp;标题</b></td>     
<td  align="center" bgcolor="#33CCFF" height="20"><b>时 间</b></td>     
<td height="20" align="center" bgcolor="#33CCFF"><b>操作</b></td>     
</tr>     
<%I=0     
p=RS.PageSize*(Curpage-1)     
do while (Not RS.Eof) and (I<RS.PageSize)     
p=p+1%>     
<tr>     
<td bordercolorlight="#C0C0C0" align=left><b><%=rs("name")%></b></td>     
<td  align="center">[<%=rs("time")%>]</td>     
<td bordercolorlight="#c0c0c0" align=middle><a href="del.asp?id=<%=rs("id")%>">删除</a>|<a href="modifile.asp?id=<%=rs("id")%>">修改</a></td>    
</tr>    
<%I=I+1    
RS.MoveNext    
Loop%>    
<tr><td colspan=4 align=middle bordercolorlight="#C0C0C0">页次： <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%>    
页数： <a href="add.asp?CurPage=<%=StartPageNum-1%>">&lt;&lt;</a>    
<% For I=StartPageNum to EndPageNum    
if I<>CurPage then %> <a href="add.asp?CurPage=<%=I%>"><%=I%></a> <% else %>    
<%=I%> <% end if %> <% Next %>    
<% if EndPageNum<RS.Pagecount then %> <a href="add.asp?CurPage=<%=EndPageNum+1%>">下页</a>    
<%end if%>|<a href="add.asp">刷 新</a></td></tr>    
</table>    
</body>    
</html>    
<%    
end if    
rs.close    
set rs=nothing    
%>    
    
    
    
