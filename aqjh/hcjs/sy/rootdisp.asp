<%if Session("aqjh_name")=""  then
  Response.Redirect "../error.asp?id=440"
else
dim conn,rs
id=request.querystring("id")
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
rs.open ("SELECT * FROM sy WHERE ID=" & id),conn,0,1
%>

<html>
<head>
<title>◎  -&gt; 查看状词</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../css.css" rel=stylesheet></head>
<body background="../../2di.jpg" bgcolor="#CCCCCC" topmargin="0">
<hr size="1">
<table width="590" border="0" cellspacing="0" cellpadding="3" align="center">
<tr>
<td>受害人<b>:<%=rs("原告")%></b>&nbsp;悲惨地写道：</td>
<td align="right"><font color="red">&lt;&lt;&nbsp; &lt;&lt; <a href="manage.asp"><font color="red">返回喊冤告状</a></font></td>  
</tr>  
</table>  
<table width="588" border="1" cellspacing="1" cellpadding="0" align="center" bordercolor="#FFFFFF" bgcolor="#000000">  
<tr bgcolor="#000000">  
<td height="67">  
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">  
<tr>  
<td bgcolor="#8EB4D9">  
<p><span class="txt"><br>  
<%=changechr(rs("状词"))%><br>  
<br>  
<br>  
对被告这种恶劣行径，希望给予<%=rs("要求")%> <br>  
<br>  
<font size="2" color="#ffffff"><%if session("aqjh_grade")>=10 then%></font>&gt;&gt;<a href="agree.asp?result=<%=rs("要求")%>&id=<%=id%>&bg=<%=rs("被告")%>&my=<%=rs("原告")%>">同意</a>  
&gt;&gt;<a href="noagree.asp?id=<%=id%>">不同意</a>  
&gt;&gt; <a href="del.asp?id=<%=id%>">删除此记录</a><a href="noagree.asp?id=<%=id%>"></a><%end if %></span></p>  
<p align="right">　</p>  
</td>  
</tr>  
</table>  
</td>  
</tr>  
</table>  
<%rs.close  
set rs=nothing  
%>  
<hr size="1">  
<p align="center">　</p>  
</body>  
</html>  
<%  
end if  
function getorder(theid)  
dim tmprs  
tmprs=conn.execute("select [Order] from bbs Where ID=" & theid)  
getorder=tmprs("Order")  
set tmprs=nothing  
end function  
  
function changechr(str)  
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")  
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")  
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")  
end function  
%>