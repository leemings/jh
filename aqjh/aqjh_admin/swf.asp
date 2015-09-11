<%@ LANGUAGE=VBScript codepage ="936"%>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("aqjh_usermdb")
if Trim(Request.QueryString("w"))="del" then
sql="delete from swf where id="&Request.QueryString("id")
conn.execute(sql)
end if 
%>
<html>
<head>
<title>后台歌曲管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from swf order by id desc"
rs.open sql,conn,1,1
page=request.querystring("page")
PageSize = 20
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
if not rs.eof then
%>
<table width=680 border=1 cellspacing=0 cellpadding=3 bordercolorlight=#000000 bordercolordark=#FFFFFF align=center>
<tr align=center>
<td colspan=3 height=30><b>后台MTV歌曲管理</b></td>
<td width=80><a href=swf_add.asp><font color=red>添加歌曲</a></td>
</tr>
<tr align=center>
<td width=30>ID</td>
<td>歌曲名称</td>
<td>地 址</td>
<td>操 作</td>
</tr>
<%do while not (rs.eof or rs.bof) and count<rs.PageSize%>
<tr>
<td align=center><%=rs("id")%></td>
<td><a href="<%=rs("路径")%>" target="_blank"><%=rs("歌名")%></a></td>
<td><%=rs("路径")%></td>
<td align=center><a href="swf_edit.asp?id=<%=rs("id")%>" title="修改歌曲资料！">改</a> | <a href="swf.asp?w=del&id=<%=rs("id")%>" title="删除该歌曲！">删</a></td>
</tr>
<%
rs.movenext
count=count+1
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>[<a href="swf.asp?page=<%=page-1%>">上一页</a>][第<%=page%>页][<a href="swf.asp?page=<%=page+1%>">下一页</a>]
<%else%>
</table>
<table width=680 border=1 cellspacing=0 cellpadding=3 bordercolorlight=#000000 bordercolordark=#FFFFFF align=center>
<tr>
<td height=30 align=center><a href=swf_add.asp><font color=red>添加歌曲</a>
<%rs.close
set rs=nothing
conn.close
set conn=nothing
end if%>
</table>