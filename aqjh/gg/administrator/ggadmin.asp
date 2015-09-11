<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if aqjh_grade<>10 or aqjh_name<>application("aqjh_user") then Response.Redirect "../../error.asp?id=439"
page=request.querystring("page")
sfsz=request.querystring("action")
fs=""
jg=""
if sfsz="L" then
	fs=lcase(trim(request.form("fs")))
	jg=trim(request.form("jg"))
end if
PageSize=30
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
if sfsz="L" and jg<>"" then
	if fs="id" then
		cx=fs&"="&jg
	else
		cx="instr("&fs&",'"&jg&"')<>0"
	end if
	sql="select * from 公告 where "&cx
else
	sql="select * from 公告 order by 时间 desc"
end if
'response.write sql
'response.end
set rs=conn.execute(sql)
rs.PageSize = PageSize
pgnum=rs.pagecount
if pgnum<1 then pgnum=1
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if rs.pagecount>0 then rs.AbsolutePage=page
count=0
%>
<html>
<head>
<title>首页站长公告管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!-- 
a { text-decoration: none} 

--> 
A:link {COLOR: #0000FF;FONT-FAMILY: 宋体; }
A:visited {COLOR: #0000FF; FONT-FAMILY: 宋体; }
A:active {FONT-FAMILY: 宋体; }
A:hover {FONT-FAMILY: 宋体;COLOR: #FF0000; }
BODY {FONT-FAMILY: 宋体; FONT-SIZE: 9pt;COLOR: #000000;}
TABLE {FONT-FAMILY: 宋体; FONT-SIZE: 9pt;COLOR: #000000;}
</style>
</head>
<body bgcolor="#FFFFFF" background="../../jhimg/bk_hc3w.gif" style="font-size: 12px">
<p align="center"><font color="#0000FF">首页站长公告管理</font></p>
<p align="center"><a href="gggl.asp">类别管理</a><font color="#0000FF">&nbsp;</font></p>
<p></p>
<table border="1" width="648" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="25" align="center">
  <tr> 
    <td align="center" nowrap height="25" width="5%">ID号</td>
    <td align="center" nowrap height="25" width="44%">标题</td>
    <td align="center" nowrap height="25" width="14%">发表时间</td>
    <td align="center" nowrap height="25" width="14%">所属类别</td>
    <td align="center" nowrap height="25" width="11%">点击次数</td>
    <td align="center" nowrap height="25" width="12%"><a href=addgg.asp>添加新公告</a></td>
  </tr>
  <%
if rs.eof or rs.bof then
%>
  <tr> 
    <td bgcolor="#85C2E0" onMouseOut="this.bgColor='#85C2E0';"onMouseOver="this.bgColor='#DFEFFF';" align="center" nowrap height="25" colspan="6"><font color="ff0000" size="2">还没有添加任何公告!</font></td>
  </tr>
  <%response.end
end if
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="#85C2E0" onMouseOut="this.bgColor='#85C2E0';"onMouseOver="this.bgColor='#DFEFFF';" font size="2";> 
    <td  align="center" nowrap height="25" width="5%"><%=rs("id")%></td>
    <td nowrap height="25" width="44%"><a href="../disp.asp?id=<%=rs("id")%>" target=_blank><%=rs("标题")%></a></td>
    <td  align="center" nowrap height="25" width="28%"><%=rs("时间")%></td>
	<td  align="center" nowrap height="25" width="28%"><%=rs("类别名称")%></td>
    <td  align="center" nowrap height="25" width="11%"><%=rs("查看次数")%></td>
    <td  align="center" nowrap height="25" width="12%"><a href="modigg.asp?id=<%=rs("id")%>">修改</a> 
      <a href="delgg.asp?id=<%=rs("id")%>" title='删除此公告'>删除</a></td>
  </tr>
  <%
	rs.movenext
loop 
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
  <tr>
    <form name="form1" method="post" action="ggadmin.asp?action=L">
      <td colspan="6"> 
        <div align="right"> 
          <select name="fs">
            <option value="标题" selected>标题</option>
            <option value="id">ＩＤ</option>
            <option value="查看次数">点击</option>
            <option value="所属类别">类别</option>
          </select>
          <input type="text" name="jg">
          <input type="submit" name="Submit" value="查找">
        </div>
      </td>
    </form>
  </tr>
</table>
<p align="center">当前第 <%=page%>/<%=pgnum%> 页，每页 <%=pagesize%> 条 
  <%ys=1
		  if page>pgnum then%>
  <a href="ggadmin.asp?page=<%=page-1%>"><br>
  上一页</a> 
  <%end if
		  do while ys<pgnum
		  	if page=ys then%>
  <B>[<%=ys%>]</B> 
  <%else%>
  <A class=black href="ggadmin.asp?page=<%=ys%>">[<%=ys%>]</A> 
  <%end if
			ys=ys+1
		  loop
		  if page<pgnum then%>
  <a href="list.asp?page=<%=page+1%>">下一页</A> 
  <%end if%>
</p>
<p align="center">本程序由永不放弃编写</p>
</body>
</html>