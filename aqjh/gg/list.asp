<html>
<STYLE type=text/css>BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 18px
}
TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 18px
}
A:visited {
	COLOR: black; TEXT-DECORATION: none
}
A:link {
	COLOR: black; TEXT-DECORATION: none
}
A:hover {
	COLOR: red; TEXT-DECORATION: none
}
</STYLE>

<head>
<TITLE>◆快乐江湖◆总站『happyjh.CoM』→打造一个新的品牌</TITLE>
<%randomize timer
mysound=int(rnd*4)+1
Response.Write "<bgsound src=mid/midi"&mysound&".mid loop=-1 volume=100>"
%>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="../index1.files/style.css" type=text/css rel=stylesheet>
<noscript><iframe src=*.html></iframe></noscript>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY oncontextmenu=window.event.returnValue=false 
onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false 
text=#000000 bgColor=#ff0000 background=../index1.files/tstdn_5.jpg>

<DIV id=Layer1 
style="Z-INDEX: 1; LEFT: 687px; WIDTH: 150px; POSITION: absolute; TOP: 22px; HEIGHT: 21px"><A 
style="CURSOR: hand" 
onclick="this.style.behavior='url(#default#homepage)';&#10;&#10;this.setHomePage('http://www.happyjh.com')" 
href="#">设为首页</A> | <A  
href="javascript:window.external.AddFavorite('http://www.happyjh.com','『快乐江湖总站』')">收藏本站</A> 
</DIV>
<DIV align=center>
<TABLE height=221 cellSpacing=0 cellPadding=0 width=775 align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=480 background=../index1.files/tstdn_2.jpg height=221>
      <DIV align=right>
      <P align=center>
 </P></DIV></TD></TR></TBODY></TABLE>
<TABLE height=82 cellSpacing=0 cellPadding=0 width=775 align=center border=0>
  <TBODY>
  <TR>
    <TD width=204 background=../index1.files/tstdn_8.jpg></TD>
    <TD background=../index1.files/tstdn_9.jpg>
      <TABLE height=27 cellSpacing=0 cellPadding=0 width="97%" align=left 
      border=0>
        <TBODY>
        <TR>
          <TD height=27>
            <DIV align=right>
            <P align=center><A href="http://www.happyjh.com">首页</A> <A 
            href="http://www.happyjh.com/bbs/" target=_blank>江湖论坛</A> <A 
            href="http://www.happyjh.com/jiaoyou/index.asp" 
            target=_blank>江湖交友</A> <A 
            href="http://www.happyjh.com/joke/index.asp" 
            target=_blank>江湖笑话</A> <A 
            href="http://www.happyjh.com/ie.exe" 
            target=_blank>清除临时文件解决掉线</A> <A 
            href="http://www.happyjh.com/jhbook/index.asp" 
            target=_blank>留言本站</A></P></DIV></TD></TR>
  </TBODY></TABLE><FONT 
      size=2>&nbsp;</FONT></TD></TR><TR>
  </table>
<table border="0" cellpadding="0" cellspacing="0" width="770" background="img/index_line_1.gif">
<tr>
<td width="185" valign="top" rowspan="4">
<table width="94%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td valign="top" colspan="2"><img src="../index1.files/000111.JPG" width="185" height="43"></td>
</tr>
<tr>
<td height="107" width="12"></td>
<%
page=request.querystring("page")
id=request.querystring("id")
page=clng(page)
if page<>"" and page<>null then
if not isnumeric(page) then
Response.Write "<script language=JavaScript>{alert('提示：页数错误，请使用数字！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End 
end if
end if
if id<>"" and id<>null then
id=0
end if
if not isnumeric(id) then
Response.Write "<script language=JavaScript>{alert('提示：id错误，请使用数字！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End 
end if

PageSize=30
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
dim tmprs
sql="select count(*) as 标题 from 公告"
if isnumeric(id) and id<>"" then
sql=sql&" where id="&id
end if
tmprs=conn.execute(sql)
musers=tmprs("标题")
set tmprs=nothing
rs.open "select * from l",conn%>
<td valign="middle" width="173" height="107">
<%if rs.eof or rs.bof then%>
<div align="center"><font color=red>还没有任何分类</font></div><br>
<br>
<%else
do while not (rs.bof or rs.eof)%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="img/bookmark.gif"><a href="list.asp?id=<%=rs("id")%>"><%=rs("lb")%></a><br>
<%rs.movenext
loop
end if
rs.close%>
</td>
</tr>
<tr>
<td height="30" colspan="2"><img src="../index1.files/00011.JPG" width="185" height="42"></td>
</tr>
<tr>
<td height="60" width="12"></td>
<form name="form1" method="post" action="search.asp?action=sumit" target="_blank">
<td valign="middle" align="center" width="173"><br>
<br>
<select name="btnr">
<option value="0" selected>不指定条件</option>
<option value="1">按标题搜索</option>
<option value="2">按内容搜索</option>
</select>
<br>
<%rs.open "select * from l"%>
<br>
<select name="zdlb">
<option value="0" selected>不指定类别</option>
<%do while not (rs.eof or rs.bof )%>
<option value="<%=rs("id")%>"><%=rs("lb")%></option>
<%rs.movenext
loop
rs.close%>
</select>
<br>
<br>
<input type="text" name="nkeycode" size="14" value="关键字">
<input type="submit" name="Submit" value="搜索">
<br>
</td>
</form>
</tr>
<tr>
<td height="31" colspan="2"><img src="../index1.files/0003.JPG" width="185" height="42"></td>
</tr>
<tr>
<td height="85" width="12"></td>
<td valign="middle" width="173" height="85">
<%rs.open "select top 10 id,标题 from 公告 order by 查看次数 desc",conn,1,2
do while not(rs.eof or rs.bof)
rmbt=rs("标题")
if len(rmbt)>9 then rmbt=left(rmbt,9)&"..."%>
&nbsp;&nbsp;<font color="ffffff">◆&nbsp;<a href="#" onclick="window.open('disp.asp?id=<%=rs("id")%>','','scrollbars=yes,resizable=yes,width=670,height=500,menubar=no')" title=<%=rs("标题")%>><font color="ffffff"><%=rmbt%></a><br>
<%rs.movenext
loop
rs.close%>
</td>
</tr>
</table>
</td>
<td width="5" height="10"></td>
<td width="570"></td>
<td width="1"></td>
<td width="10"></td>
</tr>
<tr>
<td height="41" width="5"></td>
<td valign="top" width="570"><img src="../index1.files/0010.jpg" width="539" height="41" usemap="#Map" border="0"></td>
<td width="1"></td>
</tr>

<tr>
<td height="200" width="5"></td>
<td valign="top" width="570" align="center"><br>
<TABLE cellSpacing=0 cellPadding=1 width="93%" border=0>
<TBODY>
<%if isnumeric(id) and id<>0 then
rs.open "select * from l where id="&id,conn
if not (rs.eof or rs.bof) then
lbm=rs("lb")
else
lbm="未找到此分类名"
end if
rs.close
end if
sql="select * from 公告"
if isnumeric(id) and id<>"" then
sql=sql&" where 所属类别="&id
end if
sql=sql&" order by 时间"
rs.open sql,conn
rs.PageSize = PageSize
pgnum=rs.pagecount
pgnum=clng(pgnum)
if pgnum<1 then pgnum=1
if page="" or clng(page)<1 then page=1
if clng(page) >pgnum then page=pgnum
if rs.pagecount>0 then rs.AbsolutePage=page
count=0%>
<TR>
<TD colspan="4" vAlign=top height=62>当前分类：<font color=red><%=lbm%></font>
<br>
<br>
&nbsp;&nbsp;&nbsp;共 <%=musers%>条，当前第 <%=page%>/<%=pgnum%>页，每页 
<%=pagesize%>条 </TD>
</TR>
<tr>
<td width="38%" height="28">&nbsp;&nbsp;&nbsp;<img src="img/go.gif" width="9" height="9"><font face="宋体"><b>标 
题</b></font></td>
<td width="24%" height="28">
<div align="center"><b><font face="宋体">发表时间</font></b></div>
</td>
<td width="15%" height="28">
<div align="center"><b>所属类别</b></div>
</td>
<td width="23%" height="28">
<div align="center"><font face="宋体"><b>点击次数</b></font></div>
</td>
</tr>
<%if rs.eof or rs.bof then
response.write "<td colspan='3'><br><br><br><br><div align='center'><font color=red>"&lbm&" 中还没有发表过任何主题。</font></div></td>"
else
do while not(rs.eof or rs.bof) and count<rs.PageSize%>
<tr>
<td valign=top width="38%">&nbsp;&nbsp;&nbsp;<img src="img/go.gif" width="9" height="9"><a href="#" onclick="window.open('disp.asp?id=<%=rs("id")%>','','scrollbars=yes,resizable=yes,width=670,height=500,menubar=no')"><%=rs("标题")%></a>
</td>
<td valign=top width="24%">
<div align="center"><font color="#0000FF"><%=rs("时间")%></font></div>
</td>
<td valign=top width="15%">
<div align="center"><%=rs("类别名称")%></div>
</td>
<td valign=top width="23%">
<div align="center">[<font color="#FF0000"><%=rs("查看次数")%></font>] 
</div>
</td>
<tr>
<%rs.movenext
count=count+1
loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<tr>
<td colspan="4" height="23">&nbsp;</td>
</tr>
<tr>
<td colspan="4" height="27">&nbsp;</td>
</tr>
<TR>
<TD colspan="4" align=middle>
<div align="center">共 <%=musers%>条，当前第 <%=page%>/<%=pgnum%>页，每页 <%=pagesize%>条 
<br><br>
<%if page>1 then%>
<a href="list.asp?page=<%=page-1%>">上一页</a>
<%end if
for i=1 to pgnum
if page=i then%>
<B>[<%=i%>]</B>
<%else%>
<A href="list.asp?page=<%=i%>"><b><font color=red>[<%=i%>]</font></b></A>
<%end if
next
if page<pgnum then%>
<a href="list.asp?page=<%=page+1%>">下一页</A>
<%end if%>
</div>
</TD>
</TR>
<TR>
<TD colspan="4" align=middle>&nbsp;</TD>
</TR>
</TBODY>
</TABLE>
</td>
<td width="1"></td>
</tr>
</table>
<TABLE WIDTH=770 BORDER=0 CELLPADDING=0 CELLSPACING=0 bgcolor="#B56F00">
    <TR> 
      <TD width="10" bgcolor="#B56F00"></TD>
      </TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><FONT 
      size=2></FONT></TABLE></DIV>
<DIV align=center>
<TABLE height=70 cellSpacing=0 cellPadding=0 width=775 align=center 
background=index1.files/bottom_t.jpg border=0>
  <TBODY>
  <TR>
    <TD height=70>
      <CENTER>Copyright&copy;2004 happyjh.CoM&#8482; All Rights Reserved  
      快乐江湖总站<BR>全站设计修改:快乐江湖　客服QQ:865240608  
TEL:13635716119</CENTER></TD></TR></TBODY></TABLE></DIV></FORM></BODY>
</html>
