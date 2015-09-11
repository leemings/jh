<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<%
if aqjh_grade<>10 then response.redirect "../error.asp?id=440"
PageSize=20
page=clng(request.querystring("page"))
fs=request.querystring("fs")
if page="" or clng(page)<1 then page=1
if page<1 then page=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
dim tmprs
sql="select count(*) as 标题 from letter"
tmprs=conn.execute(sql)
xjzs=tmprs("标题")
set tmprs=nothing
rs.open "select * from letter",conn,3,3
rs.PageSize=PageSize
pgnum=rs.pagecount
pgnum=clng(pgnum)
if pgnum<1 then pgnum=1
if page>pgnum then page=pgnum
if rs.pagecount>0 then rs.AbsolutePage=page
count=0
%>
<HTML>
<HEAD>
<TITLE>雪落枫飞江湖→鸿雁传情</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="../../jhimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=../../jhimg/back002.gif bgColor=#996666 leftMargin=0 oncontextmenu="return false;" text=#cccccc topMargin=0 marginheight="0" marginwidth="0">
<div align=center><br>
  <table border="0" cellpadding="0" cellspacing="0" width="655">
    <tr> 
      <td width="22"><img src="../../jhimg/jiao_z_s.gif" width="21" height="21"></td>
      <td width="628" background="../../jhimg/line_s.gif">&nbsp;</td>
      <td width="13"><img src="../../jhimg/jiao_y_s.gif" width="22" height="21"></td>
</tr>
<tr> 
      <td width="22" background="../../jhimg/line_z.gif">&nbsp;</td>
      <td width="628" align=center> 
        <TABLE bgColor=#675138 width=619 height="195">
          <TBODY> 
          <TR align=middle> 
            <TD height="35" bgcolor="#333300" width="611"> <SPAN class=title><IMG align=absMiddle height=18 src="../../jhimg/icon0.gif" width=17> 
              雪落枫飞江湖→短信管理 <IMG align=absMiddle height=18 src="../../jhimg/icon0.gif" width=17></SPAN> 
            </TD>
</TR>
<TR> 
            <TD height="180" bgcolor="#675138" width="611"> 
              <p style=" line-height: 150%" align="center">总共<%=xjzs%>条短信 
                <%if page>1 then%>
                <a href="adminletter.asp?page=<%=page-1%>&fs=<%=fs%>">上一页</a> 
                <%end if%>
                当前第<font color="#00FF00"><%=page%></font>页 
                <%if page<pgnum then%>
                <a href="adminletter.asp?page=<%=page+1%>&fs=<%=fs%>">下一页</a> 
                <%end if%>
                共<font color="#FF0000"><%=pgnum%></font>页 [ 
<%for i=1 to pgnum
	if page=pgnum then
		response.write "<b>"&i&"</b>"
	else
		response.write "<a href='adminletter.asp?page="&i&"&fs="&fs&"'>"&i&"</a> "
	end if
next%>
] <%if aqjh_name=application("aqjh_user") then%><a href="delall.asp">全部删除</a><%end if%>
<DIV align=center> 
<table bordercolorlight="ff6600" align="center" width="609" border="1" cellpadding="2" cellspacing="0">
<tr> 
<td width="212" height="25">标题</td>
<td width="61" height="25" align="center">发件人</td>
<td width="61" height="25" align="center">收件人</td>
<td width="101" height="25" align="center">时间</td>
<td width="94" height="25" align="center">删除</td>
</tr>
<%if rs.eof or rs.bof then%>
<td colspan="5" height="34"> 还没有人发过短信！</td>
<%else
do while not (rs.bof or rs.eof) and count<rs.pagesize
bt=rs("bt")
cs=rs("cs")
%>
<tr> 
<td width="212" height="29"><a href="#" onclick="javascript:window.open('../sletter.asp?id=<%=rs("id")%>','','scrollbars=yes,resizable=yes,width=550,height=400')"><%=bt%></a></td>
<td width="61" height="29" align="center"><a href='#' onclick=window.open('../../yamen/mt.asp?action=<%=rs("fjr")%>','','scrollbars=yes,toolbar=no,menubar=no,width=460,height=430')><%=rs("fjr")%></a></td>
<td width="61" height="29" align="center"><a href='#' onclick=window.open('../../yamen/mt.asp?action=<%=rs("sjr")%>','','scrollbars=yes,toolbar=no,menubar=no,width=460,height=430')><%=rs("sjr")%></a></td>
<td width="101" height="29" align="center"><%=rs("rq")%></td>
<td width="94" height="29" align="center"><a href=del.asp?id=<%=rs("id")%>>删除</a>
</td>
</tr>
<%
count=count+1
rs.movenext
loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
</table>
<br>
<input type="button" value="关闭窗口" onClick="window.close()" name="button">
<br>
</DIV>
</TD>
</TR>
</TBODY> 
</TABLE>
</td>
      <td width="13" background="../../jhimg/line_y.gif">&nbsp;</td>
</tr>
<tr> 
      <td width="22"><img src="../../jhimg/jiao_z_x.gif" width="22" height="21"></td>
      <td width="628" background="../../jhimg/line_x.gif"></td>
      <td width="13"><img src="../../jhimg/jiao_y_x.gif" width="21" height="21"></td>
</tr>
</table>
</div>
</BODY>
</HTML>