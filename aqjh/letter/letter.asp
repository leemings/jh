<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<%
pagesize=10
page=request.querystring("page")
fs=request.querystring("fs")
if page="" or clng(page)<1 then page=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻�ǽ������ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
myid=rs("id")
rs.close
dim tmprs
sql="select count(*) as ���� from letter where "
sql1="select * from letter where "
if fs="d" then
	sql=sql&"yhid="&myid&" and del=true"
	sql1=sql1&"yhid="&myid&" and del=true order by rq desc"
elseif fs="f" then
	sql=sql&"fjr='"&aqjh_name&"'"
	sql1=sql1&"fjr='"&aqjh_name&"' order by rq desc"
else
	sql=sql&"yhid="&myid&" and del=false"
	sql1=sql1&"yhid="&myid&" and del=false order by rq desc"
end if
tmprs=conn.execute(sql)
xjzs=tmprs("����")
set tmprs=nothing
rs.open sql1,conn,3,3
rs.PageSize = PageSize
pgnum=rs.pagecount
pgnum=clng(pgnum)
if pgnum<1 then pgnum=1
if page>pgnum then page=pgnum
if rs.pagecount>0 then rs.AbsolutePage=page
count=0
if fs="d" then
	xx="����<font color=red>"&xjzs&"</font>�����ű�ɾ��"
elseif fs="f" then
	xx="�����͹�<font color=#FF0000>"&xjzs&"</font>������"
else
	xx="���յ�<font color=#FF0000>"&xjzs&"</font>������"
end if
%>
<HTML>
<HEAD>
<TITLE>���㴫��</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="../jhimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=../jhimg/back002.gif bgColor=#996666 leftMargin=0 oncontextmenu="return false;" text=#cccccc topMargin=0 marginheight="0" marginwidth="0">
<div align=center><br>
<table border="0" cellpadding="0" cellspacing="0" width="544">
<tr> 
<td width="24"><img src="../jhimg/jiao_z_s.gif" width="21" height="21"></td>
<td width="518" background="../jhimg/line_s.gif"><%if aqjh_grade=10 then response.write "<a href=admin/adminletter.asp target=_blank>���Ÿ߼�����</a>"%></td>
<td width="10"><img src="../jhimg/jiao_y_s.gif" width="22" height="21"></td>
</tr>
<tr> 
<td width="24" background="../jhimg/line_z.gif">&nbsp;</td>
<td width="518" align=center> 
        <TABLE bgColor=#675138 width=538 height="195">
          <TBODY> 
          <TR align=middle> 
              <TD height="35" bgcolor="#333300" width="530"> <SPAN class=title><IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17> 
                <%=Application("aqjh_chatroomname")%>���š����㴫�� <IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17></SPAN> 
              </TD>
</TR>
<TR> 
            <TD height="180" bgcolor="#675138" width="530"> 
              <p style=" line-height: 150%" align="center"><a href="letter.asp?fs=s"><img src="../jhimg/m_inbox.gif" width="40" height="40" border="0"></a> 
<a href="letter.asp?fs=d"><img src="../jhimg/M_recycle.gif" width="40" height="40" border="0"></a> 
<a href="write.asp"><img src="../jhimg/m_write.gif" width="40" height="40" border="0"></a> 
<a href="letter.asp?fs=f"><img src="../jhimg/M_issend.gif" width="40" height="40" border="0"></a> <br>
<%=xx%> <%if page>1 then%><a href="letter.asp?page=<%=page-1%>&fs=<%=fs%>">��һҳ</a> <%end if%>
 ��ǰ��<font color="#00FF00"><%=page%></font>ҳ <%if page<pgnum then%><a href="letter.asp?page=<%=page+1%>&fs=<%=fs%>">��һҳ</a> <%end if%>
 ��<font color="#FF0000"><%=pgnum%></font>ҳ [ 
<%for i=1 to pgnum
	if page=pgnum then
		response.write "<b>"&i&"</b>"
	else
		response.write "<a href='letter.asp?page="&i&"&fs="&fs&"'>"&i&"</a>"
	end if
next%>
] 
<DIV align=center> 
<table bordercolorlight="ff6600" align="center" width="522" border="1" cellpadding="2" cellspacing="0">
<tr> 
<td width="258" height="25">����</td>
<td width="72" height="25" align="center"><%if fs<>"f" then%>������<%else%>���͸�<%end if%></td>
<td width="121" height="25" align="center">ʱ��</td>
<%if fs<>"f" then%><td width="61" height="25" align="center">ɾ��</td><%end if%>
</tr>
<%if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing%>
<td colspan="4" height="34">
<%if fs="d" then%>
û�б�ɾ���Ķ��ţ� 
<%elseif fs="f" then%>
û���ѷ��͵��ż���
<%else%>
��û���յ��κζ��ţ� 
<%end if%>
</td>
<%else
	do while not (rs.bof or rs.eof) and count<rs.pagesize
	bt=rs("bt")
	cs=rs("cs")
	if fs<>"f" and not cs then bt="<b>"&bt&"</b>"%>
<tr> 
<td width="258" height="29"><%if fs<>"f" then%><a href="#" onclick="javascript:window.open('sletter.asp?id=<%=rs("id")%>','','scrollbars=yes,resizable=yes,width=550,height=400')"><%end if%><%=bt%><%if fs<>"f" then%></a><%end if%></td>
<td width="72" height="29" align="center">
<%if fs<>"f" then
response.write "<a href='#' onclick=window.open('../yamen/mt.asp?action="&rs("fjr")&"','','scrollbars=yes,toolbar=no,menubar=no,width=460,height=430')>"&rs("fjr")&"</a>"
else
response.write "<a href='#' onclick=window.open('../yamen/mt.asp?action="&rs("sjr")&"','','scrollbars=yes,toolbar=no,menubar=no,width=460,height=430')>"&rs("sjr")&"</a>"
end if%></td>
<td width="121" height="29" align="center"><%=rs("rq")%></td>
<%if fs<>"f" then%>
<td width="61" height="29" align="center"><a href=delletter.asp?id=<%=rs("id")%>>ɾ��</a>
<%if fs="d" then%>
<a href=fyletter.asp?id=<%=rs("id")%>>��ԭ</a>
<%end if%>
</td><%end if%>
</tr>
		<%count=count+1
		rs.movenext
	loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if%>
</table>
<br>
<input type="button" value="�رմ���" onClick="window.close()" name="button">
<br>
</DIV>
</TD>
</TR>
</TBODY> 
</TABLE>
        <!--#include file="bottom.htm"-->
      </td>
<td width="10" background="../jhimg/line_y.gif">&nbsp;</td>
</tr>
<tr> 
<td width="24"><img src="../jhimg/jiao_z_x.gif" width="22" height="21"></td>
<td width="518" background="../jhimg/line_x.gif"></td>
<td width="10"><img src="../jhimg/jiao_y_x.gif" width="21" height="21"></td>
</tr>
</table>
</div>
</BODY>
</HTML>