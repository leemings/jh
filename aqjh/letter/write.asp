<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<%
id=request("id")
newbt=""
newxj=""
newsjrq=""
a=""
b=""
newsjr=""
if id<>"" then
	if not isnumeric(id) then
		Response.Write "<script Language=Javascript>alert('��ʾ��ID����');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "select id from �û� where ����='"&aqjh_name&"'",conn
	myid=rs("id")
	rs.close
	rs.open "select * from letter where yhid="&myid&" and id="&id&" and del=false",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ����û���յ��������ţ�');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	newsjrq=rs("rq")
	newbt="�ظ���"&rs("bt")
	a="������"&rs("fjr")&"��"&rs("rq")&"������ż���д����������"&chr(13)
	b="====================================="
	newxj=rs("xj")&chr(13)
	newsjr=rs("fjr")
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
%>
<HTML><HEAD><TITLE>���㴫��</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="../jhimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=../jhimg/back002.gif bgColor=#7c6244 leftMargin=0 oncontextmenu="return false;" text=#cccccc topMargin=0 marginheight="0" marginwidth="0">
<div align=center><br>
<table border="0" cellpadding="0" cellspacing="0" width="504">
<tr>
<td width="22"><img src="../jhimg/jiao_z_s.gif" width="21" height="21"></td>
<td width="468" background="../jhimg/line_s.gif"></td>
<td width="14"><img src="../jhimg/jiao_y_s.gif" width="22" height="21"></td>
</tr>
<tr>
<td width="22" background="../jhimg/line_z.gif">&nbsp;</td>
<td width="468" align=center> 
<TABLE bgColor=#675138 border=0 cellPadding=5 cellSpacing=1 width=521 height="195" style="border-left: 1 solid #816445; border-right: 1 solid #A98761; border-top: 1 solid #A98761">
<TBODY> 
<TR align=middle> 
<TD height="15" bgcolor="#333300" colspan="2" width="445"> <SPAN class=title><IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17> 
���Ͷ���Ϣ <IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17></SPAN>  
</TD>
</TR>
<form name="form1" method="post" action="writeok.asp">
<TR> 
<TD height="41" bgcolor="#675138" width="82" style="border: 1 solid #7C6244">
<p style="text-indent: 19; line-height: 150%" align="center">�ռ���</TD>
<TD height="41" bgcolor="#675138" width="407" style="border: 1 solid #7C6244">
<input type="text" name="name" value="<%=newsjr%>" size="10" maxlength="5">
</TD>
</TR>
<TR> 
<TD height="41" bgcolor="#675138" width="82" style="border: 1 solid #7C6244"> 
<p style="text-indent: 19; line-height: 150%" align="center">��&nbsp;&nbsp;��</TD>
<TD height="41" bgcolor="#675138" width="407" style="border: 1 solid #7C6244"> 
<input type="text" name="bt" size="30" maxlength="20" value="<%=newbt%>">
</TD>
</TR>
<TR> 
<TD height="41" bgcolor="#675138" width="82" valign="top" style="border: 1 solid #7C6244"> 
<p style="text-indent: 19; line-height: 150%" align="center">��&nbsp;&nbsp;��
</TD>
<TD height="41" bgcolor="#675138" width="407" style="border: 1 solid #7C6244"> 
<textarea name="txt" cols="55" rows="8"><%=a%>
<%=newxj%>
<%=b%></textarea>
</TD>
</TR>
<TR> 
<TD height="90" bgcolor="#675138" colspan="2" width="445"> 
<div align="left">˵����<br>
�� ������ʹ��Ctrl+Enter����ݷ��Ͷ���<br>
                  �� �������20���ַ����������200���ַ�<br>
�� �ڷ����������뾡��ʹ��ȫ���ַ���������Щ��ǵı����Żᱻ�Զ�����<br>
�� ÿ��һ�������շ�1���<br>
</div> 
<br>
<div align="center">
<input type="submit" name="Submit" value="����">
<input type="reset" name="Submit2" value="���">
<input type="button" name="Submit3" value="����" onclick='javascript:history.go(-1)'>
<input type="button" value="�ر�" onClick="window.close()" name="button">
<br>
</div>
</TD>
</TR>
</form>
</TBODY> 
</TABLE>
<!--#include file="bottom.htm"-->
</td>
<td width="14" background="../jhimg/line_y.gif">&nbsp;</td>
</tr>
<tr>
<td width="22"><img src="../jhimg/jiao_z_x.gif" width="22" height="21"></td>
<td width="468" background="../jhimg/line_x.gif"></td>
<td width="14"><img src="../jhimg/jiao_y_x.gif" width="21" height="21"></td>
</tr>
</table>
</div>
</BODY></HTML>
