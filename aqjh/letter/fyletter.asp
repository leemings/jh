<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<!--#include file="ffsj.asp"-->
<%
id=request.querystring("id")
call numer(id,"ID����ʹ�����֡�")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	call myclose("��ʾ���㲻�ǽ������ˣ�",1)
end if
myid=rs("id")
rs.close
rs.open "select * from letter where yhid="&myid&" and del=true and id="&id,conn
if rs.bof or rs.eof then
	call myclose("��ʾ����Ļ���վ��û�������ţ�",0)
end if
bt=rs("bt")
conn.execute "update letter set del=false where id="&id
call cls()
%>
<HTML>
<HEAD>
<TITLE>ѩ���ɽ�������ԭ����</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="../jhimg/css.css" rel=stylesheet type=text/css>
</HEAD>
<BODY background=../jhimg/back002.gif bgColor=#996666 leftMargin=0 oncontextmenu="return false;" text=#cccccc topMargin=0 marginheight="0" marginwidth="0">
<div align=center><br>
<table border="0" cellpadding="0" cellspacing="0" width="544">
<tr> 
<td width="24"><img src="../jhimg/jiao_z_s.gif" width="21" height="21"></td>
<td width="518" background="../jhimg/line_s.gif"></td>
<td width="10"><img src="../jhimg/jiao_y_s.gif" width="22" height="21"></td>
</tr>
<tr> 
<td width="24" background="../jhimg/line_z.gif">&nbsp;</td>
<td width="518" align=center> 
<TABLE bgColor=#675138 width=523 height="195">
<TBODY> 
<TR align=middle> 
              <TD height="35" bgcolor="#333300" width="511"> <SPAN class=title><IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17> 
                <%=Application("aqjh_chatroomname")%>���š����㴫�� <IMG align=absMiddle height=18 src="../jhimg/icon0.gif" width=17></SPAN> 
              </TD>
</TR>
<TR> 
<TD height="180" bgcolor="#675138" width="511"> 
<DIV align=center> �ɹ���ԭ�˱���Ϊ��<font color=red><%=bt%></font> �Ķ���Ϣ��<br>
<br>
<br>
<br>
<input type="button" value="�ر�" onClick="window.close()" name="button">
<input type="button" name="Submit3" value="����" onclick='javascript:history.go(-1)'><br>
</DIV>
</TD>
</TR>
</TBODY> 
</TABLE>
<!--#include file="bottom.htm"-->
</td>
<td width="10" background="../jhimg/line_y.gif"></td>
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