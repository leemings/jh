<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="z_fy_const.asp"-->
<!--#include file="z_plus_check.asp"-->
<%
'=========================================================
' Plusname: ��Ժ����������
' File: z_fy_fanren.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
call nav()
stats="��������"
call head_var(0,0,"������Ժ","z_fy_fayuan.asp")
if not founduser then
   Errmsg=Errmsg+"<br>"+"<li>��û�н��뱾����������Ȩ�ޡ�"
call dvbbs_error()
else
   call getfyconfig()%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr><th valign=middle colspan=2 align=center height=25><b>��������</b></td></tr>
<tr>
	<td valign=middle class=tablebody1 height=100>
  	<TABLE cellSpacing=0 cellPadding=0 border=0 align=center><TBODY>
  		<TR>
  			<TD align=middle> <IMG src=Images/img_fy/1-1.jpg width=81 height=70></TD>
  			<TD align=middle valign=middle> <IMG src=Images/img_fy/1-2.jpg width=358 height=70></TD>
				<TD align=middle> <IMG src=Images/img_fy/1-3.jpg width=82 height=70></TD>
			</TR>
			<TR>
				<TD align=center valign=middle background=Images/img_fy/2-1.jpg width=81>&nbsp;</TD>
				<TD valign=top bgcolor=#FFFBFF width=358><%
					sql="select username from [user] where lockuser=1"
					set rs=conn.execute(sql)
					dim Numcount,fyrs
					if rs.eof and rs.bof then
					  response.write("<p align=center><br><br><font color=blue>�������ΰ�����,��������û�¸�:)</font><br><br><table><tr>")
					else
						if master then
						 response.write("<p align=center><a href=z_fy_fyaction.asp?faction=allout><font color=red><b>���������¡�</b></font></a><br>")
						end if
						response.write("<p align=center>���������˵�����������������,<br><font color=red><b>�ع�Ѻ�ڴ�!!!�Ծ�Ч��!!!</b></font><br><br>")
						dim unum
						unum=0
						response.write "<table cellSpacing=0 cellPadding=0 border=0 height=300 width=100% align=center><tr height=25>"
						do while not rs.eof
							unum=unum+1
							response.write "<td align=left>��<a href=dispuser.asp?name="&rs(0)&" target=_blank title='����["&rs(0)&"]ѽ��˭�����ң�'><font color=blue>"&left(rs(0),4)&"</font></a>��</td><td align=right>"
							if master then
								response.write "<a href=# title='̽�� ["&rs(0)&"]' onClick=""javascript:window.open('z_fy_tanjian.asp?name="&rs(0)&"','','scrollbars=yes,width=450,height=510,top=0,left=300');"">"
								response.write "<b>̽</b></a> <a href=admin_lockuser.asp?action=lock_3&name="&rs(0)&" title='�ͷ� "&rs(0)&"'><b>��</b></a></td>"
							else
								if tj_set(0)="1" then
									response.write "<a href=# title='̽�� ["&rs(0)&"]' onClick=""javascript:window.open('z_fy_tanjian.asp?name="&rs(0)&"','','scrollbars=yes,width=450,height=510,top=0,left=300');"">"
									response.write "<b>̽</b></a>"
								end if
								if bs_set(0)="1" then  
									response.write " <a href=z_fy_fyaction.asp?faction=baoshi&name="&rs(0)&" title='"
									if jymaster then
										response.write "�ͷ� ["&rs(0)&"]'><b>��</b>"
									else
										response.write "������"&bs_set(1)&"Ԫ���� ["&rs(0)&"]ô��'><b>��</b>"
									end if
									response.write "</a></td>"
								end if
							end if
							if (unum mod 3=0) then response.write "</tr><tr height=25>"
							rs.movenext
						loop
					end if
					response.write "</tr><tr height=*><td colspan=3>&nbsp;</td></tr></table>"
				%></TD>
				<TD align=center valign=middle background=Images/img_fy/2-2.jpg width=82>&nbsp;</TD>
			</TR>
			<TR>
				<TD align=middle> <IMG src=Images/img_fy/3-1.jpg width=81 height=84></TD>
				<TD align=middle> <IMG src=Images/img_fy/3-2.jpg width=358 height=84></TD>
				<TD align=middle> <IMG src=Images/img_fy/3-3.jpg width=82 height=84></TD>
			</TR>
		</TBODY></TABLE>
	</td>
	<td class=tablebody2 width=30% valign=top >�� ��������<%
		if not isnull(jyname) and jyname<>"" and jyname<>"��" then response.write "<a href=dispuser.asp?name="&jyname&" target=_blank>"&jyname&"</a>(<font color=red>����</font>) "
		sql="select username from [user] where usergroupid=1"    
		set rs=conn.execute(sql)    
		do while Not RS.Eof 
			response.write "<a href=dispuser.asp?name="&rs(0)&" target=_blank>"&checkstr(rs(0))&"</a> "	
			rs.movenext
		loop 
		response.write "<br><br>" 
		response.write "<b>��ǰ�����������£�</b><br><br>"   
		if bs_set(0)="1" then
			response.write "<b>��</b> ���ͣ��������ͽ�Ϊ"&bs_set(1)&"Ԫ��<br>"
		else
			response.write "<b>��</b> ���ͣ�������<br>"
		end if
		if tj_set(0)="1" then
			response.write "<b>̽</b> ̽�ࣺ������������Ҫ"&tj_set(1)&"Ԫ��<br><br>"
		else
			response.write "<b>̽</b> ̽�ࣺ������̽��������<br><br>"
		end if
		%><table  cellspacing=1 cellpadding=2 width=100% border=2 >
			<tr  align=center>
				<td height=25 ><font color=#ff0000>�����¼�</font></td>
			</tr>
			<tr bgcolor=#444444>
				<td><MARQUEE onmouseover=this.stop() onmouseout=this.start() scrollAmount=1 scrollDelay=10 direction=up width=100% height=200 align=middle border=0><%
					set fyrs=server.createobject("adodb.recordset")
					sql="select top 20 username,dateandtime,tousername,title,content from log where class='����'  order by dateandtime desc"
					fyrs.open sql,connfy
					do while Not fyRS.Eof
						response.write "&nbsp;<font color=#ffffff>"&fyrs(1)&"</font><br><div align=left>"
						select case trim(fyrs(3))
						case "����"
							response.write "<font color=#FFFFFF>���٣�"&fyrs(4)&"</font>"
						case "̽��" 
							response.write "<font color=#FFFFFF>̽�ࣺ"&fyrs(0)&"̽�� "&fyrs(2)&"����˵:"&fyrs(4)&"</font>"
						case "����"
							response.write "<font color=yellow>������"&fyrs(2)&" "&fyrs(4)&"</font>"
						case "����" 
							response.write "<font color=green>����������"&fyrs(2)&" "&fyrs(4)&"</font>"
						case "����" 
							response.write "<font color=green>������"&fyrs(0)&" "&fyrs(4)&"</font>"
						end select
						response.write "&nbsp;</div><br>"
						fyRS.MoveNext
					Loop
				%></marquee></td>
			</tr>
		</table>
		<br><br>
		<b>�鿴����˵����</b><br>
		<li>��ǰ�������ԣ��¼�������̽��ͼ<br>
		<li><a href=# title='�鿴�ѳ�����Ա����'  onClick="javascript:window.open('z_fy_looktj.asp?action=list','','scrollbars=yes,width=450,height=500,top=10,left=300');"><font color=red>������鿴�ѳ�����Ա���Բ�</font></a><br>
		<li><a href=# title='�鿴�ҵ����Բ�'  onClick="javascript:window.open('z_fy_looktj.asp?name=<%=membername%>','','scrollbars=yes,width=450,height=500,top=10,left=300');"><font color=red>������鿴�������ԣ��¼�����</font></a><br><br><br><br><%
		if qj_set(0)="1" then 
			response.write "<li>����״̬����Ѻ����Ȩ��ûʲô���ϣ�<br>&nbsp;&nbsp;&nbsp;��ʱ���ܱ�����Ŷ~~~"
			if qj_set(1)="1" then response.write "<br>&nbsp;&nbsp;&nbsp;���ٱ��˿�"&qj_set(2)&"������"
		else
			response.write "<li>����״̬����Ѻ����Ȩ���б��ϡ�"
		end if
	%></tr>
</table>
<%end if
call activeonline()
call footer()
%>
