<%
	dim culi,daili,zhuangli,chubei,daitian,dmoney
	dim Chen_StartLixiDay		'��ʼ������Ϣ�Ĵ������		��ˮ��ɽ  2002.12.25
	dim Chen_BusinessTimeSlice  '����Ӫҵʱ��				��ˮ��ɽ  2002.12.26
	dim bankstate,bank_setting,bankuser_setting
	dim log_setting	   '���в����¼�,��¼ȫ���¼�,���������¼�����,��¼�����¼�
	dim content

    set rs=server.createobject("adodb.recordset")
    sql="select * from [bankconfig]"
    rs.open sql,conn1,1,3
    culi=rs("savedayli")     '�����Ϣ
    daili=rs("ddayli")       '������Ϣ
    zhuangli=rs("zzli")		 'ת����Ϣ 
    chubei=rs("chubei")		 '���д��� 
    daitian=rs("dkday")		 '������������   
	bankstate=rs("state")    '���е�ǰ״̬   1-Ӫҵ   0-�ر�       ��������� 2002.11.30
	bank_setting=split(rs("bank_setting"),",")    ' ��������� 2002.11.30
	log_setting=split(rs("log_setting"),",")    ' ��������� 2002.11.30
	Chen_StartLixiDay=cint(rs("StartLixiDay"))		' ��ˮ��ɽ  2002.12.25
	Chen_BusinessTimeSlice=split(bank_setting(7),"||")	'����Ӫҵ��ʼʱ��   ��ˮ��ɽ 2002.12.25
    rs.close
	
sub bankhead()
%>
<br>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
	<tr>
		<td align=left valign=middle class=TopLighNav1> <a href=Z_bank.asp><font color=navy>����Ӫҵ����</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_bank.asp?menu=1><font color=navy>������ʮ����</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_bank.asp?menu=2><font color=navy>���ж�ʮ�󴢻�</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_fuwu.asp><font color=navy>��������</font></a></td>
		<%if master then%><td align=right valign=middle class=TopLighNav1> <a href="Z_bank.asp?menu=8"><font color=red>���й�������</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_bank_user.asp><font color=red>��������</font></a>&nbsp;&nbsp;&nbsp;&nbsp;<a href=Z_fuwu.asp?action=admin><font color=red>�����������</font></a><%end if%><%if master or cint(log_setting(6))=1 then%><%if not master then%><td align=center valign=middle class=TopLighNav1><%end if%>&nbsp;&nbsp;&nbsp;<a href=Z_bank.asp?menu=13 title=�鿴���з����¼���¼><font color=blue>�¼�</font><%end if%></a></td> 
	</tr> 
</table>
<br>
<%
end sub
'-------------------------------������Ϣ-------------------------------
sub bank_err() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>���д�����Ϣ</th></tr>
	<tr><td width="100%" class=tablebody1><b>��������Ŀ���ԭ��</b><br><br><li>���Ƿ���ϸ�Ķ���<a href="boardhelp.asp?boardid=<%=boardid%>">�����ļ�</a>����������û�е�¼���߲�����ʹ�õ�ǰ���ܵ�Ȩ��
	<font color=<%=Forum_body(8)%>><%=errmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href="javascript:history.go(-1)">������һҳ</a></td></tr>
</table>
<%
end sub
'--------------------���йرո�ʾ----------------------
sub BankClose()
%>
	<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:97%">
	<tr>
		<th height=25><%=stats%></td>
	</tr>
	<tr>
		<td class=tablebody2 height=1>
			<%call bankhead()%>
			<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
				<tr><th height=25>�����Ѿ�������</th></tr>
				<tr><td width="100%" class=tablebody1>
					�װ��Ŀͻ� <font color=blue><%=membername%></font>�����ã�<br>
					&nbsp;&nbsp;&nbsp;&nbsp;���������Ѿ������ˣ����е�Ӫҵʱ���� <%=Chen_BusinessTimeSlice(0)%>:00��<%=Chen_BusinessTimeSlice(1)%>:00 ������Ӫҵʱ���ڽ����������в�����лл������ 
				</td></tr>
				<tr><td align=center height=26 class=tablebody2><a href="index.asp">������̳</a></td></tr> 
			</table>
		</td>
	</tr>
	</table>		
<%
end sub
'-------------------------------�ɹ���Ϣ-------------------------------
sub bank_suc() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>���в����ɹ�</th></tr> 
	<tr><td width="100%" class=tablebody1><b>�����ɹ���</b><br><br><li>��ӭ����<%=Forum_info(0)%>��������
	<font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%=Request.ServerVariables("HTTP_REFERER")%>>������һҳ</a></td></tr>
</table>
<%
end sub

sub logs(lclass,title,userN)
'������ 2002.12.01 д���¼�
	if UserID="" or (not isnumeric(UserID)) then UserID=0
	if content="" then content="���ɵĲ���"
	conn1.execute("insert into log ([class],UserName,UserID,Title,content,DateAndTime,IP) values('"&lclass&"','"&userN&"',"&UserID&",'"&Title&"','"&content&"',now(),'"&Request.ServerVariables("REMOTE_ADDR")&"')")
end sub
%>

