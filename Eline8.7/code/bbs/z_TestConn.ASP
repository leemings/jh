<%
	dim aconn,rUrl
	dim aconnstr
	const TestDb="data/e_testchen.asp"		'�޸����ݿ�·�������ݿ����ƣ����·����
	Set aconn = Server.CreateObject("ADODB.Connection")
	aconnstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(TestDb)
	'�����ķ��������ý��ϰ汾Access�����������������ӷ���
	'aconnstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(TestDb)
	
	aconn.Open aconnstr


'==========ͨ�ú���========================
function GetLbName(Lid)
	dim grs
	set grs=aconn.execute("Select lb From [testlb] where id="&Lid)
	if grs.eof then
		GetLbName="<font color=red>��ɾ��</font>"
	else
		GetLbName=grs(0)
	end if 
end function
'======================
sub suc()
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>���Ĵǵ� �ɹ���Ϣ</th></tr>
	<tr><td width="100%" class=tablebody1><b>�����ɹ���</b><br><br><li>��ӭ����<%=Forum_info(0)%>���Ĵǵ�
	<font color=navy><%=sucmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href=<%if rUrl="" or isnull(rUrl) then%><%=Request.ServerVariables("HTTP_REFERER")%><%else%><%=rUrl%><%end if%>>��&nbsp;&nbsp;��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=z_test.asp>���Ĵǵ�</a></td></tr>
</table>
<%
end sub

sub test_err() 
%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:75%">
	<tr><th height=25>���Ĵǵ� ������Ϣ</th></tr>
	<tr><td width="100%" class=tablebody1><b>��������Ŀ���ԭ��</b><br><br><li>���Ƿ���ϸ�Ķ���<a href="Z_TestHelp.asp">�����ļ�</a>����������û�е�½���߲�����ʹ�õ�ǰ���ܵ�Ȩ��
	<font color=<%=Forum_body(8)%>><%=errmsg%></font><br></td></tr>
	<tr><td align=center height=26 class=tablebody2><a href="javascript:history.go(-1)">������һҳ</a></td></tr>
</table>
<%
end sub

sub Admin_Head()
%>
<br>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
	<tr>
		<td align=left valign=middle class=TopLighNav1> &nbsp;<a href="Z_test.asp" title="���ؿ��Ĵǵ���ҳ">���Ĵǵ�</a> | <a href="z_testSetting.asp" title="�ǵ�������á��ǵ�������">�ǵ����</a> | <a href="Z_TestAdminUpLoad.asp" title="����ϴ�����Ŀ">�ϴ�����</a> | <a href="Z_TestEdit.asp" title="�����������">������</a>  | <a href="Z_TestAdminUser.asp" title="�����ǵ��û�">�û�����</a> | <a href="Z_TestLayIn.asp" title="�����������">�����������</a> | <a href=javascript:history.go(-1)>������һҳ</a></td>
	</tr> 
</table>
<br>
<%
end sub

sub User_Head()
%>
<br>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%">
	<tr>
		<td align=left valign=middle class=TopLighNav1> &nbsp;<a href="Z_test.asp">���Ĵǵ�</a> | <a href="Z_test.asp?action=PaiXing">����������</a> | <span lang="zh-cn"><a href="Z_TestUpload.ASP">����ѡ����</a></span> | <span lang="zh-cn"><a href="Z_TestUpload.ASP?action=addtk">���������</a></span> | <a href=javascript:history.go(-1)>������һҳ</a></td>
	</tr> 
</table>
<br>
<%
end sub
'===================
sub NewsLog(lclass,content,userN)
'������ 2002.12.01 д���¼�
	if UserID="" or (not isnumeric(UserID)) then UserID=0 
	if content="" then content="δ֪�Ĳ���"
	aconn.execute("insert into testnews ([class],UserName,UserID,content,AddTime,IP) values('"&lclass&"','"&userN&"',"&UserID&",'"&content&"',now(),'"&Request.ServerVariables("REMOTE_ADDR")&"')")
end sub
%>