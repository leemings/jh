<!--#include file="conn.asp"-->
<!--#include file="z_testconn.asp"-->

<!-- #include file="inc/const.asp" -->
<%
'------------------------
'script writed by ������
'2003-01-06     �ú����
'------------------------
	stats="���Ĵǵ� �û�����"
	
	if (not master) and Request("action")<>"" then
		errmsg=errmsg+"<br><li>��û���ڿ��Ĵǵ�����û���Ȩ�ޣ���ȷ�����Ѿ���¼���Ҿ�����Ӧ��Ȩ�ޣ�"
		founderr=true
		stats="���Ĵǵ� ������Ϣ"
		call nav()
		call head_var(0,0,"���Ĵǵ�","z_test.asp")		
		call dvbbs_error()
	else
		call nav()
		call head_var(0,0,"���Ĵǵ�","z_test.asp")
		dim ars
		select case request("action")
			case "LayIn"
				call LayIn()
			case else		
				call main()
		end select
	end if
	set aconn=nothing
	call activeonline()
	call footer()
'================================================
'-----------------���Ĵǵ��û��б�--------------
sub main()
%>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<tr><th>��ӭ <%=membername%> ���뿪�Ĵǵ� ��������</th></tr>
	<tr><td class=tablebody2> 
    <%call Admin_Head()%>
	
	<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="width:97%"> 
		<tr>
		  <th colspan="3" align=left>����������빦��</th>
		</tr>
		<tr>
			<td colspan="3" align="center" class=tablebody1>ע��:�����水Ҫ����д��������!</td>
		</tr>
		<form action="?action=LayIn" method=post>
		<tr>
			<td width="20%" height="1" align="right" valign="middle" class=tablebody1>Դ���ݿ�����</td>
			<td width="60%" height="1" class=tablebody1><input type="text" name="DBname" size="30"> �磺testchen.mdb����plus/testchen.mdb</td>
			<td width="20%" height="3" rowspan="4" class=tablebody1>
			&nbsp;<font color=blue>��</font> �������ǰ����鿴��������ݿ���������¼��Ŀ�����ı���������д����߱���<br>
			&nbsp;<font color=red>��</font> �������Ƚ�ռ�÷�������Դ�������ڱ��ؽ��У�����ѡ����������ٵ�ʱ����� 
			</td>
		</tr>
		<tr>
			<td class=tablebody1 align="right" valign="middle">Դ��������</td> 
			<td class=tablebody1><input type="text" name="LBTable" size="30" value="TestLb"> �磺TestLb</td>
		</tr>
		<tr>
			<td class=tablebody1 align="right" valign="middle">Դ��Ŀ������</td>
			<td class=tablebody1><input type="text" name="TiMuTable" size="30" value="Test"> �磺Test</td>
		</tr>
		<tr>
			<td class=tablebody1 align="right" valign="middle">����ѡ�</td>
			<td class=tablebody1><input type="radio" name="Setting"  value="0" checked>��ԭ���Ļ�����׷�� &nbsp;<input type="radio" name="Setting"  value="1">ɾ������Ŀ����</td>
		</tr>				
		<tr>
			<td class=tablebody1 align="right" valign="middle"></td>
			<td class=tablebody1 colspan="2"><input type="submit" name="Submit" value="����"></td>
		</tr>
		</form>
	</table>
	<br>
	</td>
	</tr>
</table>
<% 
end sub

sub LayIn()
	dim DBname,LBTable,TiMuTable,Tconn
	DBname=replace(trim(request.Form("DBname")),"'","")
	LBTable=replace(trim(request.Form("LBTable")),"'","")
	TiMuTable=replace(trim(request.Form("TiMuTable")),"'","")
	
	if DBname="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>����д��������ݿ���"
		call test_err()
		exit sub
	end if
	if LBTable="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>����д�����������"
		call test_err()
		exit sub
	end if
	if TiMuTable="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>����д�������Ŀ����"
		call test_err()
		exit sub
	end if

	'ȡ���µ����ݽ��и���
	on error resume next	
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBname)		'���ӵ�������ݿ�
	
	if err.number=0 then
		dim trs,nrs,NewLbID,OldTitleID,OldLbID
		set nrs=tconn.execute("select lb,id from "&LBTable&" order by id")         '�򿪵��������
		
		if cint(request.Form("Setting"))=1 then
			'ȡ��ԭ�������Ŀid�����id
			OldTitleID=aconn.execute("select top 1 id from [test] order by id desc")(0)
			OldLbID=aconn.execute("select top 1 id from [testlb] order by id desc")(0)
		end if  
		if err.number=0 then
		
			do while not nrs.eof 
				aconn.execute("insert into [testlb] (lb) values('"&nrs(0)&"')")
				NewLbID=aconn.execute("select top 1 id from [testlb] order by id desc")(0)
				set trs=tconn.execute("select * from "&TiMuTable&" where lb="&nrs(1))		'�򿪵������Ŀ��
				if err.number<>0 then exit do	
				do while not trs.eof
					aconn.execute("insert into [test] (lb,jj,[time],title,key1,key2,key3,key4,ok,tx,tcount,okcount,username) values ("&NewLbID&","&trs("jj")&","&trs("time")&",'"&replace(trs("title"),"'","��")&"','"&trs("key1")&"','"&trs("key2")&"','"&trs("key3")&"','"&trs("key4")&"','"&trs("ok")&"',"&trs("tx")&","&trs("tcount")&","&trs("okcount")&",'"&trs("username")&"')")
					trs.movenext
				loop
				trs.close
				nrs.movenext
			loop
			nrs.close
			set nrs=nothing		
			set trs=nothing
			
			if cint(request.Form("Setting"))=1 and err.number=0 then
				'ɾ��������Ŀ�����
				aconn.execute("delete from [test] where id<="&OldTitleID&"")    
				aconn.execute("delete from [testlb] where id<="&OldLbID&"") 
				sucmsg=sucmsg+"<br><li>������Ѿ�ɾ��"
			end if
			
		end if
	end if	  
	if err.number<>0 then
		ErrMsg=ErrMsg+"<Br>"+"<li>���ݿ����ʧ�ܣ���ȷ�����������Լ����ݿ�ṹ�Ƿ�һ�£�������Ϣ���£�"
		ErrMsg=ErrMsg+"<Br>"+"<li>"&err.Description
		err.clear
		call test_err()
	else
		dim LBnum,TitleNum,Total,LayInNum
		LBnum=tconn.execute("select count(id) from "&LBTable&"")(0)
		Total=tconn.execute("select count(id) from "&TiMuTable&"")(0)
		LayInNum=tconn.execute("select count(id) from "&TiMuTable&" where lb in (select id from "&LBTable&")")(0)
		
		sucmsg=sucmsg+"<br><li>����⵼��ɹ�"
		sucmsg=sucmsg&"<br><li>���������Դ��� ����<font color=red>"&LBnum&"</font>�����࣬��������<font color=red>"&Total&"</font>�⣬�ɹ����룺<font color=red>"&LayInNum&"</font>�⣬�޷�����(��Щ��Ŀ������һ�����)��<font color=red>"&(Total-LayInNum)&"</font>��"
		call suc()
	end if
	tconn.close
	set tconn=nothing				
end sub
%>