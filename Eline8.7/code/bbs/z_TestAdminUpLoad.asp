<!--#include file="conn.asp"-->
<!--#include file="z_testconn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'--------------------
'edit by ��ˮ��ɽ
'--------------------

	dim ars
	
	stats="���Ĵǵ� ����ϴ���Ŀ"
	
	if not master then
		errmsg=errmsg+"<br><li>��û���ڿ��Ĵǵ�����ϴ����Ȩ�ޣ���ȷ�����Ѿ���¼���Ҿ�����Ӧ��Ȩ�ޣ�"
		founderr=true
		stats="���Ĵǵ� ������Ϣ"
		call nav()
		call head_var(0,0,"���Ĵǵ�","z_test.asp")		
		call dvbbs_error()
	else
		call nav()
		call head_var(0,0,"���Ĵǵ�","z_test.asp")
		
		select case request("action")
			case "Auditing"
				call Auditing()
			case "ͨ�����"
				call pass()
			case "ɾ������"
				call del()
			case "passall"
				call passall()
			case "delall"
				call delall()
			case else		
				call UploadList()
		end select
	end if
	set aconn=nothing
	call activeonline()
	call footer()
'================================================
'-----------------�ϴ���Ŀ�б�--------------
sub UploadList()
%>
<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<tr><th>��ӭ <%=membername%> ���뿪�Ĵǵ� ��������</th></tr>
	<tr><td class=tablebody2>
	<%call Admin_Head()%>
		<table align=center cellspacing=1 cellpadding=3 class=tableborder1 style="width:97%">
		  <tr><th colspan="7">�ϴ���Ŀ����</th></tr> 	
		  <tr>
			<td width="60" align="center" class=TopLighNav1><b>�⡡��</b></td>
			<td width="*" align="center" class=TopLighNav1><b>�⡡Ŀ</b></td>
			<td width="100" align="center" class=TopLighNav1><b>��Ŀ���</b></td>
			<td width="50" align="center" class=TopLighNav1 ><b>����</b></td>
			<td width="90" align="center" class=TopLighNav1><b>�����ϴ���</b></td>
			<td width="65" align="center" class=TopLighNav1><b>�ϴ�ʱ��</b></td> 
			<td width="130" align="center" class=TopLighNav1><b>����</b></td> 
		  </tr>
<%
			dim HaveUpLoad
			HaveUpLoad=true

			set ars=aconn.execute("select * from [testtemp] order by AddTime desc")
			if ars.eof and ars.bof then
				HaveUpLoad=false
				response.Write "<tr><td colspan=7 class=tablebody1>��ʱû���ϴ���Ŀ��Ҫ���</td></tr>"
			else
				do while not ars.eof
%>
				  <tr>
					<td align="center" class=tablebody2><%=ars("id")%></td>
					<td class=tablebody1><a href="?action=Auditing&lbid=<%=ars("lb")%>&id=<%=ars("id")%>"><%=ars("title")%></a></td>
					<td class=tablebody1 align="center"><%=GetLbName(ars("lb"))%></td>
					<td class=tablebody1 align="center"><%if ars("tx")=1 then%>�����<%else%>ѡ����<%end if%></td> 
					<td class=tablebody1 align="center"><%=ars("username")%></td>
					<td class=tablebody1 align="center"><%=ars("addtime")%></td> 
					<td class=tablebody1 align="center"><a href="?action=Auditing&lbid=<%=ars("lb")%>&id=<%=ars("id")%>" title="�������">���</a> | <a href="?action=ͨ�����&flag=auto&id=<%=ars("id")%>" title="ֱ��ͨ�����" onclick="javascript:{if(confirm('��ȷ�����鿴������Ŀ���ݾ�ͨ������������?')){return true;}return false;}">ͨ��</a> | <a href="?action=ɾ������&id=<%=ars("id")%>" title="ֱ��ɾ������" onclick="javascript:{if(confirm('��ȷ��ɾ��������?')){return true;}return false;}">ɾ��</a></td>
				  </tr>
<%				
					ars.movenext
				loop	
			end if
			ars.close
			set ars=nothing	
%>		  
		</table> 
		<br>
	</td></tr>	
	<tr><td class=tablebody1 align="center"><%if HaveUpLoad then%><a href="?action=passall" title="ֱ��ͨ�����д���˵���Ŀ" onclick="javascript:{if(confirm('��ȷ�����鿴������Ŀ���ݾ�ͨ�����д���˵���Ŀ��?')){return true;}return false;}">ȫ��ͨ��</a> | <a href="?action=delall" title="ֱ��ɾ�����д���˵���Ŀ" onclick="javascript:{if(confirm('��ȷ��ɾ�����д���˵���Ŀ��?')){return true;}return false;}">ȫ��ɾ��</a> | <%end if%><a href="Z_test.asp">��  ��</a></td></tr>
</table>	
<%	
end sub
'------------�����Ŀ--------------
sub Auditing()
	dim id,lbid
	id=request("id")
	if not isInteger(id) then
		errmsg=errmsg+"<br><li>�������Ŀ����"
		call test_err()
		exit sub
	end if

	sql="select * from [testtemp] where id="&id
	set ars=aconn.execute(sql)
	if ars.eof and not ars.bof then
		errmsg=errmsg+"<br><li>û���ҵ���Ӧ����Ŀ"
		founderr=true
		ars.close
		call test_err()
		exit sub
	end if
	lbid=request("lbid")
%>
	<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<form method="POST" action="Z_TestAdminUpLoad.asp">
	<tr><th>��ӭ <%=membername%> ���뿪�Ĵǵ� ��������</th></tr>
	<tr><td class=tablebody2>
<%call Admin_Head()%>
<table align=center cellspacing=1 cellpadding=3 class=tableborder1 style="width:97%">
  <tr><th colspan=8>����ϴ���Ŀ</th></tr> 
  <tr>
	<td width="70" align="center" class=tablebody2><b>��ʱ���</b></td>
	<td width="82" class=tablebody1>&nbsp;<%=ars("id")%></td>
	<td width="98" align="center" class=tablebody2><b>��Ŀ�������</b></td>
	<td width="102" class=tablebody1><select name="lb"><%
	set rs=aconn.execute("select lb,id from [testlb]")
	do while not rs.eof
%>
		<option value=<%=rs(1)%> <%if cint(rs(1))=cint(lbid) then%> selected <%end if%>><%=rs(0)%></option>
<%
  		rs.movenext
	loop
	rs.close
	set rs=nothing
%></select></td>
	<td width="57" align="center" class=tablebody2 ><b>����</b></td>
	<td width="63" class=tablebody1><%if ars("tx")=1 then%>�����<%else%>ѡ����<%end if%></td>
	<td width="90" align="center" class=tablebody2><b>�����ϴ���</b></td>
    <td width="102" class=tablebody1><%=ars("username")%></td>
  </tr>
  <tr>
    <td width="70" class=tablebody2 align="center"><b>��&nbsp;&nbsp;Ŀ</b></td>
    <td width="600" colspan="7" align="left" class=tablebody1>&nbsp;<textarea rows="12" name="title" cols="70"><%=ars("title")%></textarea>
	</td>
  </tr>
  <tr>
    <td width="70"  align="center" class=tablebody2><%if ars("tx")=1 then%>��ȷ�𰸣�<%else%>�� 1��<%end if%></td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key1" size="70" value='<%=ars("key1")%>'></td>
  </tr>
 <%if ars("tx")=0 then%>
  <tr>
    <td width="70" align="center" class=tablebody2>�� 2��</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key2" size="70"  value='<%=ars("key2")%>'></td>
  </tr>
  <tr>
    <td width="70" align="center" class=tablebody2>�� 3��</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key3" size="70" value='<%=ars("key3")%>'></td>
  </tr>
  <tr>
    <td width="70"  align="center" class=tablebody2>�� 4��</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key4" size="70"  value='<%=ars("key4")%>'></td>
  </tr>
 <%end if%>
  <tr>
    <td width="70" class=tablebody2 align="center"><%if ars("tx")=0 then%>��ȷ�𰸴��ţ�<%end if%></td>
    <td width="81" class=tablebody1>&nbsp;<%if ars("tx")=0 then%><input type="text" name="ok" size="10"  value=<%=ars("ok")%>><%end if%></td>
    <td width="116" colspan="2" class=tablebody2> <b><span lang="zh-cn">��ʱ</span></b>
      <input type="text" name="atime" size="7" value=<%=ars("timelimit")%>> <b><span lang="zh-cn"> ��</span></td>
    <td width="172" colspan="2" class=tablebody2>
    <p align="center"><b>���⽱�� <input type="text" name="jj" size="7" value=<%=ars("jiangjin")%>> Ԫ</b></td>
    <td width="231" colspan="2" class=tablebody1>
		<input type=hidden name="id" value=<%=ars("id")%>>
		<input type=hidden name="tx" value=<%=ars("tx")%>>
		<input type=hidden name="username" value=<%=ars("username")%>>
	</td>
  </tr>
</table><br>
<tr><td align="center" class=tablebody2>
	<input type=submit value="ͨ�����" name=action>����<input type=submit value="ɾ������" name=action>����<input type="submit" value="������Ŀ�б�" name="go">
</td></tr>
</form></table>   
<%
		set ars=nothing	
end sub

sub pass()
'write by ������ 2003.1.1
'ͨ�����
	dim ausername
	dim id,lb,key(4),jj,atime,title,ok,tx
	id=request("id")
	if not isInteger(id) then
		errmsg=errmsg+"<br><li>�������Ŀ����"
		call test_err()
		exit sub
	end if
	if request("flag")="auto" then		'ֱ�������Ŀ
		set ars=aconn.execute("select * from [testtemp] where id="&id)
		if ars.eof and ars.bof then
			errmsg=errmsg+"<br><li>�������Ŀ������û���ҵ�����Ŀ"
			call test_err()
			ars.close
			exit sub
		end if
		lb=ars("lb"):		jj=ars("JiangJin"):		atime=ars("TimeLimit"):		title=ars("title")
		key(0)=ars("key1"):		key(1) =ars("key2"):		key(2) =ars("key3"):		key(3) =ars("key4")
		ok=ars("ok"):		tx=ars("tx"):		ausername=ars("username")
		ars.close
	else		'���������Ŀ
		lb=request("lb"):		jj=request("jj"):		atime=request("atime"):		title=request("title")
		key(0)=request("key1"):		key(1) =request("key2"):		key(2) =request("key3"):		key(3) =request("key4")
		ok=request("ok"):		tx=request("tx"):		ausername=trim(replace(request("username"),"'",""))	
	end if	
	if lb="" or (not isnumeric(lb)) then
		errmsg=errmsg+"<br><li>�������Ŀ�����������Ч���ӽ���"
		founderr=true
	end if
	if request("flag")<>"auto" then		'���������Ŀ
		if jj="" or (not isnumeric(jj)) then
			errmsg=errmsg+"<br><li>�������������"
			founderr=true
		end if
		if atime="" or (not isnumeric(atime)) then
			errmsg=errmsg+"<br><li>ʱ�ޱ���������"
			founderr=true
		elseif atime<=0 then
			errmsg=errmsg+"<br><li>ʱ�ޱ������0"
			founderr=true			
		end if
		if cint(tx)=0 then		'��ѡ������жϴ𰸴��룬����ⲻ��
			if ok="" or (not isnumeric(ok)) then
				errmsg=errmsg+"<br><li>����д��ȷ�𰸵Ĵ���"
				founderr=true
			elseif ok>4 or ok<1 then
				errmsg=errmsg+"<br><li>��ȷ�𰸵Ĵ���ֻ����д1��4������"
				founderr=true			
			end if
		end if
		if key(0)="" then 
			errmsg=errmsg+"<br><li>��һ����ѡ�𰸲���Ϊ��"
			founderr=true
		end if 							
		if title="" then
			errmsg=errmsg+"<br><li>��Ŀ����Ϊ��"
			founderr=true		
		end if
	end if
	if ausername="" then ausername="����Ա����"
	if founderr then 
		call test_err()
		exit sub
	end if	
	set ars=server.createobject("adodb.recordset")	
	sql="select * from [test] where lb="&lb
	ars.open sql,aconn,1,3
	
	ars.addnew 
	ars("lb")=lb
	ars("jj")=jj
	ars("time")=atime
	ars("title")=title
	ars("key1")=key(0)
	ars("key2")=key(1) 
	ars("key3")=key(2) 
	ars("key4")=key(3) 
	ars("ok")=ok
	ars("tx")=tx
	ars("username")=ausername
	ars.update 
	ars.close 
	
	'ɾ����ʱ���е���Ŀ
	aconn.execute("delete from [testtemp] where id="&id)
	
	if ausername<>"����Ա����" then
		'ȡ���ϴ���Ŀ�Ľ������Ϸ�� 
		dim jiangjin,auserupsinew
		set ars=aconn.execute("select userjl,userupsinew from [config]")
		jiangjin=ars(0)
		auserupsinew=ars(1)
		'�ѽ������Ϸ�Ҹ����ϴ��� 
		conn.execute ("update [user] set userWealth=userWealth+"&jiangjin&" where username='"&ausername&"'")   
		'�����ϴ��û����ϴ���Ŀͨ��������Ϸ��
		aconn.execute ("update testuser set userpass=userpass+1,usersinew=usersinew+"&auserupsinew&" where username='"&ausername&"'")
	
		'���ϴ��߷��Ͷ���Ϣ֪ͨ
		dim message
		message="���ϴ���֪ʶ�ʴ���Ŀ��"&request("title")&" ����ͨ����ˣ�ͬʱ����������"&jiangjin&"Ԫ����Ԫ��"&auserupsinew&"����Ϸ�ң�����գ�"
		sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&ausername&"','"&membername&"','����֪ʶ�ʴ���Ŀ�ɹ���','"&message&"',Now(),0,1)"
		conn.execute sql
	end if
	
	set ars=nothing	
	
	sucmsg=sucmsg+"<br><li>���������ɣ��뷵�ؽ�����������"
	rUrl="Z_TestAdminUpLoad.asp"
	call suc()
end sub

sub passall()
'write by ������  2003.1.1
'ȫ��ͨ�����
	dim jiangjin,auserupsinew,message

	'ȡ���ϴ���Ŀ�Ľ������Ϸ�� 
	set ars=aconn.execute("select userjl,userupsinew from [config]")
	jiangjin=ars(0)
	auserupsinew=ars(1)
	
	set ars=aconn.execute("select * from [testtemp]")
	if not (ars.eof and ars.bof) then 
		do while not ars.eof
			'����Ŀ������������ 
			aconn.execute("insert into [test] (lb,jj,[time],title,key1,key2,key3,key4,ok,tx,username) values('"&ars("lb")&"',"&ars("JiangJin")&","&ars("TimeLimit")&",'"&ars("title")&"','"&ars("key1")&"','"&ars("key2")&"','"&ars("key3")&"','"&ars("key4")&"','"&ars("ok")&"',"&ars("tx")&",'"&ars("username")&"')") 
			'�ѽ������Ϸ�Ҹ����ϴ���  
			conn.execute ("update [user] set userWealth=userWealth+"&jiangjin&" where username='"&ars("username")&"'")   
			'�����ϴ��û����ϴ���Ŀͨ��������Ϸ��
			aconn.execute ("update testuser set userpass=userpass+1,usersinew=usersinew+"&auserupsinew&" where username='"&ars("username")&"'")
			'���ϴ��߷��Ͷ���Ϣ֪ͨ
			message="��¼���֪ʶ�ʴ���Ŀ��"&ars("title")&"��û��ͨ����ˣ��Ѵ���ʱ����ɾ������ע�⣡"
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&ars("username")&"','"&membername&"','֪ʶ�ʴ���Ŀɾ��֪ͨ��','"&message&"',Now(),0,1)"
			conn.execute sql
			ars.movenext
		loop 
		'ɾ����ʱ���е���Ŀ
		aconn.execute("delete from [testtemp]") 
	end if
	ars.close
	set ars=nothing
	sucmsg=sucmsg+"<br><li>������Ŀ�����ɣ��뷵�ؽ�����������"
	rUrl="Z_TestAdminUpLoad.asp"
	call suc()	
end sub

sub del()
'write by ������  2003.1.1
'û��ͨ�����,��Ŀ����ʱ����ɾ��
	dim ausername,message
	ausername=replace(request("username"),"'","")
	'ɾ����ʱ���е���Ŀ
	aconn.execute("delete from [testtemp] where id="&request("id"))
	
	message="��¼���֪ʶ�ʴ���Ŀ��"&request("title")&" ��û��ͨ����ˣ��Ѵ���ʱ����ɾ��"
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&ausername&"','"&membername&"','֪ʶ�ʴ���Ŀɾ��֪ͨ��','"&message&"',Now(),0,1)"
	conn.execute sql
	
	sucmsg=sucmsg+"<br><li>��Ŀ�Ѵ���ʱ����ɾ�����뷵�ؽ�����������"
	rUrl="Z_TestAdminUpLoad.asp"
	call suc()	
end sub

sub delall()
'write by ������ 2003.1.1
'����ʱ����ɾ�����д���˵��ϴ�����Ŀ
	dim message 
	
	set ars=aconn.execute("select title,username from [testtemp]")
	if not (ars.eof and ars.bof) then 
		do while not ars.eof
			message="��¼���֪ʶ�ʴ���Ŀ��"&ars(0)&"��û��ͨ����ˣ��Ѵ���ʱ����ɾ��"
			sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&ars(1)&"','"&membername&"','֪ʶ�ʴ���Ŀɾ��֪ͨ��','"&message&"',Now(),0,1)"
			conn.execute sql
			ars.movenext
		loop 
		'ɾ����ʱ���е���Ŀ
		aconn.execute("delete from [testtemp]")
	end if
	ars.close	
	set ars=nothing	
	sucmsg=sucmsg+"<br><li>���д���˵��ϴ�����Ŀ�Ѵ���ʱ����ɾ�����뷵�ؽ�����������"
	rUrl="Z_TestAdminUpLoad.asp"
	call suc()
end sub
%>