<!--#include file="conn.asp"-->
<!--#include file="Z_TestConn.ASP"-->
<!-- #include file="inc/const.asp" -->
<%
REM ���Ĵǵ� �������ά��
REM ������ �޸�Ϊfinal�� 
REM ԭ���߲����

dim ars 

if not master then  
	errmsg=errmsg+"<br><li>��û���ڿ��Ĵǵ�ά����������Ȩ�ޣ���ȷ�����Ѿ���¼���Ҿ�����Ӧ��Ȩ�ޣ�" 
	founderr=true
	stats="���Ĵǵ� ������Ϣ"
	call nav()
	call head_var(2,0,"","")		
	call dvbbs_error()
else
	stats="���Ĵǵ� ����������"
	call nav()
	call head_var(0,0,"���Ĵǵ�","Z_test.asp")
%>
	<table align=center cellspacing=1 cellpadding="3" border=0 class=tableborder1>
	<tr><th>��ӭ <%=membername%> ���뿪�Ĵǵ� ��������</th></tr>
	<tr><td class=tablebody2>
<%
	call Admin_Head()
	select case request("action")
		case "saveedit"
			call saveedit()
		case "edit"
			call edit()
		case "del"
			call del()
		case "selaction"
			call selaction()	'��������	
		case else
			call main()
	end select
	response.Write "</td></tr></table>"
	set aconn=nothing	
end if
call activeonline()
call footer()	
'===================================	
sub main()						
%>
<table align=center cellspacing=1 cellpadding=3 class=tableborder1 style="width:97%">
	<tr><th colspan="6">����������</th></tr> 
	<tr><td colspan="6" class=TopLighNav1><font color=<%=Forum_body(8)%>>��������</font></td></tr> 
	<form method="POST" action="?action=edit" name=form1>
	<tr>
		<td class=tablebody1 colspan="2">&nbsp;���: </td>
		<td class=tablebody1 colspan="4"><input type="text" name="Tid" size="20">&nbsp;&nbsp;<input type=submit value="����" name="B1"></td>
	</tr>
	</form>
	<tr><td colspan=6 class=tablebody2 height=5></td></tr>
<%if request("action")="" or request("userSearch")="" then%>	
	<tr><td colspan="6" class=TopLighNav1><font color=<%=Forum_body(8)%>>�߼���ѯ</font></td></tr> 
	<form method="POST" action="?action=userSearch" name=form2>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;�ϴ���</td> 
		<td class=tablebody1 width=80%  colspan="4"><input type="text" name="username" size="20">&nbsp;<input type=checkbox name="usernamechk" value="yes" checked>�û�������ƥ��</td>
	</tr>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;��Ŀ����</td> 
		<td class=tablebody1 width=80%  colspan="4">
			<select size=1 name="TitleClass">
				<option value=0>�Σ���</option> 
				<option value=1>ѡ����</option>
				<option value=2>�����</option>
			</select>	
		</td>
	</tr>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;��Ŀ���</td> 
		<td class=tablebody1 width=80%  colspan="4">
			<select size=1 name="TitleLb">
				<option value=0>�Σ���</option> 
<%
				set ars=aconn.execute("select id,lb from testlb order by id")
				do while not ars.eof
					response.write "<option value="&ars(0)&">"&ars(1)&"</option>"
					ars.movenext
				loop
				ars.close
				set ars=nothing
%>
			</select>	
		</td>
	</tr>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;��Ŀ�а���</td> 
		<td class=tablebody1 width=80%  colspan="4"><input type=text name="Title" size=45></td>
	</tr>
	<tr>
		<td class=tablebody1 width=20%  colspan="2">&nbsp;����ʽ</td> 
		<td class=tablebody1 width=80%  colspan="4">
			<select size=1 name="ordername">
				<option value="id">�����</option> 
				<option value="title">����Ŀ</option>
				<option value="lb">�����</option>
				<option value="username">���ϴ���</option>
				<option value="tcount">��Ӧ������</option>
				<option value="okcount">���������</option>
				<option value="jj">������Ķ���</option>  
				<option value="time">��ʱ�����ƵĶ���</option> 
			</select>
			<select size=1 name="order">
				<option value="esc">����</option> 
				<option value="desc">����</option>
			</select>					
		</td>
	</tr>
	<tr>
		<td class=tablebody2 align=center colspan=6><input name="submit" type=submit value="   ��  ��   "></td>
	</tr>
	<input type=hidden name="userSearch" value="2">
	</form>
	
	<tr><td colspan=6 class=tablebody1 height=5></td></tr>
	<tr><td colspan="6" class=TopLighNav1><font color=<%=Forum_body(8)%>>����Զ�����</font>�����Ҳ�ɾ���ظ�����Ŀ</td></tr>  
	<form method="POST" action="?action=userSearch" name=form3>
	<tr>
		<td class=tablebody1 colspan="2">&nbsp;����˵��: </td>
		<td class=tablebody1 colspan="4">
			&nbsp;�� ������������ظ�����Ŀɾ����ֻ��������һ�⣬ɾ������Ŀ�������ɾ��ѡ�����<br>
			&nbsp;�� ������ֻ���ҳ���Ŀ��ȫһ�µ��ظ���Ŀ<br>
		 	&nbsp;�� �����Ŀ���е����ţ����򽫻���������Ŀ��������Щ���е����ŵ��ظ���Ŀֻ���ֶ�ɾ��<br>
			&nbsp;�� <font color=<%=Forum_body(8)%>>��������������ռ�÷�������Դ��CPUռ100%�����Բ�Ҫ�ڷ�������ֱ�ӽ��б�����</font><br>
			&nbsp;�� ���������ܻỨ�ѱȽϳ���ʱ�䣬����Ҫ�����Ĵ�С����Ϊ�������Ļ���������<br>   
			&nbsp;�� <font color=blue>�����ظ���Ŀ�������⣬������<u>�߼���ѯ</u>��ѡ��<u>��Ŀ</u>�����ҳ��ظ�����Ŀɾ��</font>
		</td> 
	</tr> 
	<tr>
		<td class=tablebody1 colspan="2">&nbsp;ɾ��ѡ��: </td>
		<td class=tablebody1 colspan="4"><input type=radio name="DeleteSel" value="1" checked>ɾ����ĿID�����Ŀ	<input type=radio name="DeleteSel" value="2">ɾ����ĿIDС����Ŀ</td>
	</tr> 
	<tr>
		<td class=tablebody2 align=center colspan=6><input name="submit" type=submit value="��ʼ����"></td>
	</tr>	
	<input type=hidden name="userSearch" value="1">
	</form>	
<%
elseif request("userSearch")=1 then
	call ClearTitle()	'�����ظ����� 
elseif request("userSearch")=2 then
%>
	<tr>
		<td colspan=6 class=TopLighNav1 height=23><font color=<%=Forum_body(8)%>>�������</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="?" class="cblue">�߼���ѯ</a></td>  
	</tr>
	<tr><td colspan=6 class=tablebody2 height=5></td></tr>	
<%
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	dim sqlstr
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	
	Set rs= Server.CreateObject("ADODB.Recordset")
	sqlstr=""
	if trim(request("username"))<>"" then
		if request("usernamechk")="yes" then
			sqlstr=" username='"&trim(request("username"))&"'"
		else
			sqlstr=" username like '%"&trim(request("username"))&"%'"
		end if
	end if
	if cint(request("TitleClass"))>0 then
		if sqlstr="" then
			sqlstr=" tx="&request("TitleClass")-1&""
		else
			sqlstr=sqlstr & " and tx="&request("TitleClass")-1&""
		end if
	end if
	if cint(request("TitleLb"))>0 then
		if sqlstr="" then
			sqlstr=" lb="&request("TitleLb")&""
		else
			sqlstr=sqlstr & " and lb="&request("TitleLb")&""
		end if
	end if	
	if request("Title")<>"" then
		if sqlstr="" then
			sqlstr=" Title like '%"&request("Title")&"%'"
		else
			sqlstr=sqlstr & " and Title like '%"&request("Title")&"%'"
		end if
	end if
	if sqlstr<>"" then sqlstr=" where "&sqlstr
	sqlstr=sqlstr & " order by "&request("ordername")
	if request("order")<>"esc" then sqlstr=sqlstr&" "&request("order")
	
	sql="select * from [test] "& sqlstr
	'response.write sql
	rs.open sql,aconn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=6 class=tablebody1>û���ҵ���ؼ�¼��</td></tr>"
	else		
%>		
	<FORM METHOD=POST ACTION="?action=selaction">
		  <tr>
			<td width="40" align="center" class=TopLighNav1><b>�⡡��</b></td>
			<td width="45" align="center" class=TopLighNav1><b>�⡡��</b></td
			><td width="*" align="center" class=TopLighNav1><b>�� Ŀ(�����Ŀ���б���༭)</b></td>
			<td width="90" align="center" class=TopLighNav1><b>��Ŀ���</b></td>
			<td width="80" align="center" class=TopLighNav1><b>�����ϴ���</b></td> 
			<td width="110" align="center" class=TopLighNav1><b>�١���</b></td>   
		  </tr>
<%
			rs.PageSize = Cint(Forum_Setting(11)) 
			rs.AbsolutePage=currentpage
			page_count=0
			totalrec=rs.recordcount
			while (not rs.eof) and (not page_count = Cint(Forum_Setting(11)))
%>
			  <tr>
				<td align="center" class=tablebody1><%=rs("id")%></td>
				<td class=tablebody1 align="center"><%if rs("tx")=1 then%>�����<%else%>ѡ����<%end if%></td>
				<td class=tablebody1><a href="?action=edit&Tid=<%=rs("id")%>"><%=rs("title")%></a></td> 
				<td class=tablebody1 align="center"><%=GetLbName(rs("lb"))%></td>
				<td class=tablebody1 align="center"><%=rs("username")%></td> 
				<td class=tablebody1 align="center"><a href="?action=edit&Tid=<%=rs("id")%>" title="�༭��Ŀ">�༭</a> | <a href="?action=del&Tid=<%=rs("id")%>" title="ֱ��ɾ������" onclick="javascript:{if(confirm('��ȷ��ɾ��������?')){return true;}return false;}">ɾ��</a> | <input type="checkbox" name="TitleID" value="<%=rs("id")%>"></td>
			  </tr>
<%
			page_count = page_count + 1
			rs.movenext
			wend
			Pcount=rs.PageCount
%>
			<tr><td colspan=6 class=TopLighNav1 align=center>��ҳ��
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1&userSearch="&request("userSearch")&"&username="&request("username")&"&TitleClass="&request("TitleClass")&"&TitleLb="&request("TitleLb")&"&Title="&request("Title")&"&usernamechk="&request("usernamechk")&"&ordername="&request("ordername")&"&order="&request("order")&"&action="&request("action")&""">[1]</a> ..."
			end if
			if Pcount>currentpage+3 then
			endpage=currentpage+3
			else
			endpage=Pcount
			end if
			for i=currentpage-3 to endpage
			if not i<1 then
				if i = clng(currentpage) then
				response.write " <font color=red>["&i&"]</font>"
				else
				response.write " <a href=""?page="&i&"&userSearch="&request("userSearch")&"&username="&request("username")&"&TitleClass="&request("TitleClass")&"&TitleLb="&request("TitleLb")&"&Title="&request("Title")&"&usernamechk="&request("usernamechk")&"&ordername="&request("ordername")&"&order="&request("order")&"&action="&request("action")&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&"&userSearch="&request("userSearch")&"&username="&request("username")&"&TitleClass="&request("TitleClass")&"&TitleLb="&request("TitleLb")&"&Title="&request("Title")&"&usernamechk="&request("usernamechk")&"&ordername="&request("ordername")&"&order="&request("order")&"&action="&request("action")&""">["&Pcount&"]</a>"
			end if
%>		  
<tr><td colspan=5 class=tablebody1 align=center><B>��ѡ������Ҫ���еĲ���</B>��<input type="radio" name="selaction" value=1 checked>����ɾ��&nbsp;&nbsp;<input type="radio" name="selaction" value=2>�ƶ������&nbsp;&nbsp;
<select size=1 name="lbid">
<%
		set ars=aconn.execute("select lb,id from [testlb]")
		do while not ars.eof
			response.write "<option value="&ars(1)&">"&ars(0)&"</option>"
			ars.movenext
		loop
		ars.close
%>
	</select>
	</td>
	<td class=tablebody1 align=center><input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">
	</td>
	</tr>
	<tr><td colspan=6 class=tablebody1 align=center>
	<input type=submit name=submit value="ִ��ѡ���Ĳ���"  onclick="{if(confirm('ȷ��ִ��ѡ��Ĳ�����?')){this.document.recycle.submit();return true;}return false;}">
	</td></tr>
	</FORM>		  
<%
	end if 
end if			
%>  					
</table>	
<%
end sub

sub edit()
	dim id,ausername
	dim alb,akey1,akey2,akey3,akey4,ajj,atime,atitle,aok,atx

	id=ltrim(trim(request("Tid")))
	if id="" or (not isnumeric(id)) then
		errmsg=errmsg+"<br><li>��ָ����ȷ�Ĳ���"
		call test_err()
		exit sub
	end if 

	set ars=server.createobject("adodb.recordset")
	sql="select lb,title,key1,key2,key3,key4,ok,username,id,tx,jj,time from [test] where id="&id
	ars.open sql,aconn,1,1
	if ars.eof and  ars.bof then
		ars.close
		errmsg=errmsg+"<br><li>�������Ŀ������û���ҵ�����Ŀ"
		call test_err()
		exit sub
	end if	
	alb=ars(0)
	atitle=ars(1)
	akey1=ars(2)
	akey2=ars(3)
	akey3=ars(4)
	akey4=ars(5)
	aok=ars(6)
	ausername=ars(7)
	id=ars(8)
	atx=ars(9)
	ajj=ars(10)
	atime=ars(11)
	ars.close
%>
<form method="POST" action='Z_TestEdit.asp?action=saveedit&Tid=<%=id%>'>
<table align=center cellspacing=1 cellpadding="3" class=tableborder1 style="width:97%">
  <tr><td colspan=8 class=TopLighNav1 height=23>�༭��Ŀ&nbsp; &nbsp; &nbsp; &nbsp;<a href="Z_TestEdit.asp?action=edit&Tid=<%=id-1%>"><font color=<%=Forum_body(8)%>>��һ��</font></a> | <a href="Z_TestEdit.asp?action=edit&Tid=<%=id+1%>"><font color=<%=Forum_body(8)%>>��һ��</font></a> | <a href="?">�߼���ѯ</a> | <a href="<%=Request.ServerVariables("HTTP_REFERER")%>">�� ��</a></td></tr> 
  <tr>
    <td width="70" align="center" class=tablebody2><b>�⡡��</b></td>
    <td width="82" class=tablebody1>&nbsp;<%=id%> </td>
    <td width="98" class=tablebody2><b>&nbsp;��Ŀ�������</b></td>
    <td width="102" class=tablebody1><select name="lb"><%
	set ars=server.createobject("adodb.recordset")
	sql="select lb,id from [testlb]"
	ars.open sql,aconn,1,1
	do while not ars.eof
%>
		<option value='<%=ars(1)%>' <%if ars(1)=alb then%>selected<%end if%>><%=ars(0)%></option>
<%
  		ars.movenext
	loop
	ars.close
%></select></td>
    <td width="57" class=tablebody2><b>����</b></td>
    <td width="63" class=tablebody1><%if atx=1 then%>�����<%else%>ѡ����<%end if%></td>
    <td width="90" class=tablebody2><b>�����ϴ���</b></td>
    <td width="102" class=tablebody1><input type="text" name="username" size="8" value=<%=ausername%>></td>
  </tr>
  <tr >
    <td width="70"  align="center" class=tablebody2><b>��&nbsp;&nbsp;Ŀ</b></td>
    <td width="600" colspan="7" align="left" class=tablebody1>&nbsp;<textarea rows="12" name="title" cols="70"><%=atitle%></textarea></td>
  </tr>
  <tr >
    <td width="70"  align="center" class=tablebody2><%if atx=1 then%>��ȷ��<%else%>��1<%end if%></td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key1" size="70" value='<%=akey1%>'></td>
  </tr>
  <%if atx=0 then%>
  <tr>
    <td width="70" align="center" class=tablebody2>��2</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key2" size="70"  value='<%=akey2%>'></td>
  </tr>
  <tr >
    <td width="70" align="center" class=tablebody2>��3</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key3" size="70" value='<%=akey3%>'></td>
  </tr>
  <tr>
    <td width="70"  align="center" class=tablebody2>��4</td>
    <td width="600" colspan="7" class=tablebody1>&nbsp;<input type="text" name="key4" size="70"  value='<%=akey4%>'></td>
  </tr><%end if%>
  <tr>
    <td width="70"  align="center" class=tablebody2><%if atx=0 then%>��ȷ��<%end if%></td>
    <td width="81" class=tablebody2>&nbsp;<%if atx=0 then%><input type="text" name="ok" size="10"  value='<%=aok%>'><%end if%></td>
    <td width="116" colspan="2" class=tablebody2 align=center>��ʱ <input type="text" name="atime" size="7" value=<%=atime%>> ��</td>
    <td width="172" colspan="2" class=tablebody2 align=center>���⽱��<input type="text" name="jj" size="7" value=<%=ajj%>>Ԫ</td>
    <td width="231" colspan="2" class=tablebody2></td>
  </tr>
</table>
<tr><td align=center class=tablebody2><input type="submit" value="�޸ı���" name="reaction">����<input type="submit" value="ɾ������" name="reaction"></td></tr>
</form>
<%
end sub	

sub saveedit()
	dim id
	id=request("Tid")
	if id="" or (not isnumeric(id)) then
		errmsg=errmsg+"<br><li>��ָ����ȷ�����"
		call test_err()
		exit sub	
	end if
	if request("reaction")="�޸ı���" then
		dim ausername
		dim lb,key(4),jj,atime,title,ok,tx		
		lb=request("lb"):		jj=request("jj"):		atime=request("atime"):		title=request("title")
		key(0)=request("key1"):		key(1) =request("key2"):		key(2) =request("key3"):		key(3) =request("key4")
		ok=request("ok"):		tx=request("tx"):		ausername=trim(replace(request("username"),"'",""))
			
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
		if cint(tx)=0 then
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
				
		set ars=server.createobject("adodb.recordset")
		sql="select * from [test] where id="&id
		ars.open sql,aconn,1,3
		if ars.eof and ars.bof then
			errmsg=errmsg+"<br><li>�������Ż��߸���Ŀ�Ѿ�ɾ��"
			call test_err()
			exit sub		
		end if
		ars("lb")=lb
		ars("jj")=jj
		ars("time")=atime
		ars("title")=title
		ars("key1")=key(0)
		ars("username")=ausername		
		if ars("tx")=0 then			'����ѡ����
			ars("key2")=key(1) 
			ars("key3")=key(2) 
			ars("key4")=key(3) 
			ars("ok")=ok
		end if
		ars.update
		ars.close
		set ars=nothing
		sucmsg=sucmsg+"<br><li>�༭��Ŀ�ɹ����뷵�ؽ�����������"
		'rUrl="Z_TestAdminUpLoad.asp"
		call suc()
	elseif request("reaction")="ɾ������" then
	    rUrl="Z_Testedit.asp"
		call del()
	end if
end sub

sub del()
	dim id
	id=request("Tid")
	if id="" or (not isnumeric(id)) then
		errmsg=errmsg+"<br><li>��ָ����ȷ�����"
		call test_err()
		exit sub	
	end if
	aconn.execute("delete from [test] where id="&id)
	sucmsg=sucmsg+"<br><li>ɾ����Ŀ�ɹ����뷵�ؽ�����������"
	call suc()
end sub

sub selaction()
	if request("selaction")="" then
		errmsg=errmsg+"<br><li>��ָ����ز���"
		call test_err()
		exit sub
	end if
	if request("selaction")=1 then		'����ɾ��
		aconn.execute("delete from [test] where id in ("&replace(request("TitleID"),"'","")&")")
		sucmsg=sucmsg+"<br><li>�����ɹ����뷵�ؽ�����������"
		call suc()	
	elseif request("selaction")=2 then '�����Ƶ������
		dim lbid
		lbid=request("lbid")
		aconn.execute("update [test] set lb="&lbid&" where id in("&replace(request("TitleID"),"'","")&")")
		sucmsg=sucmsg+"<br><li>�����ɹ����뷵�ؽ�����������"
		call suc()	
	else
		errmsg=errmsg+"<br><li>����Ĳ���"
		call test_err()
		exit sub	
	end if
end sub
'================�����ظ�����Ŀ====================
sub ClearTitle()
	Server.ScriptTimeOut=9999999
	set ars=server.createobject("adodb.recordset")
	dim startime1
	startime1=timer()
	if cint(request("DeleteSel"))=1 then
		sql="select id,title from [test] order by id"
		ars.open sql,aconn,1,1
		on error resume next
		for i=1 to ars.recordcount
			aconn.execute("delete from [test] where title='"&ars(1)&"' and id>"&ars(0))
			ars.movenext
		next
		ars.close		
	else
		sql="select id,title from [test] order by id desc"
		ars.open sql,aconn,1,1
		on error resume next
		for i=1 to ars.recordcount
			aconn.execute("delete from [test] where title='"&ars(1)&"' and id<"&ars(0))
			ars.movenext
		next
		ars.close		
	end if	
	sucmsg=sucmsg+"<br><li>��Ŀ�������! ���β���ִ��ʱ�䣺"&FormatNumber((timer()-startime1)*1000,3)&" ����"
	call suc()	
end sub
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>