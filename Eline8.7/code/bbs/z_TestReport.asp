<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="Z_TestConn.ASP"-->
<%
'=========================================================
' Script Written by ��ˮ��ɽ
'=========================================================
dim ID
stats="���Ĵǵ� �ٱ����������Ŀ"

if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ�������Ŀ��"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ�����Ŀ������"
else
	ID=request("id")
end if
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>���¼�������ز�����"
end if
if founderr then
	call nav()
	call head_var(0,0,"���Ĵǵ�","Z_test.asp")
	call test_err()
else
	call nav()
	call head_var(0,0,"���Ĵǵ�","Z_test.asp")
	if request("action")="send" then
		call send()
	else
		call main()
	end if
	call activeonline()
	if founderr then call test_err()
end if
call footer()
'=========================================
sub main()
	dim master_2
 	set rs=conn.execute("select adduser from [admin] order by id")
 	if not(rs.bof and rs.eof) then
		do while not rs.eof 
			master_2=""&master_2&"<option value="""&rs(0)&""">"&rs(0)&"</option>&nbsp;"
			rs.movenext
		loop	
	else
		master_2="�޹���Ա"
	end if
	rs.close
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <form action="?action=send&id=<%=id%>" method=post>
    <tr>
    	<th valign=middle colspan=2 height=25>�ٱ����������Ŀ������Ա</th>
	</tr>
    <tr>
    	<td class=tablebody1 width="40%"><b>���淢�͸��ĸ�����Ա��</b></td>
    	<td class=tablebody1 width="60%"><select name=master size=1><%=master_2%></select></td>
    </tr>
	<tr>
    	<td class=tablebody1 valign="top"><b>��Ϣ���ݣ�</b><br>��д�����Ŀ����������<br>�������Ϊ�𰸲���ȷ���������ȷ�Ĵ𰸣�������ԭ��<br>������ԵĻ���д�������ҵ���ȷ�𰸵��鼮�ȣ��������Ա��֤<br>����֤�����ȷʵ�д�Ļ����ٱ��߽���õ�һ���Ľ���<br>�Ǳ�Ҫ����²�Ҫʹ�������</td>
    	<td class=tablebody1><textarea name="content" cols="55" rows="6">����Ա�����ã���������ԭ�������㱨��������������Ŀ��</textarea></td>
    </tr>
	<tr>
    	<td colspan=2  class=tablebody2 align=center><input type=submit value="�� ��" name="Submit"></td>
	</tr>
	</form>
</table>
<%
end sub

sub send()
	dim body
	dim writer
	dim incept
	dim topic
	body=checkStr(request("content"))
	writer=membername
	incept=checkStr(request("master"))
	sql="select title from [test] where ID="&ID
	set rs=aconn.execute(sql)
	if not(rs.bof and rs.eof) then
		topic="�ٱ����������Ŀ"
		body=body & chr(10) &"---------------------------------------"& chr(10) &"�ٱ�����Ŀ���£�"& replace(rs(0),"'","��") & chr(10) &"�༭�����Ŀ��[URL=Z_TestEdit.asp?action=edit&Tid="&ID&"]Z_TestEdit.asp?action=edit&Tid="&ID&"[/URL]"
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>��ָ������Ŀ������</li>"
		exit sub
	end if
	rs.close
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept&"','"&membername&"','"&topic&"','"&body&"',Now(),0,1)"
	conn.execute(sql)
	sucmsg="<li>���淢�ͳɹ���<br><li>����֤�����ȷʵ�д�Ļ���������õ�һ���Ľ���"
	
	call suc()
	set rs=nothing
end sub
%>