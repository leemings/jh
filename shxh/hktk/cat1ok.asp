<%@ LANGUAGE = VBScript %>
<!--#include file="jha.asp"-->
<%
nl=request.form("nl")
if nl=0 then
	mess="û��������ô����?"
else
	Jname=session("Ba_jxqy_username")
	sql="select * from �û� where ����='" & Jname & "' and ����>" & nl & " and ����>300 and ״̬<>'��' and ״̬<>'��'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
		mess=Jname & "���㲻�ܹ���,��������������������300"
else
	
        if rs("����")>nl then
		mess=Jname & "����û����ô��������"
	else
			randomize timer
			r=int(rnd*nl)
			m=int(rnd*100)
			rou=int(rnd*2)
			if r>m then
				mess=Jname&"�����è,�õ�����" & rou & "��"
				sql="update �û� set ����=����-" & nl & ",����=����+" & rou & " where ����='" & Jname & "'"
				conn.execute sql
			elseif r<m then
				mess=Jname&"��èצץ��"
				sql="update �û� set ����=����-" & nl & ",����=����-" & m & " where ����='" & Jname & "'"
				conn.execute sql
			end if
        end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end if
%>
<TITLE>����ѵ����</TITLE>
<style>td{font-size:14}</style>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=xuetang.asp>����</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>