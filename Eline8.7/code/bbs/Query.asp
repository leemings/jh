<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
stats="��̳����"
if boardmaster or master then
	guestlist=""
else
	guestlist=" lockboard<>1 and "
end if
call nav()
if BoardID="" or not isInteger(BoardID) then
	boardid=0
	call head_var(2,0,"","")
else
	boardid=cLng(Boardid)
	call head_var(1,BoardDepth,0,0)
end if
call main()
if founderr then call dvbbs_error()
call footer()

sub main()
	if Cint(GroupSetting(14))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳������Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
		founderr=true
		exit sub
	end if%>
<form action=queryresult.asp method=post>
<table cellpadding=5 cellspacing=1 align=center class=tableborder1>
	<tr>
		<th valign=middle colspan=2 >������Ҫ�����Ĺؼ���</th>
	</tr>
	<tr>
		<td class=tablebody1 colspan=2 align=center valign=middle>(ͬʱ��ѯ�������ʹ��'<font  color=<%=Forum_body(8)%>>or</font>' �ָ��ؼ��֣���ѯͬʱ����ĳ����ʹ��'<font  color=<%=Forum_body(8)%>>and</font>'�ָ��ؼ���)<br><br><input type=text size=40 name=keyword></td>
	</tr>
	<tr>
		<td class=tablebody2 valign=middle colspan=2 align=center><b>����ѡ��</b></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>��������</b>&nbsp;<input name=sType type=radio value=1></td>
		<td class=tablebody1 align=left valign=middle><select name=nSearch><option value=1>������������<option value=2>�����ظ�����<option value=3>���߶�����</select></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>�ؼ�������</b>&nbsp;<input name=sType type=radio value=2 checked></td>
		<td class=tablebody1 align=left valign=middle><select name=pSearch><option value=1>�������������ؼ���<!--<option value=2>�����������������ؼ���<option value=3>���߶�����--></select></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>���ڷ�Χ</b></td>
		<td class=tablebody1 align=left valign=middle><select name=SearchDate class=smallsel><option value=ALL>��������<option value=1>��������<option  selected value=5>5������<option value=10>10������<option value=30>30������</option></select><select name=Stable size=1><%
			For i=0 to ubound(AllPostTable)
				response.write "<option value="""&AllPostTable(i)&""""
				if trim(NowUseBBS)=trim(AllPostTable(i)) then response.write " selected "
				response.write ">"&AllPostTablename(i)&"</option>"
			next
		%></select></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>����������</b></td>
		<td class=tablebody1 align=left valign=middle><input type=text size=20 name=scoreuser></td>
	</tr>
	<tr>
		<td class=tablebody1 align=right valign=middle><b>����ֵ����</b></td>
		<td class=tablebody1 align=left valign=middle><select name=SearchScore class=smallsel><option value=ALL selected>���������<option value=NONE>��δ������<option value=A5>������ϵļӷ���<option value=A4>�ķ����ϵļӷ���<option value=A3>�������ϵļӷ���<option value=A2>�������ϵļӷ���<option value=A1>һ�����ϵļӷ���<option value=Z0>�����<option value=B1>һ�����ϵļ�����<option value=B2>�������ϵļ�����<option value=B3>�������ϵļ�����<option value=B4>�ķ����ϵļ�����<option value=B5>������ϵļ�����</select></td>
	</tr>
	<tr>
		<td class=tablebody1 valign=middle colspan=2 align=center><b>��ѡ��Ҫ��������̳</b></td>
	</tr>
	<tr>
		<td class=tablebody1 colspan=2 valign=middle align=center><b>������̳��</b>&nbsp;<select name=boardid size=1><option value=0>������̳</option><%
			set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
			do while not rs.EOF
				response.write "<option value="""&rs(0)&""" "
				if boardid=rs(0) then
					response.write "selected"
				end if
				response.write ">"
				select case rs(2)
				case 0
					response.write "��"
				case 1
					response.write "&nbsp;&nbsp;��"
				end select
				if rs(2)>1 then
					for i=2 to rs(2)
						response.write "&nbsp;&nbsp;��"
					next
					response.write "&nbsp;&nbsp;��"
				end if
				response.write rs(1)&"</option>"
				rs.MoveNext
			loop
			set rs=nothing
		%></select></td>
	</tr>
	<tr>
		<td class=tablebody2 valign=middle colspan=2 align=center><input type=submit value=��ʼ����></td>
	</tr>
</table>
</form>
<%call activeonline()
end sub%>