<%@ LANGUAGE=VBScript%>
<!--#include file="jhb.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("sjjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "������������ٽ����Ʊ��"
window.close()
</script>
<%Response.End
end if

Jname=sjjh_name
sql="delete * from ���� where ʱ��<date()"
conn.execute sql
sql="select * from ���� where �ʺ�='" & Jname & "' "
set rs=conn.execute(sql)

%>
<table border=1 bgcolor="#bee0fc" align=center width=550 cellpadding="10" cellspacing="13">
	<option><p align=center><font size=5 color=#00ff00></font></p>
	<tr><td bgcolor=#cce8d6>
		<table border=0 cellpadding=1 cellspacing=1 bgcolor="#51a8ff" width="100%" >		</table>
		
		<P align=center><front size=3pt>�����˽���Ϊ������Ľ������£��۳�1%��Ӷ��������ڹ��е�����</front></p>
		<table border=0 cellpadding=1 cellspacing=1 bgcolor="#51a8ff" width="100%" 
     >
	
		
		
        <P align=center></P>
        <TBODY>
			<tr bgcolor=#c4deff>
				<td>��Ʊ����</td><td>��������</td><td>��������</td><td>ʱ��</td><td>����</td>
			</tr>
<%DO UNTIL RS.EOF%>	
			<tr bgcolor=#f7f7f7>
				<td><%=rs("��ҵ")%>
				<td><%=rs("����")%></td>
				<td><%=rs("������")%></td>
				<td><%=rs("ʱ��")%></td>
				<td><%=formatnumber(rs("�ջ�"),0)%></td>
				
		    </tr>
<%
rs.movenext
loop
conn.close
%></p>

		</table>