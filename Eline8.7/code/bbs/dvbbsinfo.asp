<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<%boardid=request("boardid")
if not isnumeric(boardid) or boardid="" or boardid=0 then
	boardid=0
end if
Stats="��̳������Ϣ"
call activeonline()
call nav()
call head_var(1,BoardDepth,0,0)%>
<BR>
<table cellpadding=6 cellspacing=1 align=center class=tableborder2>
	<tr><td class=tablebody1>��̳������Ϣ</td></tr>
	<tr><td class=tablebody1><a href=?view=Setting&boardid=<%=boardid%>>��������</a> | <a href=?view=info&boardid=<%=boardid%> >������Ϣ</a> | <a href=?view=Group&boardid=<%=boardid%>>�û���Ȩ��</a> | <a href=?view=css&boardid=<%=boardid%>>��ɫCSS</a> | <a href=?view=pic&boardid=<%=boardid%>>��̳ͼƬ</a> | <a href=?view=boardpic&boardid=<%=boardid%>>��̳����ͼ��</a> | <a href=?view=TopicPic&boardid=<%=boardid%>>��̳����ͼ��</a> | <a href=?view=userset&boardid=<%=boardid%>>�û��趨</a> | <a href=?view=boardset&boardid=<%=boardid%>>�ְ��趨</a></td></tr>
</table>
<%if  request("view")="Setting" then
	call Setting()
elseif request("view")="info" then
	call info()
elseif request("view")="Group" then
	call Group()
elseif request("view")="css" then
	call css()
elseif request("view")="pic" then
	call pic()
elseif request("view")="boardpic" then
	call boardpic()
elseif request("view")="TopicPic" then
	call TopicPic()
elseif request("view")="userset" then
	call userset()
elseif request("view")="boardset" then
	if boardid>0 then
		call boardset()
	else
		call Setting()
	end if
else
	call Setting()
end if
call footer()

sub Setting()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_Setting()  ����������Ϣ</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_Setting()  ����������Ϣ  ��ǰֵ��������</td></tr>
    <%i=-1%>
    <tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(0) ������ʱ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(1) �ű���ʱʱ��(��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(2) �����ʼ����  0����֧�� 1��JMAIL 2��CDONTS  3��ASPEMAIL</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(3) �Ƿ����ü�¼�����Ķ��û�����  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(4) ��ҳģʽ   0-�°���ʽ 1-�ɰ���ʽ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(5) ��ҳ��ʾ��̳���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(6) �Ƿ������û��Զ���ͷ��   0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(7) �Ƿ�����ͷ���ϴ�   0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(8) ɾ������û�ʱ��(����)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(9) ��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(10) �¶���Ϣ��������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(11) �����б��������ӳ����ÿҳ��ʾ����¼</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(12) ��ҳ�Ƿ���ʾ����б�(ͣ��)   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(13) ��ҳ�Ƿ���ʾ����վ   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(14) ����������ʾ�û�����  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(15) ����������ʾ��������  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(16) ��ҳ�Ƿ���ʾ����   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(17) ��ҳ���ǵ�һ�е���ʾ��ʽ<BR>1-���� 2-���� 3-���� 4-���� 5-������ 6-������� 7-���Ů��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(18) ��ҳ���ǵڶ��е���ʾ��ʽ<BR>1-���� 2-���� 3-���� 4-���� 5-������ 6-������� 7-���Ů��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(19) ��ˢ�»���   0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(20) ���ˢ��ʱ����(��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(21) ��̳��ǰ״̬</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(22) ͬһIPע����ʱ��(��)   0-����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(23) Email֪ͨ����   0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(24) һ��Emailֻ��ע��һ���ʺ� 0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(25) ע����Ҫ����Ա��֤  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(26) ����ͬʱ������(��)   0-����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(27) �����й��ģʽ   0����̬��� 1����������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(28) �����б��Ƿ���ʾ�û�IP</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(29) �Ƿ���ʾ�����ջ�Ա  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(30) �Ƿ���ʾҳ��ִ��ʱ��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(31) ���ٵ�¼λ��   1������ 2���ײ� 0������ʾ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(32) �Ƿ�����̳����  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(33) ���������б���ʾ��ǰλ��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(34) ���������б���ʾ��¼�ͻʱ��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(35) ���������б���ʾ������Ͳ���ϵͳ  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(36) ���������б���ʾ��Դ  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(37) �Ƿ��������û�ע��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(38) Ĭ��ͷ����(����)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(39) Ĭ��ͷ��߶�(����)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(40) ����û�������(�ַ�)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(41) ��û�������(�ַ�)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(42) �������ǩ��   0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(43) �Ƿ���ʾ��������Ա  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(44) ��Ϊ���Ż�����������ֵ(�ظ���)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(45) �������Ƿ���ʾ���͹�Ʊ  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(46) �������Ż�ӭ��ע���û�  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(47) ����ע����Ϣ�ʼ�  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(48) �༭����������ʾ"��xxx��yyy�༭"����Ϣ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(49) ����Ա�༭����ʾ"��XXX�༭"����Ϣ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(50) �ȴ�"��XXX�༭"��Ϣ��ʾ��ʱ��(����)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(51) �༭����ʱ��(����)   0-����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(52) ��ֹ���ʼ���ַ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(53) �����û�ʹ��ͷ��   0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(54) ʹ���Զ���ͷ������ٷ�����(��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(55) ���������վ���ϴ�ͷ��  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(56) ��������ͷ���ļ���С(KB)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(57) ���ͷ����(����)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(58) ���ͷ��߶�(����)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(59) �û�ͷ����󳤶�(�ַ�)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(60) �Զ���ͷ�����ٷ�����������(��)   0-����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(61) �Զ���ͷ��ע����������(��)   0-����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(62) �Զ���ͷ������������������һ������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(63) �Զ���ͷ����Ҫ���εĴ���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(64) ��ˢ�¹�����Ч��ҳ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(65) �û�ǩ���Ƿ���UBB����  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(66) �û�ǩ���Ƿ���HTML���� 0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(67) �û��Ƿ�����ͼ��ǩ  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(68) �û������б����(��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(69) �Ƿ�ʱ������̳  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(70) ��̳����ʱ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(71) �����˵�ģʽ  0-�����̶� 1-��໬��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(72) �����л�Ա������ʾ��ʽ  0-������ϸ 1-����� 2-ͷ����ϸ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(73) �Ƿ�����������  0-���� 1-���������� 2-���ƻظ��� 3-����������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(74) �������Ʊ���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(75) ��ҳ�Ƿ���ʾ������ϲ������̳��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(76) ���ͼƬ�Ƿ��뵭��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Forum_Setting(77) ��ҳLogo������̳��ʽ  0-������������ 1-ֻ���������� 2-ֻ���������� 3-�ȹ����ֵ���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
		</tr>
		<%for i=78 to ubound(Forum_Setting)%>
			<tr>
				<td width="70%" class=tablebody2><li>Forum_Setting(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Setting(i)%></b></font></td>
			</tr>
		<%next%>
    <tr><th height="24" colspan="2"></th></tr>
  </table>
<%end sub 

sub info()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_Info()  ������Ϣ</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_Info()   ������Ϣ  ��ǰֵ��������</td></tr>
    <%i=-1%>
    <tr>
			<td width="40%" class=tablebody2><li>Forum_Info(0) ��̳����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(1) ��̳��url</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(2) ��ҳ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(3) ��ҳURL</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(4) SMTP Server��ַ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(5) ��̳����ԱEmail</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(6) ��̳��ҳLogo��ַ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(7) ��̳ͼƬĿ¼</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(8) ��̳����Ŀ¼</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(9) ��̳����ʱ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(10) ��������Ŀ¼</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(11) ��̳ͷ��Ŀ¼</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<tr>
			<td width="40%" class=tablebody2><li>Forum_Info(12) ��̳�Ľ�վ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
		</tr>
		<%for i=13 to ubound(Forum_Info)%>
			<tr>
				<td width="40%" class=tablebody2><li>Forum_Info(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_Info(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Info(i)%></b></font></td>
			</tr>
		<%next%>
		<tr><th height="24" colspan="2"></th></tr>
	</table>
<%end sub 

sub Group()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">GroupSetting()[reGroupSetting()]  �û���Ȩ������</th></tr>
    <tr><td class=tablebody1 colspan="2">GroupSetting()[reGroupSetting()]  �û���Ȩ������  ��ǰֵ��������</td></tr>
    <%i=-1%>
    <tr>
			<td width="70%" class=tablebody2><li>GroupSetting(0)  ���������̳   0���� 1����</td>
			<td width="*" class=tablebody2>GroupSetting(0) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(0)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(1)  ���Բ鿴��Ա��Ϣ  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(2)  ���Բ鿴�����˷��������� 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(3)  ���Է���������   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(4)  ���Իظ��Լ�������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(5)  ���Իظ������˵�����  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(6)  ��������̳�������ֵ�ʱ���������0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(7)  �����ϴ�����   0���� 1���� 2-�������ϴ� 3-�������ϴ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(8)  ���Է�����ͶƱ   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(9)  ���Բ���ͶƱ   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(10) ���Ա༭�Լ�������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(11) ����ɾ���Լ�������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(12) �����ƶ��Լ������ӵ�������̳ 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(13) �Դ�/�ر��Լ����������� 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(14) ����������̳   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(15) ����ʹ�á����ͱ�ҳ�����ѡ����� 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(16) �����޸ĸ�������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(17) ���Է���С�ֱ�   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(18) ����ɾ������������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(19) �����ƶ�����������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(20) ���Դ�/�ر�����������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(21) ���Թ̶�/����̶�����  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(22) ���Խ���/�ͷ������û�  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(23) ���Ա༭����������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(24) ���Լ���/�����������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(25) ���Է�������   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(26) ���Թ�����   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(27) ���Թ���С�ֱ�   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(28) ��������/����/��������û� 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(29) ����ɾ���û�1��10������������ 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(30) ���Բ鿴����IP����Դ  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(31) �����޶�IP����   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(32) ���Է��Ͷ���   0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(33) ��෢���û�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(34) �������ݴ�С����(�ֽ�)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(35) �����С����(KB)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(36) �Ƿ���������ӵ�Ȩ�� 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(37) �Ƿ��н���������̳��Ȩ�� 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(38) ���Խ��������̶ܹ����� 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(39) ���������̳�¼�  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(40) һ������ϴ��ļ�����(��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(41) ���������������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(42) ���Թ����û�Ȩ��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(43) ���Խ���/�ͷ��û�  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(44) �ϴ��ļ���С����(KB)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(45) ��������ɾ�����ӣ�ǰ̨�� 0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(46) ����С�ֱ������Ǯ(Ԫ)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(47) �������������Ǯ(Ԫ)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(48) ����̳�ļ�����Ȩ��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(49) ���������̳չ��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(50) һ������ϴ��ļ�����(��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>GroupSetting(51) ������������ɫȨ��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
		</tr>
		<%for i=52 to ubound(GroupSetting)%>
			<tr>
				<td width="70%" class=tablebody2><li>GroupSetting(<%=i%>) </td>
				<td width="*" class=tablebody2>GroupSetting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=GroupSetting(i)%></b></font></td>
			</tr>
		<%next%>
		<tr><th height="24" colspan="2"></th></tr>
	</table>
<%end sub 

sub css()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_Body()  ������ɫCSS�ı���</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_Body()  ������ɫCSS�ı���  ��ǰֵ��������</td></tr>
		<%i=-1%>
    <tr>
			<td width="60%" class=tablebody2><li>Forum_Body(0) ���߿��壩CSS����һ ���ã�Class=TableBorder1</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(1) ���߿��壩CSS����� ���ã�Class=TableBorder2</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(2) ������CSS����һ������� ���ã�th</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(3) ������CSS�������ǳ������ ���ã�Class=TableTitle2</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(4) �����CSS����һ ���ã�Class=TableBody1</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(5) �������ɫ��(1��2��ɫ����ʾ�д���) ���ã�Class=TableBody2</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(6) ��̳�����ܵ�CSS����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(7) ��������������CSS���壨����� ���ã�id=TableTitleLink</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(8) ��������������ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(9) ��ʾ���ӵ�ʱ��������ӣ�ת�����ӣ��ظ��ȵ���ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(10) ��̳����ܵ�CSS����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(11) ��̳BODY��ǩ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(12) �����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(13)  </td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(14) ��ҳ������ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(15) һ���û�����������ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(16) һ���û������ϵĹ�����ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(17) ��������������ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(18) ���������ϵĹ�����ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(19) ����Ա����������ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(20) ����Ա�����ϵĹ�����ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(21) �������������ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(22) ��������ϵĹ�����ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(23) �����˵����CSS����(Logo & Banner�·�) ���ã�Class=TopLighNav</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(24) �����˵����CSS����(Logo & Banner�Ϸ�) ���ã�Class=TopDarkNav</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(25) Body��CSS����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(26) �����˵����CSS����(�����˵�) ���ã�Class=TopLighNav1</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(27) ���߿���ɫ���� �����涨��һ�Ͷ�CSS������Ʋ����ĵط�����ñ��ֺͱ߿�CSS����һ��ͬ������ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_Body(28) �����˵����CSS����(Logo & banner����) ���ã�Class=TopLighNav2</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
		</tr>
		<%for i=29 to ubound(Forum_Body)%>
			<tr>
				<td width="60%" class=tablebody2><li>Forum_Body(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_Body(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Body(i)%></b></font></td>
			</tr>
		<%next%>
		<tr><th height="24" colspan="2"></th></tr>
	</table>
<%end sub 

sub pic()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr><th height="24" colspan="2">Forum_Pic()  ��̳ͼƬ����</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_Pic()  ��̳ͼƬ����  ��ǰֵ��������</td></tr>
    <%i=-1%>
    <tr>
			<td width="30%" class=tablebody2><li>Forum_Pic(0) ��̳����Ա</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(1) ��̳����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(2) ��̳���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(3) ��ͨ��Ա</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(4) ���˻������Ա</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(5) ͻ����ʾ�Լ�����ɫ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(6) ������̳������������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(7) ������̳������������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(8) �����б����˵��ǰ���ͼ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(9) �����б��н��շ�����ͼ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(10) �����б����ܷ�����ͼ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(11) �����б�����������ͼ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(12) ��̳����ͼ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(13) �����б��Ϸ������б�ͼ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(14) ��֤��̳������������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(15) ��֤��̳������������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_Pic(16) ��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
		</tr>
		<%for i=17 to ubound(Forum_Pic)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_Pic(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_Pic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_Pic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_Pic(i)%> border=0></td>
			</tr>
		<%next%>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>
<%end sub 

sub boardpic()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr><th height="24" colspan="2">Forum_BoardPic()  ��̳����ͼ��</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_BoardPic()  ��̳����ͼ��  ��ǰֵ��������</td></tr>
    <%i=-1%>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(0) ������̳</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(1) ��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(2) ����ͶƱ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(3) С�ֱ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(4) �ظ�����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(5) ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(6) ������ҳ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(7) �����ϼ�Ŀ¼</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(8) ��ǰĿ¼</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(9) �µĶ���Ϣ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(10) �ҷ��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(11) �����������ģʽ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(12) ƽ�����������ģʽ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(13) ��һƪ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(14) ��һƪ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(15) ˢ�����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_BoardPic(16) ��̳����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
		</tr>
		<%for i=17 to ubound(Forum_BoardPic)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_BoardPic(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_BoardPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_BoardPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_BoardPic(i)%> border=0></td>
			</tr>
		<%next
		i=-1%>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(0) �򿪵�����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(1) ���ŵ�����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(2) ����������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(3) �̶�������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(4) ����������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(5) �����б��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(6) �����б��з�ҳͼ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(7) ˢ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(8) ������ʾ���ļ���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(9) �̶ܹ�������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(10) </td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(11) </td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_StatePic(12) ͶƱ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
		</tr>
		<%for i=13 to ubound(Forum_StatePic)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_StatePic(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_StatePic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_StatePic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_StatePic(i)%> border=0></td>
			</tr>
		<%next
		i=-1%>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(0) ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(1) б��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(2) �»���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(3) ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(4) URL����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(5) Email��ַ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(6) ��ͼ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(7) ��Flash</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(8) ��ShockWave</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(9) ��RMӰƬ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(10) ��Media PlayerӰƬ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(11) ��QuickTimeӰƬ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(12) ��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(13) ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(14) �ƶ���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(15) ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(16) ��Ӱ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(17) �����б�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<tr>
			<td width="30%" class=tablebody2><li>Forum_UBB(18) ��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
		</tr>
		<%for i=19 to ubound(Forum_UBB)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_UBB(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_UBB(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_UBB(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_UBB(i)%> border=0></td>
			</tr>
		<%next%>
    <tr><th height="24" colspan="2"></th></tr>
  </table>
<%end sub

sub topicpic()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_TopicPic()  ��̳����ͼ��</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_TopicPic()  ��̳����ͼ��  ��ǰֵ��������</td></tr>
    <%i=-1%>
    <tr>
			<td width="30%" class=tablebody2><li>Forum_TopicPic(0) ���浱ҳΪ�ļ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(1) ���汾��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(2) ��ʾ�ɴ�ӡ�İ汾</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(3) �ѱ�������ʵ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(4) �ѱ���������̳�ղ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(5) ���ͱ�ҳ�������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(6) �ѱ�������IE�ղ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(7) ���Ͷ��Ÿ�����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(8) �����û���������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(9) ����û���Ϣ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(10) �û�����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(11) �û�OICQ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(12) �û�ICQ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(13) �û�MSN</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(14) �û���ҳ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(15) ��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(16) �༭����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(17) ɾ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(18) ��������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(19) ���뾫������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(20) �û�IP</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(21) ��Ϊ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<tr>
      <td width="30%" class=tablebody2><li>Forum_TopicPic(22) ���ٻظ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
		</tr>
		<%for i=23 to ubound(Forum_TopicPic)%>
			<tr>
				<td width="30%" class=tablebody2><li>Forum_TopicPic(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_TopicPic(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_TopicPic(i)%></b></font>&nbsp;&nbsp;<img src=pic/<%=Forum_TopicPic(i)%> border=0></td>
			</tr>
		<%next%>
		<tr><th height="24" colspan="2"></th></tr>
	</table>
<%end sub

sub userset()
	dim i
	%><table cellpadding=4 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Forum_User() �û��趨</th></tr>
    <tr><td class=tablebody1 colspan="2">Forum_User() �û��趨  ��ǰֵ��������</td></tr>
    <%i=-1%>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(0) ע���Ǯ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(1) �������ӽ�Ǯ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(2) �������ӽ�Ǯ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(3) ɾ�����ٽ�Ǯ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(4) ��¼���ӽ�Ǯ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(5) ע�ᾭ��ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(6) �������Ӿ���ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(7) �������Ӿ���ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(8) ɾ�����پ���ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(9) ��¼���Ӿ���ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(10) ע������ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(11) ������������ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(12) ������������ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(13) ɾ����������ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(14) ��¼��������ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(15) �������ӽ�Ǯ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(16) ������������ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(17) �������Ӿ���ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(18) ��ͨ��Ա�����л�Ա���ļ۸�(ͣ��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(19) ��ͨ��Ա����һ��Ա���ļ۸�(ͣ��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(20) VIP��Ա�����л�Ա���ļ۸�(ͣ��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(21) VIP��Ա����һ��Ա���ļ۸�(ͣ��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(22) ��ͨ��Ա�޸�һ��ͷ��ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(23) VIP��Ա�޸�һ��ͷ��ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(24) ��ͨ��Ա�޸�һ��ǩ���ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(25) VIP��Ա�޸�һ��ǩ���ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(26) ��ͨ��Ա���Ͷ��ŵļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(27) VIP��Ա���Ͷ��ŵļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(28) ��ͨ��Ա�ظ������֪ͨ�ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(29) VIP��Ա�ظ������֪ͨ�ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(30) ��ͨ��Ա�ظ����ʼ�֪ͨ�ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(31) VIP��Ա�ظ����ʼ�֪ͨ�ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(32) ��ͨ��Ա��������ע��ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(33) VIP��Ա��������ע��ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(34) ��ͨ��Ա�ϴ�ͷ��ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(35) VIP��Ա�ϴ�ͷ��ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(36) ��ͨ��Ա�ϴ��ļ��ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(37) VIP��Ա�ϴ��ļ��ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(38) ��ͨ��Ա�ϴ���Ƭ�ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(39) VIP��Ա�ϴ���Ƭ�ļ۸�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(40) ͶƱ���ӽ�Ǯ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(41) ͶƱ���Ӿ���ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<tr>
			<td width="60%" class=tablebody2><li>Forum_User(42) ͶƱ��������ֵ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
		</tr>
		<%for i=43 to ubound(Forum_User)%>
			<tr>
				<td width="60%" class=tablebody2><li>Forum_User(<%=i%>) </td>
				<td width="*" class=tablebody2>Forum_User(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Forum_User(i)%></b></font></td>
			</tr>
		<%next%>
    <tr><th height="24" colspan="2"></th></tr>
  </table>
<%end sub

sub boardset()
	dim i
	%><table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
		<tr><th height="24" colspan="2">Board_Setting() �ְ���̳�趨</th></tr>
		<tr><td class=tablebody1 colspan="2">Board_Setting() �ְ���̳�趨  ��ǰֵ��������</td></tr>
		<%i=-1%>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(0) ��̳����/����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(1) ��̳����/���������û������趨�Ƿ�ɽ���������̳��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(2) ��̳����/�����ܣ�������̳��Ҫ�趨�ɽ����û���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(3) ��̳��˲�����/����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(4) ����ģʽ��/�߼�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(5) HTML֧�ֿ���/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(6) UBB���ܿ���/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(7) ��ͼ��ǩ����/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(8) �����ǩ����/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(9) ��ý���ǩ����/�رգ�����Flash,RM,AVI�ȣ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(10) ��Ǯ������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(11) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(12) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(13) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(14) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(15) �ظ�������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(16) ������������ֽ���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(17) �����󷵻أ���ҳ/��̳/���ӣ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(18) ����ͬʱ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(19) �ϴ��ļ�����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(20) </td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(21) �������涨ʱ���Ź���</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(22) ��̳����ʱ�䣨0��24�㣩</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(23) �������ԣ���������/��������/Ӣ�ģ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(24) ��̳������������/Ĭ�Ϸ��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(25) ���������Ԥ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(26) �����б�ÿҳ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(27) �������ÿҳ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(28) ���������ֺ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(29) ���������м��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(30) ����ˮ����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(31) ÿ�η���ʱ����(��)</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(32) ���ͶƱ��Ŀ</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(33) ������������ɾ������</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(34) �����������޸�CSS����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(35) ���а��������޸�CSS����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(36) �Ƿ����̳�¼��еĲ����߱��ܣ��Թ���Ա��Ч��</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(37) ��̳Ĭ�϶�ȡ��������������ʱ���ڣ�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(38) Ĭ���������ظ�ʱ��/����/������/�ظ���/�����/����ʱ�䣩</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(39) �Ƿ���ü�����б�  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(40) �Ƿ�̳��ϼ�����Ȩ��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(41) ������б�һ�а�����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(42) �Ƿ���ʾ�¼���̳����  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(43) ��Ϊ������̳�Ƿ�������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(44) �Ƿ񿪷ŷ�����������  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(45) ����������Ĭ��״̬  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(46)  �Ƿ񿪷ŷ�������Żظ�����  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(47) ��������Żظ���Ĭ��״̬  0���ر� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(48) �Ƿ񿪷Ŷ�������ע�⹦��  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(49) �Ƿ񿪷Ű������ֹ���  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(50) �Ƿ�ͬѧ¼��̳  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(51) �Ƿ�VIP��̳  0���� 1����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(52) �Ƿ��Ա���̳  0����ͨ 1��Ů�� 2-����</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(53) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(54) ��Ա������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(55) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(56) �Ա�������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(57) �߼�������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(58) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(59) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(60) ��ʾ������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<tr>
			<td width="70%" class=tablebody2><li>Board_Setting(61) ����������/�ر�</td>
			<%i=i+1%>
			<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
		</tr>
		<%for i=62 to ubound(Board_Setting)%>
			<tr>
				<td width="70%" class=tablebody2><li>Board_Setting(<%=i%>) </td>
				<td width="*" class=tablebody2>Board_Setting(<%=i%>) ��ǰֵΪ��<font color=<%=Forum_Body(8)%>><b><%=Board_Setting(i)%></b></font></td>
			</tr>
		<%next%>
    <tr><th height="24" colspan="2"></th></tr>
  </table>
<%end sub%>