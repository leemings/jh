<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
dim comm(30),commti(30),commsm(30)
comm(0)="/���ŷ�Ǯ$"
commti(0)="���ţ����ŷ���(�Ϲ�)�����ϸ����ŵ��ӷ�Ǯ��"
commsm(0)="���ŷ�Ǯ"
comm(1)="/ָ������$"
commti(1)="���ţ����ŷ���(�Ϲ�)�����ϸ����ŵ��Ӽ�������������"
commsm(1)="ָ������"
comm(2)="/��Ϣ$"
commti(2)="�鿴���˵���Ϣ����Ҫ5�������ӵģ�"
commsm(2)="�鿴��Ϣ"
comm(3)="/�Ĳ�$ ��Ϣ"
commti(3)="�鿴�ĳ�����Ϣ���Ĳ������"
commsm(3)="�ĳ���Ϣ"
comm(4)="/����$ ��������1-1000"
commti(4)="�����Լ��������!"
commsm(4)="E�߷��"
comm(4)="/����$ ��������1-1000"
commti(4)="�����Լ��������!"
commsm(4)="E�߷��"
<option style="color:red" value="/���$ �氮���">ʾ�����
<option style="color:red" value="/����$">���Ż���
<% if Application("sjjh_baowu")>0 then%>
<option style="color:#FF0000" value="/����$ ">��������
<%end if%>
<%if chatinfo(5)=0 then%>
<option style="color:#FF0000" value="/����$ ">��������
<%end if%>
<option style="color:green" value="/����$ ">ǧ�ﴫ��
<option style="color:green" value="/�Ķ�$ ">�Ķ��о�
<option style="color:green" value="/ŭ��$ ">��ʨŭ��
<option style="color:green" value="/����$ ">��������
<option style="color:blue" value="/��ʦ$ ">��ʦϰ��
<option style="color:blue" value="/��ͽ$ ">����ͽ��
<option style="color:blue" value="/����$ ">��������
<option style="color:blue" value="/��Ŀ$ ">��Ŀ����
<option style="color:blue" value="/����$ ">���ϰ��
<option style="color:orange" value="/����$ 100">��������
<%if chatinfo(5)=0 then%>
<option style="color:orange" value="/���Ǵ�$ ">��ȡ����
<%end if%>
<option style="color:orange" value="/����$ 100">���ھ���
<option style="color:black" value="/��Ǯ$ 1000">�����ֽ�
<option style="color:black" value="/����$ д����Ʒ��">������Ʒ
<option style="color:black" value="/ʹ��$ ��Ƭ��">ʹ�ÿ�Ƭ
<%if chatinfo(5)=0 then%>
<option style="color:black" value="/͵Ǯ$ ">͵ȡǮ��
<option style="color:red" value="/�¶�$ ��ҩ��">͵͵�¶�
<option style="color:red" value="/Ͷ��$ ������">Ͷ������
<option style="color:red" value="/����$ ��ʽ��">���й���
<%end if%>
<option style="color:#B000D0" value="/��Ǯ$ 0">���д�Ǯ
<option style="color:#B000D0" value="/ȡǮ$ 0">����ȡǮ
<option style="color:#B000D0" value="/ת��$ 10000">����ת��
<option style="color:blue" value="/����$ ����">�����ֵ�
<option style="color:blue" value="/����$ �鿴">�鿴�ֵ�
<option style="color:blue" value="/����$ ɾ��">���۶���
<option style="color:blue" value="/���$ Ҫ����ĺ�������">����ֵ�
<%
if jhmp="" or jhmp="����" or jhmp="��" then%>
<option value="/��������$">��������
<%end if%>
<%if chatinfo(5)=0 then%>
	<%if left(jhzy,2)="С͵" then%>
		<option style="color:blue" value="/С͵$ ">͵ȡ��Ʒ
	<%else%>
		<option style="color:blue" value="/ץС͵$ ">ץ С ͵
	<%end if %>
<%end if%>
<%if jhsf="����" then%>
<option value="/�������$">�������
<option style="color:blue" value="/������$ ��������Ϣ">�� �� ��
<option value="/���$ �����">������
<option value="/���$ ����">��ⳤ��
<option value="/���$ ����">��⻤��
<option value="/���$ ����">�������
<option value="/���$ ����">ȡ�����
<%end if
if (jhsf="����" or jhsf="����") and sjjh_grade>=4 then%>
<option style="color:blue" value="/��Ѩ$ ��д��ԭ�� ">��ѵ����
<option style="color:blue" value="/����$ 100 ">��������
<%end if
if sjjh_grade>=6 then%>
<option style="color:blue" value="/��Ѩ$ ��д��ԭ�� ">��Ѩ����
<option style="color:blue" value="/��Ѩ$ ��д��ԭ�� ">��Ѩ����
<option style="color:blue" value="/��Ѩ$ �û���">��Ѩ����
<option style="color:blue" value="/����$ ��д��ԭ�� ">�������
<option style="color:blue" value="/����$ ��д��ԭ�� ">��������
<option style="color:blue" value="/����$ ��д��ԭ�� ">���β���
<option style="color:blue" value="/��ip$ ">�鿴IP
<option style="color:blue" value="/����$ ��д��ԭ�� ">���㵷��
<option style="color:blue" value="/ը��$">�� ը ��
<%
end if
if sjjh_grade>=8 then%>
<option style="color:blue" value="/���$ ��д��ԭ�� ">�������
<option style="color:blue" value="/�ͷ�$ �û���">������
<%
end if
if sjjh_grade>=9 then%>
<option style="color:red" value="/�鿴��Ʒ$">�鿴��Ʒ
<option style="color:red" value="/������Ʒ$ ��Ƭ">����鿴
<option style="color:red" value="/�Ŵ�$ ">�Ŵ�����
<option style="color:red" value="/״̬$ ">�鿴״̬
<option style="color:red" value="/����$ ">վ������
<option style="color:red" value="/ն��$ ԭ�� ">ն���˷�
<%

%>
<html>
<head>
<title>�����������</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}
</script>
<style></style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" background="<%=chatimage%>" bgproperties="fixed">
<div align="center"><font color="#FFFFFF"><span style='font-size:9pt'><font size="-1">��</font></span><font size="-1">��������<span style='font-size:9pt'>��</span></font></font><font size="3"><br>
</font><br>
  <table cellpadding="3" cellspacing="0" border="1" bgcolor="#CCCCCC" align="center" width="100%" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr > 
      <td> 
        <div align="center"><font color="#330033" size="2">��������</font></div>
        </td>
    </tr>
    <tr > 
      <td > 
        <div align="center"><a href="javascript:s('/���ŷ�Ǯ$')" title="���š����ŷ���(�Ϲ�)�����Ͽ��Բ�����">���ŷ�Ǯ</a></div>
      </td>
    </tr>
  </table>
</div>
<p class=p1 align="center"><font color="#FFFFFF" size="2">ɱ��������<br>
��������+������<br>
<input type="checkbox" name="liji" id="liji" value="on">
�������� </font></p>
</body>
</html>
