<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>��E�߽�����sqlָ��ϵͳ��wWw.51eline.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<SCRIPT LANGUAGE="JavaScript">
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.sqlstr.value;
revisedTitle = addTitle;
document.PostTopic.sqlstr.value=revisedTitle;
document.PostTopic.sqlstr.focus();
return; }
</script>

<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin="0">
<p align="center"><%=Application("sjjh_chatroomname")%>���ݿ�<br>
  <br>
  <b><font color="#FF0000">ע�⣺�˴���������Ӱ�쵽���ݿ�����ϵͳ��<br>
  ���鲻���׵�վ����Ҫʹ�ã��ǳ�Σ�գ�</font></b> </p>
<p align="left">��������Ҫִ�е�ָ�<br>
  �磺update �û� set grade=1,���='����' where ���&lt;&gt;'����' and ����&lt;&gt;'�ٸ�'<br>
  �ǽ�������ݲ�Ϊ���ţ����ɲ�Ϊ�ٸ��Ĺ���ȼ����1��,������ó�Ϊ���ӣ��� <br>
  �磺delete * from w where i=0<br>
  �ǽ���Ʒ���ݿ��е���ƷΪ0�ļ�¼ɾ����<br>
  �磺delete * from �书 where ����='�ٸ�'<br>
  �ǽ��ٸ����书ɾ����<br>
  �磺delete * from �û� where allvalue&lt;=2 and month(��¼)&lt;&gt;month(now())<br>
  �ǽ��Ǳ���,�ݵ�С�ڵ���2���û�ɾ����<br>
<font color=#cc0000>������Ǯ��SQL</font> <br>
update �û� set ���=���+150000000 where grade=5 <br>
<font color=#cc0000>��������1000�����</font> <br>
update �û� set allvalue=allvalue+1000 <br>
<font color=#cc0000>�ϲ�����</font> <br>
update �û� set ����='������A' where ����='������B'  <br>

����˵�� ������B �ϲ������� A��  <br>
ִ�����֮ǰ��ִ��  <br>
update �û� set grade=1, ���='��' where ����='����A'���='����' <br> 
update �û� set grade=1, ���='��' where ����='����A'���='����'  <br>
update �û� set grade=1, ���='��' where ����='����A'���='����'  <br>
update �û� set grade=1, ���='��' where ����='����B'���='����'  <br>
update �û� set grade=1, ���='��' where ����='����B'���='����'  <br>
update �û� set grade=1, ���='��' where ����='����B'���='����'  <br>
update �û� set grade=1, ���='��' where ����='����A'���='����'  <br>
update �û� set grade=1, ���='��' where ����='����A'���='����'  <br>
�����ִ�л����ܶ೤������������ <br>
�ٰѱ��ϲ������ɵ�����Ȩ���û�����OK�� <br>

<font color=#cc0000>�õȼ��ߵĽ��������ֶ����Լ�����</font> <br>
update �û� set allvalue=174050 where allvalue>174050   <br>   
�����û�=59��                       ִ���������ȼ�����59 <br>
��һ���û� <br>
update �û� set allvalue=174050 where id=���� <br>

<font color=#cc0000>���г���1��Ǯ���˵�Ǯ����1��</font> <br>
update �û� set ����=0,���=100000000 where ����<>'�ٸ�' and ����+���>100000000 <br>
<font color=#cc0000>��ĳ��һ�����齣</font> <br>
update �û� set W3='���齣|1;' where ����='С��ʮһ' <br>

<font color=#cc0000>��ĳ��һ��������������Ҫע��Ŷ����Ȼ���ġ�</font> <br>
update �û� set zw='����|����|2002-7-22 21:05:06|75000|21000|14000|4|2002-9-20 4:13:49' where ����='С��ʮһ' <br>

<font color=#cc0000>��ĳ��һ��С����ע������żҲҪ��</font> <br>
update �û� set boy='����|baby|2003-5-4 18:44:26|235|40830|11380|20|2003-10-19 21:38:32',boysex='images/boy.gif' where ����='һ����' <br>

<font color=#cc0000>���ip�⿪</font> <br>
update �û�  set ״̬='����',�¼�ԭ��='��' where lastip='����ip'<br>

<font color=#cc0000>�ᱦ��ʼ��</font><br>
update �ᱦ set �ھ�=0,��ȡ=false,��������=0,����=false,ʱ��=now() where ����=true  <br>

<font color=#cc0000>�»��ֽ���</font><br>
update �û� set ��Ա��=��Ա��+30,���=���+100,����=����+200,����=����+100000000,allvalue=allvalue+1000,��=��+500,ľ=ľ+500,ˮ=ˮ+500,��=��+500,��=��+500 where mvalue>=?(�鿴���а�) <br>

<font color=#cc0000>�ܻ��ֽ���</font><br>
update �û� set ��Ա��=��Ա��+50,���=���+200,����=����+500,����=����+500000000,allvalue=allvalue+2000,��=��+1000,ľ=ľ+1000,ˮ=ˮ+1000,��=��+1000,��=��+1000 where allvalue>=?(�鿴���а�) <br>

  <br>
</p>
<%if sjjh_name=Application("sjjh_user") then%>
<form method="POST" action="sqlcommok.asp" name="PostTopic">
 <div align="center">
<SELECT name=fs onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">����ָ��</OPTION> 
<OPTION value="update b set o=1000 where a='��Ʒ��' and n<>'��'">[�޸ľ�������]</OPTION>
<OPTION value="update �û� set ����=0,���=0 where (����+���)>1000000000">[����ʲ�����10���û��ʲ�Ϊ0]</OPTION>
<OPTION value="update �û� set grade=1,���='����' where ����='����'">[��������ɹ�����Ա]</OPTION>
<OPTION value="delete * from y where b='�ٸ�'" >[ɾ���ٸ��书]</OPTION>
<OPTION value="delete * from l " >[ɾ�����в�����¼]</OPTION>
<OPTION value="delete * from l where d='����'" >[ɾ�����н���������¼]</OPTION>
<OPTION value="delete * from j " >[ɾ�������ռ�����]</OPTION>
<OPTION value="delete * from h where c='���'" >[ɾ���������ϼ�¼]</OPTION>
<OPTION value="delete * from h where c='���'" >[ɾ����������¼]</OPTION>
<OPTION value="delete * from t where b='����'" >[ɾ���������ĺ��弼]</OPTION>
<OPTION value="delete * from �û� where allvalue<=0" >[ɾ�����allvalueΪ0���û�]</OPTION>
<OPTION value="delete * from �û� where ��Ա��>500" >[ɾ����Ա�𿨴���500���û�]</OPTION>
</SELECT><br>
    <br>
    ������ָ� 
    <input type="text" name="sqlstr" size="50" maxlength="250">
    <br>
    <input type="submit" value="ִ��" name="B1" class="p9">
    <input type="reset" name="Submit" value="���">
  </div>
  </form>
 <%else%>
  <div align="center">������Ȩվ���������ֹʹ�ã�</div>
 <%end if%>
<p align="center">��E�߽�����</p>
</body>
</html>