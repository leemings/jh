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
<title>�û����ݿ�����wWw.51eline.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>
<SCRIPT LANGUAGE="JavaScript">
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.tiaojian.value;
revisedTitle = addTitle;
document.PostTopic.tiaojian.value=revisedTitle;
document.PostTopic.tiaojian.focus();
return; }
function DoTitlel(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.show.value;
revisedTitle = addTitle;
document.PostTopic.show.value=revisedTitle;
document.PostTopic.show.focus();
return; }

</script>
<body bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin="0">
<p align="center"><%=Application("sjjh_chatroomname")%>���ݿ�</p>
<p align="left">������������ѯ�û���<br>
  �磺���='����' ��ѯ���Ϊ���ŵ��û���<br>
  ���磺 ����&gt;10000 and �书&gt;10000 �鿴��������10000�����书����10000���û���<br>
  �ڲ�ѯ�п���ʹ�ã�&quot;and&quot; &quot;or&quot; &quot;&gt;&quot; &quot;&lt;&quot; &quot;&lt;&gt;&quot; 
  &quot;&gt;=&quot; &quot;&lt;=&quot; &quot;=&quot; ��ϵ����<br>
  �ڲ�ѯ��Ʒʱ����ѯ�ֶε�ֵ��Ч�������룺ӵ����='СС��'<br>
  <font color="#0000FF">��ѯ�ֶΣ���ֵΪ��һֵ,�磺�����������ȡ�����Ϊ�������� and ���� �� ����,������</font></p>
<div align="center">�û������޸ĳ��� </div>
<form method="POST" action="seeuserok.asp" name="PostTopic">
  <div align="center"><br>
    �������ѯ������
<SELECT name=fs onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">ѡ���ѯ</OPTION> 
<OPTION value="��Ա=True">[��ѯ�ݵ��ƻ�Ա]</OPTION>
<OPTION value="��Ա�ȼ�<>0">[��ѯ�ȼ��ƻ�Ա]</OPTION>
<OPTION value="����='�ٸ�' or grade>=6">[�ٸ���Ա]</OPTION>
<OPTION value="(����+���)>1000000000">[�ʲ�10��]</OPTION>
<OPTION value="�书>80000">[�书��8��]</OPTION>
<OPTION value="grade>2 and grade<=5" >[���ɹ���Ա]</OPTION>
<OPTION value="���='����' or grade=5">[��������]</OPTION>
<OPTION value="�ȼ�>50" >[ս����50]</OPTION>
<OPTION value="regip='�����ѯ��ip��ַ' or lastip='�����ѯ��ip��ַ'" >[ip��ѯ]</OPTION>
<OPTION value="��Ա��>0">[��Ա��]</OPTION>
<OPTION value="ʦ��='СС��'">[ʦ����ѯ]</OPTION>
<OPTION value="���>=100">[��Ҵ���100]</OPTION>
<OPTION value="��Ա�ȼ�=3">[��ѯ3����Ա]</OPTION>
<OPTION value="��Ա�ȼ�>0">[��ѯ���л�Ա]</OPTION>
</SELECT>
    <input type="text" name="tiaojian" size="50" maxlength="250">
    <br>
    ���뽫��ѯ�ֶΣ� 
    <input type="text" name="show" size="10" maxlength="12">
<SELECT name=fs onchange=DoTitlel(this.options[this.selectedIndex].value)> 
<OPTION selected value="">��ѯ����</OPTION> 
<OPTION value="����">[����]</OPTION>
<OPTION value="���">[���]</OPTION>
<OPTION value="��Ա����">[��Ա����]</OPTION>
<OPTION value="��Ա����">[��Ա����]</OPTION>
<OPTION value="lasttime">[�������ʱ��]</OPTION>
<OPTION value="������">[������]</OPTION>
<OPTION value="����">[����]</OPTION>
</SELECT>

    <br>
    <input type="submit" value="��ѯ" name="B1" class="p9">
    <input type="reset" name="Submit" value="���">
  </div>
  </form>
<p align="center">��E�߽�����</p>
</body>
</html>