<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>�û����ݿ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><%=Application("aqjh_chatroomname")%>���ݿ�</p>
<p align="left">����˵����<br>
  <font color="#0000FF">����ս���ȼ����㹫ʽ���ȼ�x�ȼ�x50���磺��һ��ս��10�����û���10x10x50=5000�����ֵд�뵽���ݿ�ģ�<b>[<font color="#FF0000">�ܻ���</font></b>]��������û���һ�ν���ʱս���ȼ����Զ���������</font><font color="#FF0000"><br>
  </font><font color="#0000FF">��Աʱ���ǽ�ֹʱ�䣬��Ա�ȼ�Ϊ��1,2,3,4�����ֵΪ0�����ʾ���ǻ�Ա��</font><br>
  <font color="#FF0000"> �书�ӣ������ӣ������ӣ������ӵȡ���һЩֵ�ǲ�Ҫ�Ҹĵģ���Щ�������ڽ����ϴ�������Ŀ�õ�����Щֵ��Ҫ�Ķ���<br>
  <font color="#000000">�����ݷ��ƻ�Ա˵����</font> <br>
  [<b><font color="#0000FF">��Ա</font></b>]����ΪTrue ʱ��Ϊ�ݷ��ƻ�Ա��[<b><font color="#0000FF">��Ա����</font></b>]�����ݷ��ƻ�Ա�Ľ���ʱ�䡣�й��ݷ��ƻ�ԱĬ��Ϊ2����Ҫ���޸Ŀ����޸��ݷ��ļ��õ���</font></p>
<table width=80% align=center><tr><td>
<form method="POST" action="showuser.asp">
    <select name="sjcz">
      <option value="����" selected>����</option>
      <option value="ID">ID</option>
      <option value="��ż">��ż</option>
      <option value="����">����</option>
      <option value="Oicq">OIcq</option>
    </select>
    <input type="text" name="search" size="10" maxlength="10">
    <input type="submit" value="��ѯ" name="B1" class="p9">
    <input type="reset" name="Submit" value="���"> [�û����ϲ�ѯ]
</form>
<form method="POST" action="setzt.asp">
    <select name="t1">
      <option value="����" selected>����</option>
      <option value="���">���</option>
      <option value="���">���</option>
      <option value="��Ա��">��</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="�书">�书</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="ת��">ת��</option>
      <option value="����">����</option>
      <option value="����">����</option>
      <option value="�Ṧ">�Ṧ</option>
    </select>
    ������<input type="text" name="name" size="10" maxlength="10">
    <select name="t2">
      <option value="����" selected>����</option>
      <option value="����">����</option> 
</select>
<input type="text" name="sl" size="10" maxlength="9">
	<input type="submit" value="ȷ��" name="B1" class="p9">
    <input type="reset" name="Submit" value="���"> [����״̬����]
</form>
<form method="POST" action="laren.asp">
    <input type="text" name="larenseek" size="10" maxlength="10">
    <input type="submit" value="��ѯ" name="B12" class="p9">
    <input type="reset" name="Submit2" value="���"> [���˲鿴����]
</form>
<form method="POST" action="chuwu.asp">
    <input type="text" name="chuwuseek" size="10" maxlength="10">
    <input type="submit" value="��ѯ" name="B12" class="p9">
    <input type="reset" name="Submit2" value="���"> [�鿴�����䣬��ID��ѯ]
</form>
<form method="POST" action="pass.asp">
    <input type="text" name="cpass" size="10" maxlength="10">
    <input type="submit" value="�޸�" name="B12" class="p9">
    <input type="reset" name="Submit2" value="���"> [������������¼����ĳɣ�123456]
</form>
<form method="POST" action="pass2.asp">
    <input type="text" name="cpass" size="10" maxlength="10">
    <input type="submit" value="�޸�" name="B12" class="p9">
    <input type="reset" name="Submit2" value="���"> [�����������ڶ�����ĳɣ�123456]
</form>
<form method="POST" action="yuejlok.asp">
    <input type="text" name="jl_name" size="10" maxlength="10">
    <input type="submit" value="ȷ��" name="B12" class="p9">
    <input type="reset" name="Submit2" value="���"> [���½���(�����»��ֺ���������)]
</form></td></tr></table></body></html>