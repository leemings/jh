<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>��Ʒ��۸�������</title>
<link rel="stylesheet" href="../setup.css">
<script LANGUAGE="JavaScript">
<!--

function FrmAddLink_onsubmit() {
if(document.FrmAddLink.moneybei.value=="")
{
alert("��������û��д���޷���ɳ���")
document.FrmAddLink.moneybei.focus()
return false
}
}

//-->
</script>
</head><LINK href="css/css.css" type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<form method=post action="wpmoney1.asp" LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
<table  border=1 cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#006699" width="229">
<tr>
<td colspan="2" height="29">
<div align="center"><font color="#FFFFFF">��Ʒ��۸�������</font></div>
</td>
</tr>
<tr>
<td width="66">
<div align="center"><font color="#FFFFFF">����</font></div>
</td>
<td width="154"><font color="#FFFFFF">
<select name="lx">
<option value="����" selected>�ֳֵ���</option>
<option value="����">�ֳֻ���</option>
<option value="����">����</option>
<option value="ͷ��">ͷ��</option>
<option value="˫��">˫��</option>
<option value="����">����</option>
<option value="װ��">װ��</option>
<option value="ҩƷ">����ҩ</option>
<option value="��ҩ">��ҩ</option>
</select>
</font></td>
</tr>
<tr>
<td width="66">
<div align="center"><font color="#FFFFFF">��������</font></div>
</td>
<td width="154"><font color="#FFFFFF" size="2">
<input type="text" name=moneybei value="0">
</font></td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<input type=submit value="ȷ ��" name="submit">
<font color="#CCCCCC">------- </font>
<input  onClick="javascript:window.document.location.href='binqi.asp'" value="�� ��" type=button name="back">
</div>
</td>
</tr>
</table>
</form>
<div align="center"><br>
ע��ѡ��װ�����ͣ�ѡ������������ȷ�¾Ϳ����Զ�������Ʒ��Ǯ��<br>
��Ʒ��Ǯ= (����+����+����+����)*��������<br>
����ڡ��塢�������ȳ��ָ��������Զ�ת����������</div>
</body></html>