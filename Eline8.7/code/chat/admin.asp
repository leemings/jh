<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if sjjh_grade<10 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ʲô����ֻ��վ���ſ��Բ�����');window.close();}</script>"
	Response.End 
end if
if session("advtime")<>"" then
if session("advtime")>now()-12 then  Response.Redirect "../error.asp?id=492"
end if
session("advtime")=now()
%>
<html>
<head>
<title>վ�����Ų�����һ�������wWw.51eline.com��</title>
<style>
body{
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#993300" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#FFFF00" size="2">վ������</font></div></td>
      </tr>
    </table>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#9999CC" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr > 
      <td id="sm"> 
        <div align="center"><font color="#FFFFFF" size="-1"><a href="fafang.asp?cz=����&value=10000000" target="d" title="�������������ֵΪ1000����!">����</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=���&value=3" target="d" title="���Ž�ң����ֵΪ3��!">���</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=10" target="d" title="�������ң����ֵΪ10��!">����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=1000" target="d" title="���ŵ��£����ֵΪ1000��!">����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=1000" target="d" title="�������������ֵΪ1000��!">����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=1000" target="d" title="���ŷ��������ֵΪ1000��!">����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=�Ṧ&value=1000" target="d" title="�����Ṧ�����ֵΪ1000��!">�Ṧ</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=8000" target="d" title="�ų��������ô�Ҵ�!">����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=20000" target="d" title="�ų������ô�Ҵ�!">����</a></font></div>
      </td>
    </tr>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=10000" target="d" title="���Ź����������5���!">����</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=��ʿ&value=50000" target="d" title="�ų���ʿ���ô�Ҵ�!">��ʿ</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=�˹�����&value=80000" target="d" title="�ų��˹��������ô�Ҵ�!">����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=��Ů&value=10000" target="d" title="�ų���Ů���ô�Ҵ�!">��Ů</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=10000" target="d" title="�ų����У��ô�Ҵ�!">����</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> <div align="center"><font size="-1"><a href="fafang.asp?cz=��å&value=30000" target="d" title="������磬�ֲܿ����������3���!">��å</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=��ҩ&value=5" target="d" title="���Ų�ҩ�����ֵΪ5��!">��ҩ</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang.asp?cz=��ҩ&value=1" target="d" title="����ҩ!">��ҩ</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm" height="22"> <div align="center"><font size="-1"><a href="fafang.asp?cz=ƻ��&value=9000" target="d" title="����ƻ��������9000����!">ƻ��</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm" height="22"> <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=50000" target="d" title="���ű���������5������!">����</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=1" target="d" title="������!">����</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=allvalue&value=100" target="d" title="���Ŵ�㣬���ֵΪ100��!">���</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=�ݶ�����&value=100" target="d" title="���Ŷ��㣬���ֵΪ100��!">����</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=�ϻ�&value=8000" target="d" title="�ų��ϻ����ô�Ҵ�!">�ϻ�</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
      <div align="center"><font size="-1"><a href="fafang.asp?cz=��ǹ&value=8000" target="d" title="������ǹ���ֲܿ����������1���!">��ǹ</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=��ʬ&value=8000" target="d" title="���Ž�ʬ���ֲܿ����������1���!">��ʬ</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=˥��&value=8000" target="d" title="����˥�磬�ֲܿ����������1���!">˥��</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=Ԫ��&value=100000" target="d" title="����Ԫ�����������10��!">Ԫ��</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=5000" target="d" title="������������߷���5ǧ��!">����</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=5000" target="d" title="������������߷���5ǧ��!">����</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=�书&value=5000" target="d" title="�����书����߷���5ǧ��!">�书</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=�����&value=100" target="d" title="��������㣬���ֵΪ100��!">�����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=������&value=100" target="d" title="���������㣬���ֵΪ100��!">������</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=����&value=2" target="d" title="�������𿨣����ֵΪ2Ԫ!">����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=��&value=200" target="d" title="���Ž����ֵΪ500��!">������</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=ľ&value=200" target="d" title="����ľ�����ֵΪ200��!">ľ����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=ˮ&value=200" target="d" title="����ˮ�����ֵΪ200��!">ˮ����</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=��&value=200" target="d" title="���Ż����ֵΪ200��!">������</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang.asp?cz=��&value=200" target="d" title="�����������ֵΪ200��!">������</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm">
        <div align="center"><font color="#FF0000" size="-1">����㷢��.���̫��</font></div>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
