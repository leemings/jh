<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn
if rs("grade")<10 then
   Response.Write "<script Language=Javascript>alert('�㲻��վ�����ܹ�������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<html>
<head>
<title>�ٸ����Ų�����</title>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:#ccccff;text-decoration:none;}
a:hover{color:#ffffff;text-decoration:none;CURSOR:url('1.cur');}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed">
<div align="center"><br><table width="70%" border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=����&value=1000000" target="d" title="�������������ֵΪ10����!">����</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=Ԫ��&value=300000" target="d" title="����Ԫ�����������1��!">Ԫ��</a></div>
      </td>
    </tr>
<tr >
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=����&value=100" target="d" title="���ŵ��£����ֵΪ100��!">����</a></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><a href="fafang.asp?cz=����&value=100" target="d" title="�������������ֵΪ100��!">����</a></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><a href="fafang.asp?cz=����&value=100" target="d" title="���ŷ��������ֵΪ100��!">����</a></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><a href="fafang.asp?cz=�Ṧ&value=100" target="d" title="�����Ṧ�����ֵΪ100��!">�Ṧ</a></div>
      </td>
    </tr>
    <tr >
      <td id="sm">
        <div align="center"><a href="fafang.asp?cz=��&value=50" target="d" title="���Ž����ԣ����ֵΪ50��!">������</a></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=ľ&value=50" target="d" title="����ľ���ԣ����ֵΪ50��!">ľ����</a></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=ˮ&value=50" target="d" title="����ˮ���ԣ����ֵΪ50��!">ˮ����</a></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=��&value=50" target="d" title="���Ż����ԣ����ֵΪ50��!">������</a></div>
</td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=��&value=50" target="d" title="���������ԣ����ֵΪ50��!">������</a></div>
      </td>
    </tr>
	<tr > 
      <td id="sm" height="22"> 
        <div align="center"><a href="fafang.asp?cz=����&value=100000" target="d" title="���ŵ��⣬����10������!">����</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=����&value=5000" target="d" title="������������߷���5ǧ��!">����</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=����&value=5000" target="d" title="������������߷���5ǧ��!">����</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=�书&value=5000" target="d" title="�����书����߷���5ǧ��!">�书</a></div>
      </td>
    </tr>
	
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=�ϻ�&value=10" target="d" title="�ų��ϻ����ô�Ҵ�!">�ϻ�</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><a href="fafang.asp?cz=˥��&value=10000" target="d" title="����˥�磬�ֲܿ����������1���!">˥��</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><a href="fafang.asp?cz=��ʬ&value=10000" target="d" title="���Ž�ʬ���ֲܿ����������1���!">��ʬ</a></div>
      </td>
    </tr>
<tr > 
      <td id="sm" height="22"> 
        <div align="center"><a href="fafang.asp?cz=��ǹ&value=10000" target="d" title="������ǹ��ҪС��!">��ǹ</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=�ݶ�����&value=100" target="d" title="���Ŷ��㣬���ֵΪ100��!">����</a></div>
      </td>
    <tr> 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=������&value=50" target="d" title="�����㣬�������50!">������</a></div>
      </td>
    </tr>
	 <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=����&value=50000" target="d" title="���Ź������������5���!">����</a></div></td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=����&value=20000" target="d" title="�ų������ô�Ҵ�!">����</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=����&value=50000" target="d" title="���ŷ����������5���!">����</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=��ʿ&value=50000" target="d" title="������ʿ���������5���!">��ʿ</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=�˹�����&value=50000" target="d" title="�˹����������й����������5����!">����</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=����&value=50000" target="d" title="�������ˣ��������5����!">����</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=��Ů&value=50000" target="d" title="������Ů���������5���!">��Ů</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=����&value=50000" target="d" title="�������У��������5���!">����</a></div></td>
    </tr>
	<tr > 
      <td id="sm" height="22"> <div align="center"><a href="fafang.asp?cz=��å&value=100000" target="d" title="������磬�ֲܿ����������10���!">��å</a></div></td>
    </tr>
	<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=��Ƭ&value=1" target="d" title="�ſ�Ƭ!">��Ƭ</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=��ҩ&value=1" target="d" title="�Ų�ҩ!">��ҩ</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=��ҩ&value=1" target="d" title="����ҩ!">��ҩ</a></div>
      </td>
    </tr>
 <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=ƻ��&value=10000" target="d" title="ƻ����������!">ƻ��</a></div>
      </td>
    </tr>
	<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=����&value=1" target="d" title="������!">����</a></div>
      </td>
    </tr><tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=�±�&value=1" target="d" title="�����±�">�����±�</a></div>
      </td>
    </tr>
 <tr > 
      <td id="sm" height="22"> <div align="center"><a href="fafang.asp?cz=����&value=50000" target="d" title="���ű���������5������!">����</a></div></td>
    </tr>
	<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=վ��&value=90000" target="d" title="���ź��䣬����10��!">վ��</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=�����&value=50" target="d" title="���ݵ㣬�������50!">�����</a></div>
      </td>
    </tr>
	 <tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=�̻�&value=50000" target="d" title="�����̻����������5���!">�̻�</a></div></td>
    </tr>
	<tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=ʨ��&value=50000" target="d" title="����ʨ�ӣ��������5���!">ʨ��</a></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=�ǵ�&value=50000" target="d" title="�ǵ䲡��5���!">�ǵ�</a></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=�˷�&value=50000" target="d" title="�˷�������5���!">�˷�</a></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><a href="fafang.asp?cz=С͵&value=50000" target="d" title="С͵����5���!">С͵</a></div></td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=���&value=10" target="d" title="���Ž�ң����ֵΪ10��!">���</a></div>
     
      </td>
    </tr>
	<tr > 
      <td id="sm"> 
        <div align="center"><a href="fafang.asp?cz=����&value=5" target="d" title="���Ż�Ա�𿨣����ֵΪ5��!">����</a></div>
      </td>
    </tr>
 <tr>
      <td id="sm"> 
        <div align="center"><a href="sen.asp" target="d" title="���ŵȼ�">�ȼ�</a></div>
      </td>
    </tr>
 <tr>
      <td id="sm"> 
        <div align="center"><a href="#" onClick="window.open('zmbk.asp','listbak','scrollbars=yes,resizable=yes,width=400,height=250')">����</a></div>
      </td>
    </tr>
  </table></div></body></html>