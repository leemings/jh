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
if aqjh_grade<7 then
	Response.Write "<script Language=Javascript>alert('����ȼ���������Ҫ�ż�����!');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<html>
<head>
<title>�ٸ��Ź�</title>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:#ccccff;text-decoration:none;}
a:hover{color:#ffffff;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed">
<div align="center"><br><table width="70%" border="1" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="0" align="center">
    <tr > 
      <td id="sm"> 
        <div align="center"><a href="#" onClick="window.open('listbak.asp','listbak','scrollbars=yes,resizable=yes,width=400,height=250')">��Ǯ</a></div>
      </td>
    </tr>    <tr > 
      <td id="sm"> 
        <div align="center"><a href="sendnl.asp" target="_blank">����</a></div>
      </td>
    </tr><tr> 
      <td id="sm"> 
        <div align="center"><a href="sendfy.asp" target="_blank">����</a></div>
      </td>
    </tr><tr> 
      <td id="sm"> 
        <div align="center"><a href="sendtl.asp" target="_blank">����</a></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=�ϻ�&value=10" target="d" title="�ų��ϻ����ô�Ҵ�!">�ϻ�</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=˥��&value=10000" target="d" title="����˥�磬�ֲܿ����������1���!">˥��</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=��ʬ&value=10000" target="d" title="���Ž�ʬ���ֲܿ����������1���!">��ʬ</a></font></div>
      </td>
    </tr>
<tr > 
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=��ǹ&value=10000" target="d" title="������ǹ��ҪС��!">��ǹ</a></font></div>
      </td>
    </tr>
	 <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=����&value=50000" target="d" title="���Ź������������5���!">����</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=����&value=20000" target="d" title="�ų������ô�Ҵ�!">����</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=����&value=50000" target="d" title="���ŷ����������5���!">����</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=��ʿ&value=50000" target="d" title="������ʿ���������5���!">��ʿ</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=�˹�����&value=50000" target="d" title="�˹����������й����������5����!">����</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=����&value=50000" target="d" title="�������ˣ��������5����!">����</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=��Ů&value=50000" target="d" title="������Ů���������5���!">��Ů</a></font></div></td>
    </tr>
    <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=����&value=50000" target="d" title="�������У��������5���!">����</a></font></div></td>
    </tr>
	<tr > 
      <td id="sm" height="22"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=��å&value=100000" target="d" title="������磬�ֲܿ����������10���!">��å</a></font></div></td>
    </tr>
 <tr > 
      <td id="sm"> 
        <div align="center"><font size="-1"><a href="fafang2.asp?cz=ƻ��&value=10000" target="d" title="ƻ����������!">ƻ��</a></font></div>
      </td>
    </tr>
	 <tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=�̻�&value=50000" target="d" title="�����̻����������5���!">�̻�</a></font></div></td>
    </tr>
	<tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=ʨ��&value=50000" target="d" title="����ʨ�ӣ��������5���!">ʨ��</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=�ǵ�&value=50000" target="d" title="�ǵ䲡��5���!">�ǵ�</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=�˷�&value=50000" target="d" title="�˷�������5���!">�˷�</a></font></div></td>
    </tr>
<tr > 
      <td id="sm"> <div align="center"><font size="-1"><a href="fafang2.asp?cz=С͵&value=50000" target="d" title="С͵����5���!">С͵</a></font></div></td>
    </tr>
  </table>
</div>
</body></html>