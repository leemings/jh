<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade<5 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻������û���ʸ��������');window.close();}</script>"
	Response.End 
end if
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>

<html>
<head>
<title>���ŷ��Ų�����</title>
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
      <td id="sm" height="22"> 
        <div align="center"><font size="-1"><a href="fafang1.asp?cz=Ͱ��&value=5000" target="d" title="�ͷ�κ������ֻ��κ���˿��Դ�!">Ͱ��</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
         <div align="center"><font size="-1"><a href="fafang1.asp?cz=ľ��&value=5000" target="d" title="�ͷ��������ֻ������˿��Դ�!">ľ��</a></font></div>
      </td>
    </tr>
    <tr > 
      <td id="sm" height="22"> 
       <div align="center"><font size="-1"><a href="fafang1.asp?cz=����&value=5000" target="d" title="�ͷ��������ֻ������˿��Դ�!">����</a></font></div>
      </td>
    </tr>
  </table>
</div>
</body>
</html>