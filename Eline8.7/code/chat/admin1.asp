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
if sjjh_grade<5 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���㲻������û���ʸ��������');window.close();}</script>"
	Response.End 
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>

<html>
<head>
<title>���ŷ��Ų�����һ�������wWw.happyjh.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor="#006699" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#0099CC" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#FFFF00" size="2">���Ƿ���</font></div></td>
      </tr>
    </table>  
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#336699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
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
    <tr > 
      <td id="sm">
        <div align="center"><font color="#FFFFFF" size="-1">ֻ�����Ųſɷ������</font></div>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
