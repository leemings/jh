<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
session.Abandon
id=Trim(Request.QueryString("id"))
Select Case id
Case "000"
	errormsg="�����ڼ������رգ�����ԭ�������볢�����INTERNET��ʱ�ļ�"
Case "001"
	errormsg="�����ڼ������رա�ԭ��ʱ��"
Case "002"
	errormsg="�����ڼ������رա�ԭ������Ϊ�㱻�ɶ��"&Request.QueryString("from")&"�����ˣ�"
Case "003"
	errormsg="�����ڼ������رա�ԭ������Ϊ�������ж���������"
Case "004"
	errormsg="��û�е�¼��ʱ�Ͽ����ӣ������µ�¼��"
Case "005"
	errormsg="�����ڼ������رա�ԭ������Ϊ�㱻�ɶ��"&Request.QueryString("from")&"���������"
Case "006"
	errormsg="�����ڼ������رա�ԭ������Ϊ�㱻�ɶ��"&Request.QueryString("from")&"���������ˣ�����취����ͨ��һ��ȥ�ɣ�"	
Case "007"
	errormsg="�����ڼ������رա�ԭ������Ϊ�ɶ��"&Request.QueryString("from")&"���������񼫣�������Ͷ�������̷�ר�ü�����"	
Case "008"
	errormsg="�����ڼ������رա�ԭ������Ϊ���IP�����������Զ��������"	
Case else
	errormsg="�Բ��𣬸ó�������δ���Ǽǡ�"
End Select
%>
<html>
<head><link rel="stylesheet" href="css.css" type="text/css">
<script language=javascript>if(window!=window.top)top.location.href=location.href;</script>
<title> �� �� �� ʾ</title>
</head>
<body oncontextmenu=self.event.returnValue=false background="<%=bgimage%>" bgcolor="<%=bgcolor%>" topmargin="150" leftmargin="200">
<table width="350" border="0" cellspacing="2" cellpadding="4">
  <tr> 
    <td align="center"><img src="../images/error.gif" width="32" height="32"> 
      <%=errormsg%></td>
  </tr>
  <tr> 
    <td align="center" valign="middle"> 
      <input type="button"  value="ȷ��" onClick=javascript:window.close() name="button">
    </td>
  </tr>
</table>
</body>
</html>
