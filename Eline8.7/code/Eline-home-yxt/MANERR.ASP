<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
id=Trim(Request.QueryString("id"))
chatroombgimage=Application("hxf_c_chatroombgimage")
chatroombgcolor=Application("hxf_c_chatroombgcolor")
nl=""
Select Case id
Case "000"
 nl="��������Ŀ¼��������Ŀ¼���������д���Global.asa û�б�ִ�С���������Ҫ����Ŀ¼��֧�֣�"
Case "100"
 nl="�Բ�������δ��¼���Ѿ���ʱ�Ͽ����ӣ���ص���¼ҳ�����������û�����������е�¼��"
Case "101"
 nl="����ӵ�� 3 ������Ȩ�޲�������Դ��Ķ�����"
Case "102"
 nl="������Ŀ�����ܷſհף�����д������"
Case "103"
 nl="�������ݱ������ԡ�//����ͷ����䣡"
Case "104"
 nl="�������ݲ��ܳ��֡�//##�����⽫���·��Ժ����������ظ����뽫 ## ȥ����"
Case "105"
 nl="�������ƻ������ݳ������Ƶĳ��ȣ�"
Case "106"
 nl="�������ݺ��С�%%������������Ӧ��Ϊ��1 �Ա��ˡ���"
Case "107"
 nl="��������Ϊ��1 �Ա��ˡ�������������ȴû�г��֡�%%����"
Case "108"
 nl="��ӵĶ��������У����ܰ�����ǵġ�\��������������" & chr(34) & "����"
Case "109"
 nl="�������Ѿ����ڣ������ظ���ӡ�"
Case "120"
 nl="��û���������񹫿�����Ȩ�ޡ�"
Case "121"
 nl="û�г��� 7 ��ļ�¼�����������"
Case "130"
 nl="��û�С��ʺŹ�����Ȩ�ޡ�"
Case "131"
 nl="û�д����ʺſɹ�ɾ����"
Case "132"
 nl="�������û�����"
Case "133"
 nl="���ѡ���ɱ�����ʺ���û���ҵ����û��������ָܻ���"
Case "134"
 nl="���ָܻ��û�����<font color=red>" & Request.QueryString("name") & "</font>����Ϊϵͳ��������ͬ���û������ڡ������ȷʵ��Ҫ�ָ����û����������á�ɾ���ʺš����ܣ���ɾ��ϵͳ����ͬ���û������ٻָ����û�����"
Case "135"
 nl="��������ɾ�����û�����"
Case "136"
 nl="�û��������ڣ�����ɾ����"
Case "137"
 nl="�û��������ڡ�"
Case "138"
 nl="���û����������ñ����������ظ�������"
Case "139"
 nl="���û���δ�����ñ���������ȡ����"
Case "140"
 nl="���û��������û���������Ϊ�ա�"
Case "141"
 nl="���ܽ��û�����Ϊ��<font color=red>" & Request.QueryString("name") & "</font>����Ϊϵͳ��������ͬ���û������ڡ������ȷʵ��Ҫ�ĳ�����û����������á�ɾ���ʺš����ܣ���ɾ��ϵͳ����ͬ���û�����"
Case "142"
 nl="�û����������������Ϊ�ա�"
Case "150"
 nl="��û�С�����ѹ������Ȩ�ޡ�"
Case "151"
 nl="���ݿ���δ�رգ��������´����ݿ⡣"
Case "152"
 nl="���ݿ���δ�򿪣����عر����ݿ⡣"
Case "160"
 nl="������Ҫ�������û�����"
Case "200"
 nl="��û�С����ˡ���Ȩ�ޡ�"
Case "201"
 nl="�㲻���������У�����ִ�С����ˡ�������"
Case "202"
 nl="��ָ��Ҫ�߳��Ķ���"
Case "203"
 nl="�����޹����ˣ�������ԭ��"
Case "204"
 nl="�û�����<font color=red>" & Request.QueryString("kickname") & "</font> �����������У����������ˡ�"
Case "205"
 nl="�Ǻǣ�ë�������Լ��氡��"
Case "210"
 nl="��û�С�IP������Ȩ�ޡ�"
Case "211"
 nl="�㲻���������У����ܽ��С�IP������"
Case "212"
 nl="��ָ��Ҫ������ IP��"
Case "213"
 nl="�����Լ��� IP����ɵ�ˣ�"
Case "214"
 nl="�����޹ʷ��� IP��������ԭ��"
Case "215"
 nl="�� IP �Ѿ��������ˣ������ظ�������"
Case "216"
 nl="����������� IP ��ԭ��"
Case "217"
 nl="�� IP δ�����������ܽ��н�����"
Case "218"
 nl="��ָ��Ҫ������ IP��"
Case "219"
 nl="Ҫ������IP���û�������Ӧ��"
Case "220"
 nl="��û�С�ը����������Ȩ�ޡ�"
Case "221"
 nl="�㲻���������У����ܽ��С�ը����������"
Case "222"
 nl="��ָ��Ҫ��ը�Ķ���"
Case "223"
 nl="�����㲻�ǲ������������������ɡ�"
Case "224"
 nl="�����޹�����ը����������ԭ��"
Case "225"
 nl="�û�����<font color=red>" & Request.QueryString("bombname") & "</font> �����������У�ը�����ˡ�"
Case "230"
 nl="��û�и��ġ�ϵͳ��������Ȩ�ޡ�"
Case "231"
 nl="��ֵ���ֵ��ȫ��ͬ�����ؽ��и��ġ�"
Case "232"
 nl="�����ֵ���Ϸ���"
Case "240"
 nl="��û�н��С����������Ȩ�ޡ�"
Case "241"
 nl="�㲻���������У����ܽ��С����������"
Case "242"
 nl="��û��ִ�С�������������Ȩ�ޡ�"
Case "243"
 nl="��û��ִ�С�������������Ȩ�ޡ�"
Case "244"
 nl="�û�������Ϊ�ա�"
Case "245"
 nl="�Ҳ����û�����<font color=red>" & Request.QueryString("username") & "</font>��"
Case "246"
 nl="���û����ļ��𲻱���ͣ����ܶ�����в�����"
Case "247"
 nl="ѡ���ļ���ֵ���Ϸ���"
Case "248"
 nl="������ԭ��"
Case "250"
 nl="��û�в鿴��HTML���á���Ȩ�ޡ�"
Case "255"
 nl="��û�и��ġ�վ�����桱��Ȩ�ޡ�"
Case "256"
 nl="���ݲ���Ϊ�ա�"
Case "260"
 nl="��û�С�����������Ȩ�ޡ�"
Case "261"
 nl="�Ҳ����ö�����"
Case "270"
 nl="��û�С����Թ�����Ȩ�ޡ�"
Case "271"
 nl="�����Բ����ڡ�"
Case "280"
 nl="��û�С��������𡱵�Ȩ�ޡ�"
Case "281"
 nl="�㲻���������в��ܡ��������𡱡�"
Case "282"
 nl="�û�������Ϊ�ա�"
Case "283"
 nl="��������������ԭ��"
Case "300"
 nl="��û�й���ͶƱϵͳ����Ȩ�ޡ�"
Case "301"
 nl="������ͶƱ����������ͶƱ��"
Case "302"
 nl="��ʽ����"
Case "303"
 nl="��ѡ����֧�ֵĺ�ѡ�ˡ�"
Case "304"
 nl="��ѡ�˲����ڡ�"
Case "305"
 nl="��ѡ�˲���Ϊ�ա�"
Case "306"
 nl="��ѡ���Ѿ����ڣ������ظ���ӡ�"
Case "310"
 nl="��û�С����÷������ɣе�Ȩ�ޡ�"
Case "311"
 nl="��ʽ����"
Case "320"
 nl="IP����Ϊ�ա�"
Case else
 nl="�Բ��𣬸ó�������δ���Ǽǡ�"
End Select
nl="<p>����" & nl & "</p>"%><html>
<head>
<title>������ʾ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center"> 
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
              <table border="0" bgcolor="#3399FF" cellspacing="0" cellpadding="2" width="350">
                <tr>
<td width="342"><font color="FFFFFF">�������ʾ</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:history.go(-1)" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="�ر�"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr> 
                  <td width="59" align="center" valign="top"><font face="Wingdings" color="#0066FF" style="font-size:32pt">L</font></td>
<td width="269">
<%=nl%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value="��ȷ ����" onclick=javascript:history.go(-1)>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html>
