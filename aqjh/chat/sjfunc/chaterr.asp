<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
id=Trim(Request.QueryString("id"))
if id="001" or id="002" then
	if aqjh_name<>"" then
	to1=aqjh_name
	mycd=DateDiff("n",Session("aqjh_savetime"),now())
	if mycd>50 then mycd=50
	if mycd<1 then mycd=1
	Application.Lock
	Application("aqjh_diuqi")="v|"&to1&"|"&mycd&"|"&now()
	Application.UnLock
	diuqi=aqjh_name&"���������ң���δ�������ܼ�:<a href='dq.asp' target='d'><b><font color=red>+"&mycd&"</font></b></a>��,˭������˭�ġ�"
	towhoway=0
	saycolor="660099"
	addwordcolor="660099"
	addsays="��"
	says1="<b>��������Ϣ��</b>"&diuqi
	says1=replace(says1,"'","\'")
	says1=replace(says1,chr(34),"\"&chr(34))
	saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & "�߳�" & chr(39) &","& chr(39) & to1 & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & to1 & chr(39) &"," & chr(39) & says1 & chr(39) &",0," & session("nowinroom") & ");</script>"
	addmsg saystr
	Function Yushu(a)
		Yushu=(a and 31)
	End Function
	Sub AddMsg(Str)
		Application.Lock()
		Application("SayCount")=Application("SayCount")+1
		i="SayStr"&YuShu(Application("SayCount"))
		Application(i)=Str
		Application.UnLock()
	End Sub
	end if
end if
Session.Abandon
Select Case id
Case "000"
	nl="�����Բ��𣬳�������Ŀ¼��������Ŀ¼���������д���Global.asa û�б�ִ�С���������Ҫ����Ŀ¼��֧�֣�"
Case "001"
	nl="�����ڼ������رա�<br><br>��ɴ������ԭ����Ҫ�У�<br>�������������紫�����⣬������Ļ���������������������޷��������ݣ�ϵͳ������Ϊ��ʱ���Ͽ������ӣ�<br>�����������ˡ�������¼���ػص�¼ҳ�棬��û���ȹرմ˴��ڣ�<br>�������㱻�߳��������ҡ�<br>�����������ʹ�ñ����ڽ��벻�����������⣬�����´��ڽ���ͳ����������Ļ�������ͳ�����Ļ������Ե������´����޷��̳��ϼ����ڵ�ֵ��<br><br>���������<br>�����رմ˴��ڣ����µ���¼ҳ�������û�����������е�¼��������Ǳ��߳������ң����㷽�����õ��û����� 5 �����ڲ���ʹ�á�����ǵڣ��������������������ɣŵ���ʱ�ļ���Ȼ����������������"
Case "002"
	nl="�����ڼ������رա�<br><br>ԭ��<br>�����㳬�� " & Application("aqjh_maxtimeout") & " ����û�з��ԣ�Ϊ���������������ϵͳ�Զ��������ռ�õ���Դ��<br><br>���������<br>�����رմ˴��ڣ����µ���¼ҳ�������û�����������е�¼��"
Case else
	nl="�����Բ��𣬸ó�������δ���Ǽǡ�"
End Select
nl="<p>" & nl & "</p>"%><html>
<head>
<title>������ʾ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%>>
<script LANGUAGE="JavaScript">if(window!=window.top)top.location.href=location.href;</script>
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
<table border="0" bgcolor="#FF0099" cellspacing="0" cellpadding="2" width="350">
<tr>
<td width="342"><font color="FFFFFF">�������ʾ</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="�ر�"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr>
<td width="59" align="center" valign="top"><font face="Wingdings" color="#FF0000" style="font-size:32pt">L</font></td>
<td width="269">
<%=nl%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value="��ȷ ����" onclick=javascript:window.close()>
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