<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="lib/pub_MainStyle.css">
<%
	Dim theInstalledObjects(19)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
	theInstalledObjects(18) = "MSXML2.XMLHTTP"
	theInstalledObjects(19) = "AspHTTP.Conn"

	call main()

sub main()
%>
<table cellpadding=3 cellspacing=1 border=0 width=80%>
   <tr>
      <td width="80%" valign=top><%call servervar()%></td>
   </tr>
</table>
<%
end sub

sub servervar()
%>
              <table width="100%" border="0" cellspacing="3" cellpadding="0">
                <tr> 
                  <td colspan="2"  height=20> 
                    <b>�������йصı���</b>
                  </td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>��ʾ�ͻ�����������HTTP����</td>
                  <td width="70%"><%=request.ServerVariables("All_Http")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>��ȡISAPIDLL��metabase·��</td>
                  <td width="70%"><%=request.ServerVariables("APPL_MD_PATH")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>��ʾվ������·��</td>
                  <td width="70%"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>·����Ϣ</td>
                  <td width="70%"><%=request.ServerVariables("PATH_INFO")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>��ʾ�������IP��ַ</td>
                  <td width="70%"><%=request.ServerVariables("REMOTE_ADDR")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>������IP��ַ</td>
                  <td width="70%"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>��ʾִ��SCRIPT������·��</td>
                  <td width="70%"><%=request.ServerVariables("SCRIPT_NAME")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>���ط���������������DNS��������IP��ַ</td>
                  <td width="70%"><%=request.ServerVariables("SERVER_NAME")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>���ط�������������Ķ˿�</td>
                  <td width="70%"><%=request.ServerVariables("SERVER_PORT")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>Э������ƺͰ汾</td>
                  <td width="70%"><%=request.ServerVariables("SERVER_PROTOCOL")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>�����������ƺͰ汾</td>
                  <td width="70%"><%=request.ServerVariables("SERVER_SOFTWARE")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>����������ϵͳ</td>
                  <td width="70%"><%=Request.ServerVariables("OS")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>�ű���ʱʱ��</td>
                  <td width="70%"><%=Server.ScriptTimeout%> ��</td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>������CPU����</td>
                  <td width="70%"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>��������������</td>
                  <td width="70%"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
                </tr>
                <tr> 
                  <td colspan="2"  height=20> 
                    <b>���֧�����</b>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
<%
Dim strClass
    strClass = Trim(Request.Form("classname"))
    If "" <> strClass then
    Response.Write "<br>��ָ��������ļ������"
      If Not IsObjInstalled(strClass) then 
        Response.Write "<br><font color=red>���ź����÷�������֧��" & strclass & "�����</font>"
      Else
        Response.Write "<br><font color=green>��ϲ���÷�����֧��" & strclass & "�����</font>"
      End If
      Response.Write "<br>"
    end if
%>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
                    <b>����IIS�Դ����</b>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
	<tr height=22 align=center>
	<td width="70%">�� �� �� ��</td><td width="15%">֧ ��</td><td width="15%">��֧��</td>
	</tr>
    <%
    dim i
    For i=0 to 10
      Response.Write "<TR align=center height=22><TD align=left>&nbsp;" & theInstalledObjects(i) & "<font color=#888888>&nbsp;"
	  select case i
		case 9
		Response.Write "(FSO �ı��ļ���д)"
		case 10
		Response.Write "(ACCESS ���ݿ�)"
	  end select
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<td></td><td><font color=red><b>��</b></font></td>"
      Else
        Response.Write "<td><b>��</b></td><td></td>"
      End If
      Response.Write "</TR>" & vbCrLf
    Next
    %>
	</table>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
                    <b>���������������</b>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
	<tr height=22 align=center>
	<td width="70%">�� �� �� ��</td><td width="15%">֧ ��</td><td width="15%">��֧��</td>
	</tr>
    <%

    For i=11 to UBound(theInstalledObjects)
      Response.Write "<TR align=center height=18><TD align=left>&nbsp;" & theInstalledObjects(i) & "<font color=#888888>&nbsp;"
	  select case i
		case 11
		Response.Write "(SA-FileUp �ļ��ϴ�)"
		case 12
		Response.Write "(SA-FM �ļ�����)"
		case 13
		Response.Write "(JMail �ʼ�����)"
		case 14
		Response.Write "(CDONTS �ʼ����� SMTP Service)"
		case 15
		Response.Write "(ASPEmail �ʼ�����)"
		case 16
		Response.Write "(LyfUpload �ļ��ϴ�)"
		case 17
		Response.Write "(ASPUpload �ļ��ϴ�)"
		case 18
		Response.Write "(XMLHTTP Զ��Http��ȡ)"
		case 19
		Response.Write "(ASPHTTP Զ��Http��ȡ)"

	  end select
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<td></td><td><font color=red><b>��</b></font></td>"
      Else
        Response.Write "<td><b>��</b></td><td></td>"
      End If
      Response.Write "</TR>" & vbCrLf
    Next
	if IsObjInstalled(theInstalledObjects(18)) then
		Response.Write "<p align=center><font color=red><b>��ϲ,���ķ�����֧���Զ��ɼ�����Ԥ��!</b></font></p>"
	else
		Response.Write "<p align=center><font color=red><b>�Բ���,���ķ�������֧��XMLHTTP,�޷��Զ��ɼ�����Ԥ��!</b></font></p>"
    end if
    %>
	</table>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
                    <b>�����������֧�������⣺</b>��������������������Ҫ���������ProgId��ClassId��
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
<FORM action="Server.asp" method=post id=form1 name=form1>
	  <tr>
		<td align=center height=30>
<input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value="ȷ ��" class=stbtmshort id=submit1 name=submit1>
<INPUT type=reset value="�� ��" id=reset1 class=stbtmshort name=reset1> 
</td>
	  </tr>
</FORM>
	</table>
                  </td>
                </tr>
              </table>
<%
end sub
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>