<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Dim RefreshTime, IdleTime, TotalUsers, OnlineUser(), Tmp(), Num, I, ID
RefreshTime = 60           '�趨��ҳ�Զ�����ʱ��Ϊ10��       
IdleTime = RefreshTime * 3 '�趨����ʱ��Ϊ�Զ�����ʱ���3��
Application.Lock
'OnlineUser���м�¼���������ߵ�����ҳ֮�������SessionID
'����������ߵ�����ҳ�������, Ȼ��Ŀǰ����֮�������SessionID�������е����
If Application(Session.SessionID & "LastAccessTime") = Empty Then
  If Application("TotalUsers") = Empty Then Application("TotalUsers") = 0
  ReDim Tmp(Application("TotalUsers") + 1)
  Num = 0
  If Application("TotalUsers") > 0 Then
    For I = LBOUND(Application("OnlineUser")) To UBOUND(Application("OnlineUser"))
      ID = Application("OnlineUser")(I)
      If ID <> Session.SessionID Then
        Tmp(Num) = ID
        Num = Num + 1
      End If
    Next
  End If
  Tmp(Num) = Session.SessionID
  Application("TotalUsers") = Num + 1
  ReDim Preserve Tmp(Application("TotalUsers"))
  Application("OnlineUser") = Tmp
End If

'��¼Ŀǰ����֮������������ȡʱ��
Application(Session.SessionID & "LastAccessTime") = Timer

'����������ߵ�����ҳ֮������������ȡʱ��, ����Ŀǰʱ�����30������, ��ʾ��������
ReDim Tmp(Application("TotalUsers"))
Num = 0
For I = 0 To Application("TotalUsers") - 1
  ID = Application("OnlineUser")(I)
  If (Timer - Application(ID & "LastAccessTime")) < IdleTime Then
    Tmp(Num) = ID
    Num = Num + 1
  Else
    Application(ID & "LastAccessTime") = Empty
  End If
Next

'Num��ʾĿǰ��������, ����Application("TotalUsers")��ͬ, ��ʾ�м����˶���
If Num <> Application("TotalUsers") Then
  ReDim Preserve Tmp(Num)
  Application("OnlineUser") = Tmp
  Application("TotalUsers") = Num
End If

Application.UnLock
%>
<html><head>
<meta http-equiv="Content-Language" content="zh-tw">
<META HTTP-EQUIV="Refresh" CONTENT="<%= RefreshTime %>, URL=<%= Request.ServerVariables("PATH_INFO") %>">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>���߹���</title>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}
function rc(list){if(list!="0"){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}}
</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body bgcolor="#000000" text="#FFFFFF" background="../bg.gif">
<div align="center">
  <center>
  <table border="0" width="100%" bordercolor="#996600" height="323" bordercolorlight="#996600" bordercolordark="#996600" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%" height="317" valign="top">
        ��
        <table border="1" width="100%" cellspacing="3" cellpadding="3" bordercolor="#FBEFD7">
           <tr>
            <td width="100%" align="center" valign="top"><font color="#FFFF00">������<img border="0" src="../img/ki17.gif">���</font></td>
          </tr>
          <tr>
            <td width="100%" align="center"><font color="#FFFF00">����ͼ����</font></td>
          </tr>
          <tr>
            <td width="100%">
              <table border="0" width="100%">
                <tr>
                  <td width="50%" align="right" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ����')" title="���ϱ��ƶ����ƶ��Ṧ��100!"><font color="#FFFFFF">����</font></a></td> 

                  <td width="50%" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ����')" title="�������ƶ����ƶ��Ṧ��100!"><font color="#FFFFFF">����</font></a></td>
                </tr>
                <tr>
                  <td width="50%" align="right" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ����')" title="�������ƶ����ƶ��Ṧ��100"><font color="#FFFFFF">����</font></a></td>
                  <td width="50%" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ �Ҷ�')" title="���Ҷ��ƶ����ƶ�һ���Ṧ��100"><font color="#FFFFFF">�Ҷ�</font></a></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td width="100%">
              <table border="1" width="100%" bordercolor="#F8AC6D">
                <tr>
                  <td width="33%" align="right" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ����')" title="�������ƶ����ƶ��Ṧ��100"><font color="#FFFFFF">����</font></a></td>
                  <td width="33%" align="center"><a href="javascript:s('/�ƶ�$ ��')"><img border="0" src="img/np1.gif"></a></td>
                  <td width="34%" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ����')" title="�������ƶ����ƶ��Ṧ��100"><font color="#FFFFFF">����</font></a></td>
                </tr>
                <tr>
                  <td width="33%" align="right" bgcolor="#008080"></td>
                  <td width="33%" align="center"><a href="javascript:s('/�ƶ�$ ��')"><img border="0" src="img/np3.gif"></a><a href="javascript:s('/�ƶ�$ ����')"><img border="0" src="img/np4.gif"></a><a href="javascript:s('/�ƶ�$ ��')"><img border="0" src="img/np5.gif"></a></td>
                  <td width="34%" bgcolor="#008080"></td>
                </tr>
                <tr>
                  <td width="33%" align="right" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ����')" title="�������ƶ����ƶ��Ṧ��100"><font color="#FFFFFF">����</font></a></td>
                  <td width="33%" align="center"><a href="javascript:s('/�ƶ�$ ��')"><img border="0" src="img/np2.gif"></a></td>
                  <td width="34%" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ����')" title="�������ƶ����ƶ��Ṧ��100"><font color="#FFFFFF">����</font></a></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td width="100%">
              <table border="0" width="100%">
                <tr>
                  <td width="50%" align="right" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ��')" title="�����ƶ����ƶ��Ṧ��100"><font color="#FFFFFF">ֱ��</font></a></td>
                  <td width="50%" bgcolor="#008080"><a href="javascript:s('/�ƶ�$ ��')" title="�����ƶ����ƶ��Ṧ��100"><font color="#FFFFFF">ֱ��</font></a></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td width="100%">
              <table border="0" width="100%">
                <tr>
                  <td width="50%" align="right" bgcolor="#008080"><font color="#FFFFFF">����</font></td>
                  <td width="50%" bgcolor="#008080"><font color="#FFFFFF">����</font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
         
            <td width="100%" align="center"><font color="#FFFF00">ʹ���Ṧ</font></td>
          </tr>
          <tr>
            <td width="100%" align="center"><a title=�鿴��ͼ�Է������� 
      onClick="window.open('mp.htm','','scrollbars=yes,resizable=yes,width=630,height=500,menubar=no,top=0,left=50')" 
      href="#" target=_self><font color="#FFFF00">���ĵ�ͼ</font></a></td>
          </tr>
          <tr>
            <td width="100%" align="center"><font color="#FFFF00">Ŀǰ��</font><font color="#FFFFFF">��<%= Application("TotalUsers") %>��</font><font color="#FFFF00">λ����</font>
</td>
          </tr>   </table>
      </td>
    </tr>
  </table>
  </center>
</div></body></html>