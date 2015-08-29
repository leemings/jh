<%
Dim RefreshTime, IdleTime, TotalUsers, OnlineUser(), Tmp(), Num, I
RefreshTime =10         
IdleTime = RefreshTime * 3   
'NewUser = Request.cookies("regbname")  
'NewUser = session("username")  
NewUser=session("zpid")
UserMsg = NewUser & "Msg"    
Application.Lock
'OnlineUser?数组记录了所有连线到此网页的使用者名称
'清点所有连线到此网页的使用者，然后将目前的使用者名称放入数组的最后面
If Application(NewUser & "LastAccessTime") = Empty Then
  If Application("TotalUsers") = Empty Then Application("TotalUsers") = 0
  ReDim Tmp(Application("TotalUsers") + 1)
  Num = 0
  If Application("TotalUsers") > 0 Then
    For I = LBOUND(Application("OnlineUser")) To UBOUND(Application("OnlineUser"))
      User = Application("OnlineUser")(I)
      If User <> NewUser AND User <> Session("rName") Then
        Tmp(Num) = User
        Num = Num + 1
      Else
        Application(User & "LastAccessTime") = Empty
      End If
    Next
  End If
  Session("rName") = NewUser
  Tmp(Num) = Session("rName") 

  Application("TotalUsers") = Num + 1
  ReDim Preserve Tmp(Application("TotalUsers"))
  Application("OnlineUser") = Tmp
End If

'记录目前使用者的最近存取时间
Application(Session("rName") & "LastAccessTime") = Timer

'检查所有连线到此网页的使用者的最近存取时间，若与目前时间相差30秒以上，表示离线
ReDim Tmp(Application("TotalUsers"))
Num = 0
For I = 0 To Application("TotalUsers") - 1
  User = Application("OnlineUser")(I)
  If (Timer - Application(User & "LastAccessTime")) < IdleTime Then
    Tmp(Num) = User
    Num = Num + 1
  Else
    Application(User & "LastAccessTime") = Empty
    Application(UserMsg) = Empty
  End If
Next

'Num表示目前在线人数，若与Application("TotalUsers")不同，表示中间有人离线
If Num <> Application("TotalUsers") Then
  ReDim Preserve Tmp(Num)
  Application("OnlineUser") = Tmp
  Application("TotalUsers") = Num
End If
Application.UnLock
%>
<%   
      '检查是否有人传送信息给目前的使用者，有的话就调用ShowMsg.asp显示信息  
      If Application(UserMsg) <> "" then    
       Application("ShowMsg") = Application(UserMsg)          
    %> <%  
      '根据使用者名称及数字来决定显示信息的窗口的ID, ID若相同，信息窗口将被覆盖掉  
      Session("WinID") = Session("WinID") + 1  
      Session("UserWinID") = NewUser & Session("WinID")  
    %> 
<SCRIPT LANGUAGE="VBScript">          
      Window.open "ShowMsg.asp", "<%= Session("UserWinID") %>", "Menubar=0, Width=400, Height=180"           
    </SCRIPT>  
    <%  
      '信息显示完毕之后就将存放信息的变量值恢复为Empty   
      Application(UserMsg) = Empty   
      End If   
    %>     

<html>
<head>
    <META HTTP-EQUIV="Refresh" CONTENT="<%= RefreshTime %>, URL=<%= Request.ServerVariables("PATH_INFO") %>?rName=<%= Request("rName") %>">

<title>...在线通讯...</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style1.css" rel=stylesheet>

</head>

<body bgcolor="#CCCCFF">
<div align="center">
  <table width="100%" border="0" cellspacing="0" height="168">

    <tr>


      <td width="10%" height="335">&nbsp;</td>
      <td width="80%" bgcolor="#FFFFFF" valign="top" height="335"> 
        <div align="center">
<%
ii=1
 For I = 0 To (Application("TotalUsers") - 1) 
'ii=i+1
%> 
<img src="<%=ii%>.gif"   body="0"> <font color="#6666FF"><%= Application("OnLineUser")(I) %></font>
<% 
response.write "<br>"
Next %> 
</div>
      </td>

      <td width="10%" height="335"> 
        <div align="center"></div>
      </td>

    </tr>


  </table>
  <font color="#6666FF"> 
  </font></div>

        
</body>
</html>
