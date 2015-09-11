<%@ LANGUAGE=VBScript codepage ="936" %>
<%
Dim RefreshTime, IdleTime, TotalUsers, OnlineUser(), Tmp(), Num, I, ID
RefreshTime = 60           '设定网页自动更新时间为10秒       
IdleTime = RefreshTime * 3 '设定闲置时间为自动更新时间的3倍
Application.Lock
'OnlineUser阵列记录了所有连线到此网页之浏览器的SessionID
'清点所有连线到此网页的浏览器, 然后将目前开启之浏览器的SessionID放入阵列的最后
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

'记录目前开启之浏览器的最近存取时间
Application(Session.SessionID & "LastAccessTime") = Timer

'检查所有连线到此网页之浏览器的最近存取时间, 若与目前时间相差30秒以上, 表示结束连线
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

'Num表示目前线上人数, 若与Application("TotalUsers")不同, 表示中间有人断线
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
<title>行走功能</title>
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
        　
        <table border="1" width="100%" cellspacing="3" cellpadding="3" bordercolor="#FBEFD7">
           <tr>
            <td width="100%" align="center" valign="top"><font color="#FFFF00">【行走<img border="0" src="../img/ki17.gif">命令】</font></td>
          </tr>
          <tr>
            <td width="100%" align="center"><font color="#FFFF00">【地图规则】</font></td>
          </tr>
          <tr>
            <td width="100%">
              <table border="0" width="100%">
                <tr>
                  <td width="50%" align="right" bgcolor="#008080"><a href="javascript:s('/移动$ 北上')" title="往上北移动：移动轻功扣100!"><font color="#FFFFFF">北上</font></a></td> 

                  <td width="50%" bgcolor="#008080"><a href="javascript:s('/移动$ 下南')" title="往下南移动：移动轻功扣100!"><font color="#FFFFFF">下南</font></a></td>
                </tr>
                <tr>
                  <td width="50%" align="right" bgcolor="#008080"><a href="javascript:s('/移动$ 左西')" title="往左西移动：移动轻功扣100"><font color="#FFFFFF">左西</font></a></td>
                  <td width="50%" bgcolor="#008080"><a href="javascript:s('/移动$ 右东')" title="往右东移动：移动一次轻功扣100"><font color="#FFFFFF">右东</font></a></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td width="100%">
              <table border="1" width="100%" bordercolor="#F8AC6D">
                <tr>
                  <td width="33%" align="right" bgcolor="#008080"><a href="javascript:s('/移动$ 西北')" title="往西北移动：移动轻功扣100"><font color="#FFFFFF">西北</font></a></td>
                  <td width="33%" align="center"><a href="javascript:s('/移动$ 北')"><img border="0" src="img/np1.gif"></a></td>
                  <td width="34%" bgcolor="#008080"><a href="javascript:s('/移动$ 东北')" title="往东北移动：移动轻功扣100"><font color="#FFFFFF">东北</font></a></td>
                </tr>
                <tr>
                  <td width="33%" align="right" bgcolor="#008080"></td>
                  <td width="33%" align="center"><a href="javascript:s('/移动$ 西')"><img border="0" src="img/np3.gif"></a><a href="javascript:s('/移动$ 环绕')"><img border="0" src="img/np4.gif"></a><a href="javascript:s('/移动$ 东')"><img border="0" src="img/np5.gif"></a></td>
                  <td width="34%" bgcolor="#008080"></td>
                </tr>
                <tr>
                  <td width="33%" align="right" bgcolor="#008080"><a href="javascript:s('/移动$ 西南')" title="往西南移动：移动轻功扣100"><font color="#FFFFFF">西南</font></a></td>
                  <td width="33%" align="center"><a href="javascript:s('/移动$ 南')"><img border="0" src="img/np2.gif"></a></td>
                  <td width="34%" bgcolor="#008080"><a href="javascript:s('/移动$ 东南')" title="往东南移动：移动轻功扣100"><font color="#FFFFFF">东南</font></a></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td width="100%">
              <table border="0" width="100%">
                <tr>
                  <td width="50%" align="right" bgcolor="#008080"><a href="javascript:s('/移动$ 上')" title="往上移动：移动轻功扣100"><font color="#FFFFFF">直上</font></a></td>
                  <td width="50%" bgcolor="#008080"><a href="javascript:s('/移动$ 下')" title="往下移动：移动轻功扣100"><font color="#FFFFFF">直下</font></a></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td width="100%">
              <table border="0" width="100%">
                <tr>
                  <td width="50%" align="right" bgcolor="#008080"><font color="#FFFFFF">丢弃</font></td>
                  <td width="50%" bgcolor="#008080"><font color="#FFFFFF">查找</font></td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
         
            <td width="100%" align="center"><font color="#FFFF00">使用轻功</font></td>
          </tr>
          <tr>
            <td width="100%" align="center"><a title=查看地图以方便行走 
      onClick="window.open('mp.htm','','scrollbars=yes,resizable=yes,width=630,height=500,menubar=no,top=0,left=50')" 
      href="#" target=_self><font color="#FFFF00">查阅地图</font></a></td>
          </tr>
          <tr>
            <td width="100%" align="center"><font color="#FFFF00">目前有</font><font color="#FFFFFF">【<%= Application("TotalUsers") %>】</font><font color="#FFFF00">位行走</font>
</td>
          </tr>   </table>
      </td>
    </tr>
  </table>
  </center>
</div></body></html>