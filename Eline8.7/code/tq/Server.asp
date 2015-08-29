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
                    <b>服务器有关的变量</b>
                  </td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>显示客户发出的所有HTTP标题</td>
                  <td width="70%"><%=request.ServerVariables("All_Http")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>检取ISAPIDLL的metabase路径</td>
                  <td width="70%"><%=request.ServerVariables("APPL_MD_PATH")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>显示站点物理路径</td>
                  <td width="70%"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>路径信息</td>
                  <td width="70%"><%=request.ServerVariables("PATH_INFO")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>显示请求机器IP地址</td>
                  <td width="70%"><%=request.ServerVariables("REMOTE_ADDR")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>服务器IP地址</td>
                  <td width="70%"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>显示执行SCRIPT的虚拟路径</td>
                  <td width="70%"><%=request.ServerVariables("SCRIPT_NAME")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>返回服务器的主机名，DNS别名，或IP地址</td>
                  <td width="70%"><%=request.ServerVariables("SERVER_NAME")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>返回服务器处理请求的端口</td>
                  <td width="70%"><%=request.ServerVariables("SERVER_PORT")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>协议的名称和版本</td>
                  <td width="70%"><%=request.ServerVariables("SERVER_PROTOCOL")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>服务器的名称和版本</td>
                  <td width="70%"><%=request.ServerVariables("SERVER_SOFTWARE")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>服务器操作系统</td>
                  <td width="70%"><%=Request.ServerVariables("OS")%></td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>脚本超时时间</td>
                  <td width="70%"><%=Server.ScriptTimeout%> 秒</td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>服务器CPU数量</td>
                  <td width="70%"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> 个</td>
                </tr>
                <tr> 
                  <td width="30%" valign=top>服务器解译引擎</td>
                  <td width="70%"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
                </tr>
                <tr> 
                  <td colspan="2"  height=20> 
                    <b>组件支持情况</b>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
<%
Dim strClass
    strClass = Trim(Request.Form("classname"))
    If "" <> strClass then
    Response.Write "<br>您指定的组件的检查结果："
      If Not IsObjInstalled(strClass) then 
        Response.Write "<br><font color=red>很遗憾，该服务器不支持" & strclass & "组件！</font>"
      Else
        Response.Write "<br><font color=green>恭喜！该服务器支持" & strclass & "组件。</font>"
      End If
      Response.Write "<br>"
    end if
%>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
                    <b>－－IIS自带组件</b>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
	<tr height=22 align=center>
	<td width="70%">组 件 名 称</td><td width="15%">支 持</td><td width="15%">不支持</td>
	</tr>
    <%
    dim i
    For i=0 to 10
      Response.Write "<TR align=center height=22><TD align=left>&nbsp;" & theInstalledObjects(i) & "<font color=#888888>&nbsp;"
	  select case i
		case 9
		Response.Write "(FSO 文本文件读写)"
		case 10
		Response.Write "(ACCESS 数据库)"
	  end select
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
      Else
        Response.Write "<td><b>√</b></td><td></td>"
      End If
      Response.Write "</TR>" & vbCrLf
    Next
    %>
	</table>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
                    <b>－－其他常见组件</b>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
	<tr height=22 align=center>
	<td width="70%">组 件 名 称</td><td width="15%">支 持</td><td width="15%">不支持</td>
	</tr>
    <%

    For i=11 to UBound(theInstalledObjects)
      Response.Write "<TR align=center height=18><TD align=left>&nbsp;" & theInstalledObjects(i) & "<font color=#888888>&nbsp;"
	  select case i
		case 11
		Response.Write "(SA-FileUp 文件上传)"
		case 12
		Response.Write "(SA-FM 文件管理)"
		case 13
		Response.Write "(JMail 邮件发送)"
		case 14
		Response.Write "(CDONTS 邮件发送 SMTP Service)"
		case 15
		Response.Write "(ASPEmail 邮件发送)"
		case 16
		Response.Write "(LyfUpload 文件上传)"
		case 17
		Response.Write "(ASPUpload 文件上传)"
		case 18
		Response.Write "(XMLHTTP 远程Http获取)"
		case 19
		Response.Write "(ASPHTTP 远程Http获取)"

	  end select
	  Response.Write "</font></td>"
      If Not IsObjInstalled(theInstalledObjects(i)) Then 
        Response.Write "<td></td><td><font color=red><b>×</b></font></td>"
      Else
        Response.Write "<td><b>√</b></td><td></td>"
      End If
      Response.Write "</TR>" & vbCrLf
    Next
	if IsObjInstalled(theInstalledObjects(18)) then
		Response.Write "<p align=center><font color=red><b>恭喜,您的服务器支持自动采集天气预报!</b></font></p>"
	else
		Response.Write "<p align=center><font color=red><b>对不起,您的服务器不支持XMLHTTP,无法自动采集天气预报!</b></font></p>"
    end if
    %>
	</table>
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
                    <b>－－其他组件支持情况检测：</b>在下面的输入框中输入你要检测的组件的ProgId或ClassId。
                  </td>
                </tr>
                <tr> 
                  <td colspan="2" height=20> 
	<table border=0 width="95%" cellspacing=1 cellpadding=0>
<FORM action="Server.asp" method=post id=form1 name=form1>
	  <tr>
		<td align=center height=30>
<input class=input type=text value="" name="classname" size=40>
<INPUT type=submit value="确 定" class=stbtmshort id=submit1 name=submit1>
<INPUT type=reset value="重 填" id=reset1 class=stbtmshort name=reset1> 
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