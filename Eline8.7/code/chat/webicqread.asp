<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
sjjh_name=Session("sjjh_name")
useronlinename=Application("sjjh_useronlinename" & session("inroom"))
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(","&useronlinename,"," & sjjh_name & ",")=0 then Response.Redirect "chaterr.asp?id=001"
chatroombgimage=Application("sjjh_chatimage")
chatroombgcolor=Application("sjjh_chatbgcolor")
s0=trim(Request.QueryString("s0"))
s1=trim(Request.QueryString("s1"))
s2=trim(Request.QueryString("s2"))
s3=trim(Request.QueryString("s3"))
s4=trim(Request.QueryString("s4"))
s5=trim(Request.QueryString("s5"))
s6=trim(Request.QueryString("s6"))
s7=clng(Request.QueryString("s7"))
s8=clng(Request.QueryString("s8"))
%><html><head><title><%=sjjh_name%>，聊友在呼叫你</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="0">
<table width=100% border=1 align=center cellspacing=0 bordercolorlight=000000 bordercolordark=FFFFFF bgcolor=E0E0E0>
  <tr> 
    <td> <table border=0 bgcolor=009900 cellspacing=0 cellpadding=2 width=100%>
        <tr> 
          <td width=376>Anjh Web 
            ICQ 1.01 - <%=Application("sjjh_name")%></td>
          <td width=18> <table border=1 bordercolorlight=666666 bordercolordark=FFFFFF cellpadding=0 bgcolor=E0E0E0 cellspacing=0 width=18>
              <tr> 
                <td width=16><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="退出"><font color="000000">×</font></a></b></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="E0E0E0" cellpadding="5">
        <tr valign="middle" align="center"> 
          <td class=p9><div align="left">时间：<font color=red><%=s3%></font> (点击姓名即可回复)</div>
            <table border="1" align="center" cellspacing="0" bordercolorlight="#CCCCFF" bordercolordark="#FFFFFF" width="100%">
              <tr align="center" bgcolor="#CCFFCC"> 
                <td width="65">姓名：</td>
                <td width="91" align="left"><a href="#" onClick="javascript:window.open('webicq.asp?towho=<%=s1%>','anjh','Status=no,scrollbars=no,resizable=no,width=380,height=320')"><%=s1%></a></td>
                <td width="609">消　息　内　容</td>
              </tr>
              <tr align="left"> 
                <td colspan="3" class=p9><%=s2%></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<div align="center">
<%if s5="on" then
if len(s4)<30 then
	Response.Write "联接地址：<a href=" & s4 & " target=_blank> " & s4 & "</a><br>"  
else
	Response.Write "联接地址：<a href=" & s4 & " target=_blank> " & left(s4,30) & "……</a><br>"  
end if
ss=right(s4,3)
select case ss
case ".rm","ram"
if s6="on" then
%>
<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=<%=s7%> height=<%=s8%>>
<PARAM NAME=SRC VALUE="<%=s4%>">
<PARAM NAME=CONSOLE VALUE=Clip1>
<PARAM NAME=CONTROLS VALUE=imagewindow>
<PARAM NAME=AUTOSTART VALUE=true></OBJECT><br>
<%end if%>
<OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=<%=s7%>>
<PARAM NAME=SRC VALUE="<%=s4%>">
<PARAM NAME=AUTOSTART VALUE=-1>
<PARAM NAME=CONTROLS VALUE=controlpanel>
<PARAM NAME=CONSOLE VALUE=Clip1>
</OBJECT>
<%case "asf","wma","mid","mp3","wav"
if s6="on" then%>
<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width="<%=s7%>" height="<%=s8%>">
<param name=Filename value="<%=s4%>">
<embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src="<%=s4%>" width="<%=s7%>" height="<%=s8%>">
</embed>
</object>
<%else%>
<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width="200" height="40">
<PARAM NAME=SRC VALUE="<%=s4%>">
<PARAM NAME=AUTOSTART VALUE=-1>
<PARAM NAME=CONTROLS VALUE=controlpanel>
<PARAM NAME=CONSOLE VALUE=Clip1>
</OBJECT>
<%end if%>
<%case "jpg","gif"%>
	<img src="<%=s4%>">
<%case "swf"%>
	<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="<%=s7%>" height="<%=s8%>">
    <param name="movie" value="<%=s4%>">
    <param name="quality" value="high">
    <embed src="<%=s4%>" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="<%=s7%>" height="<%=s8%>"></embed></object>
<%end select%>
<%end if%>
</div>
</body>
</html>

