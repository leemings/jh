<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if session("sjjh_name")="" then 
	Response.Write "<script Language=Javascript>top.location.href='http://www.51eline.com';alert('��ʾ���Բ�������û�е�½������');</script>"
	Response.End
end if
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������Ƚ��뽭�������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT ����,���� FROM �û� WHERE ����='"& session("sjjh_name") &"'",conn,1,3
if rs("����")<50000 or rs("����")<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "����û50000���������Ҳ���1�����������裡"
	response.end
end if
rs("����")=rs("����")-50000
rs("����")=rs("����")-1
rs.update
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<%PageName="Box"%>
<!--#include file="conn.asp"-->
<body leftmargin="0" topmargin="0" onLoad="drop(300)">
<%
if request("id")<>"" then
set rs=server.createobject("adodb.recordset")
id=request("id")
sql="select * from MusicList where id="&id
rs.open sql,conn,1,3
if rs.eof then
errmsg="<li>�Բ��𣡸ø��������ڣ������Ѿ�������Աɾ����</li>"
call error()
Response.End 
else
rs("hits")=rs("hits")+1
rs.Update
id=rs("id")
Wma=rs("Wma")
hits=rs("hits")
Singer=rs("Singer")
MusicName=rs("MusicName")
end if
rs.Close

else
errmsg= "<li>��ѡ�������</li>"
call error()
Response.End 
end if

set rs=nothing
conn.close
set conn=nothing
%><html>
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>һ������ �� wWw.51Eline.COM</title>
<style type="text/css">

 <!--

 td {  font-size: 9pt; color: #000000}

 a:link {  font-size: 9pt; color:#000000; text-decoration: none}

 a:active {  font-size: 9pt; ; text-decoration: underline}

 a:visited {  font-size: 9pt; color: #000000; text-decoration: none}

 a:hover {  font-size: 9pt; color:#000000; text-decoration: underline; background-color: #00FFFF}

 -->

 </style>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0"><BODY bgColor=#000000 leftMargin=0 topMargin=0 oncontextmenu="return false" ondragstart="return false" onselectstart="return false"> 
<table width="300" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><img src="files/play.jpg" width="300" height="13"></td>
  </tr>
  <tr> 
    <td><object id="mPlayer1" width=300 height=250
 classid="CLSID:6BF52A52-394A-11D3-B153-00C04F79FAA6" type=application/x-oleobject standby="Loading Windows Media Player components...">
        <param name="URL" value="Yxwma.asp?id=<%=id%>">
                          <param name="rate" value="1">
                          <param name="balance" value="0">
                          <param name="currentPosition" value="0">
                          <param name="defaultFrame" value="">
                          <param name="playCount" value="100">
                          <param name="autoStart" value="-1">
                          <param name="currentMarker" value="0">
                          <param name="invokeURLs" value="-1">
                          <param name="baseURL" value="">
                          <param name="volume" value="100">
                          <param name="mute" value="0">
                          <param name="uiMode" value="full">
                          <param name="stretchToFit" value="0">
                          <param name="windowlessVideo" value="-1">
                          <param name="enabled" value="-1">
                          <param name="enableContextMenu" value="0">
                          <param name="fullScreen" value="0">
                          <param name="SAMIStyle" value="">
                          <param name="SAMILang" value="">
                          <param name="SAMIFilename" value="">
                          <param name="captioningID" value="">
                          <param name="enableErrorDialogs" value="0">
                          <param name="_cx" value="12435">
                          <param name="_cy" value="1640">
                        </object></td>
  </tr>
  <tr> 
    <td><table width="300" height="32" border="0" cellpadding="2" cellspacing="0">
        <tr>
          <td width="180"><marquee>
                        <font color="#FFFFFF"><b>��վ����ȫΪWMA��ʽ���粻���������ţ��������������ų��򣬷�������</b></font></marquee></td>
          <td width="120"><a href="http://windowsmedia.com/download" target="_blank"><img src="images/WMP9series_Free.gif" alt="�����������վ��" width="120" height="32" border="0"></a></td>
        </tr>
      </table></td>
  </tr>
</table>
</BODY></HTML>
<script>
      function rf()
      {return false; }
    document.oncontextmenu = rf
    function keydown()
      {if(event.ctrlKey ==true || event.keyCode ==93 || event.shiftKey ==true){return false;} }
      document.onkeydown =keydown
    function drag()
      {return false;}
    document.ondragstart=drag 
function stopmouse(e) { 
		if (navigator.appName == 'Netscape' && (e.which == 3 || e.which == 2))  
 		return false; 
      else if  
      (navigator.appName == 'Microsoft Internet Explorer' && (event.button == 2 || event.button == 3)) {  
 		alert("�벻Ҫ���ñ�վ���ӣ����IP�Ѿ�����¼��");
		return false;  
 		}
		return true; 
      } 
      document.onmousedown=stopmouse;  
      if (document.layers) 
      window.captureEvents(Event.MOUSEDOWN); 
       window.onmousedown=stopmouse;  
      </script>
<script language="javascript">kstatus();function kstatus(){self.status="һ������� | wWw.51Eline.com | ���Ժ�����ʹ������������ʱ�վ.";setTimeout("kstatus()",0);}</script>