<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ��Ա�ȼ� from �û� where ����='"&aqjh_name&"'",conn,2,2
hy=rs("��Ա�ȼ�")
if hy<2 then
ss=0-hy
	Response.Write "<script Language=Javascript>alert('���������Ա�ȼ�����2����Ҫ�ڽ�ң��͸Ͽ������ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ�� from �û� where ����='"&aqjh_name&"'",conn,2,2
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<15 then
ss=15-sj
	Response.Write "<script Language=Javascript>alert('�����"&ss&"�������ڱ����ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>-��Ա�ڱ�-</title>
<meta http-equiv="refresh" content="10">
<style></style>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
</head>
<BODY oncontextmenu=self.event.returnValue=false BGCOLOR="#000000" text="#FFFFFF">
<DIV ID="Layer1" STYLE="position:absolute; left:46px; top:139px; width:192px; height:165px; z-index:3">
<%
if Session("diaoyu")=true then
	Session("diaoyu")=now()
end if 
abc=DateDiff("s",Session("diaoyu"),now())
if abc>=115 then
	Response.Write "<script Language=Javascript>alert('������ǿ�������ˡ���');location.href = 'wano.asp';</script>"
	response.end
end if
if abc<=90 then
%>
<IMG SRC="wa.gif" width="50" height="50"><br>
    <font color="#FFFFFF" size="2">Ŭ���ڰ�,��Ҫ�ڽ��!</font>
<%
else
%>
    <a href="waok.ASP"><IMG SRC="waok.gif" border="0" width="30" height="50"></a> 
    <font color="#0000FF"><br>
    </font> 
    <font color="#FFFFFF" size="2">��,�ڵ��󱦲ذ�,����ѽǿ��Ҫ����!</font>
<%
end if
%>
</div>
</DIV>
<DIV ID="Layer2" STYLE="position: absolute; left: 76px; top: 101px; width: 101px; height: 64px; z-index: 1"> 
  <p align="center"><img border="0" src="wolf-admin.gif"></DIV>  
<span id="liveclock" style="position: absolute; left: 196; top: 139; width: 76; height: 29"> 
<div align="center"><font color="#FFFFFF" size="2">��������: <%=abc%></font> </div>
</span> 
<div id="Layer5" style="position:absolute; width:262px; height:21px; z-index:6; left: 119px; top: 382px"><font color="#FFFFFF" size="2">�ڰ��齭����վ��</font><br>
</div>
</body> 
</html>