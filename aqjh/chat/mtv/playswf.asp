<%@ LANGUAGE=VBScript.Encode%>
<% 
Response.Expires=0
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
chatbgcolor=Session("afa_chatbgcolor")
chatimage=Application("aqjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
id=LCase(trim(request.querystring("id")))
songid=request.querystring("name")
toname=request.querystring("toname")
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End 
end if
if toname<>aqjh_name and toname<>"���" then
	Response.Write "<script language=JavaScript>{alert('������ʲô���˼Ҳ�û�е����㣡');close();}</script>"
	Response.End
end if
Set connt=Server.CreateObject("ADODB.CONNECTION")
connt.open Application("aqjh_usermdb")
sql="select * from swf where id="&songid
Set Rs=connt.Execute(sql)
%>
<html>
<head>
<title>MTV ����</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link rel="stylesheet" href="../readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../<%=chatimage%>" bgproperties="fixed" text="#FFFFFF">
<center>
<h4><b>���ĺ��� <font color=yellow><%=id%></font> �����㲥�ĸ�����<%=rs("����")%></b></h2>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="400" height="300"><param name=movie value="<%=rs("·��")%>"><param name=quality value=high><embed src="<%=rs("·��")%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="1" height="1"></embed></object>
</html>