<%@ LANGUAGE=VBScript codepage ="936" %>
<%response.expires=0

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
name=sjjh_name
face=left(trim(request.form("face")),3)
if InStr(face,"or")<>0 or InStr(face,"'")<>0 or InStr(face,"`")<>0 or InStr(face,"=")<>0 or InStr(face,"-")<>0 or InStr(face,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');window.close();}</script>"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("sjjh_usermdb")
conn.Execute "update �û� set ����ͷ��='"&face&"' where ����='" & sjjh_name &"'"
conn.close
set conn=nothing
%><html>
<head>
<title>����ͷ���wWw.51eline.com��</title>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type='text/css'>
body{
font-size:9pt;
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff;
}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#006699" bgproperties="fixed" leftmargin="20" topmargin="30">
<p align="center"><font size="+1" color="#FFFFFF"><img src="../ico/<%=face%>-2.gif"><br><br>
  <font size="3" color="#FFFFFF">���Ѿ��ɹ��޸����Լ���ͷ��</font><font size="3">���˳����������½������Ϳ��Կ������ĺ��ͷ���ˣ�</font></font></p>
<p><br>
</p>
</body>
</html>