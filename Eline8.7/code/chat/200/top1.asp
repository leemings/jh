<%
RefreshTime =600
' session("tclock")=empty

 if  session("tclock")= ""  then session("tclock")=1
 session("tclock")= session("tclock")+1
'response.write  session("tclock")

%>

<html>
<head>
<title>���ֽ���--��Ϸ����-happyjh.com</title>
   <META HTTP-EQUIV="Refresh" CONTENT="<%= RefreshTime %>, URL=top1.asp">
<style type="text/css">
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</head>

<body bgcolor="#FFFFFF" background="gif/bigbg.gif">
 <%'bgsound src="midi2.mid" loop="20"%>


<%
for i= 1 to 14
if session("players")=i then
if i=1 then regname="С����"
if i=2 then regname="��������"
if i=3 then regname="Ǯ����"
if i=4 then regname="ɳ¡��˹"
if i=5 then regname="��С��"
if i=6 then regname=" ɳ������"
if i=7 then regname="����"
if i=8 then regname="����"
if i=9 then regname="�𱴱�"
if i=10 then regname="��̫��"
if i=11 then regname="Լ����"
if i=12 then regname="���ز�"
session("regname")=regname


 %>
<table width="100%" border="0" cellspacing="0" bgcolor="#FF9830" cellpadding="1">
  <tr> 
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td rowspan="2" width="8%"><img src="<%="gif/big"&trim(cstr(session("players")))&".gif"%>" width="71" height="61" border="0"></td>
          <td rowspan="2" height="41" background="gif/bigbg.gif"> 
            <div align="center"><%=session("regname")%></div>
          </td>
        </tr>
        <tr> </tr>
      </table>
    </td>
  </tr>
</table>
<%end if
next
 %> 
</body>
</html>
