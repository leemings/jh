<html>
<head>
<title>E�߽���--��Ϸ����-51eline.com</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="forum.css">
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

<body bgcolor="#FFFFFF" background="gif/msgbg.gif">

<div align="center">
  <table width="100%" border="0" cellspacing="0" align="center">
    <tr> 
      <td> 
        <div align="left"><a href="nan.asp">����</a></div>
      </td>
      <td> 
        <div align="left"><a href="hermsg.asp">�鿴</a></div>
      </td>

    </tr>
  </table>
<%for i =1 to 4 
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
%>
  <table width="100%" border="0" cellspacing="0" align="center">
    <tr> 
      <td width="20%" height="51"> 
        <div align="center"> 
          <table width="100%" border="0" cellspacing="0">
            <tr> 
              <td width="21%"><img src="gif/big<%=i%>.gif" width="52" height="45"></td>
            </tr>
          </table>
        </div>
      </td>
      <td width="80%" height="51"> 
        <div align="center"> 
          <table width="100%" border="0" cellspacing="0">
            <tr> 
              <td width="20%">��</td>
              <td width="20%"><img src="gif/house<%=i%>1.gif" width="27" height="39"></td>
              <td width="20%"><img src="gif/house<%=i%>2.gif" width="28" height="40"></td>
              <td width="20%"><img src="gif/house<%=i%>3.gif" width="28" height="42"></td>
              <td width="20%"><img src="gif/house<%=i%>4.gif" width="30" height="44"></td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
    <tr> 
      <td width="20%"><font color="#009933"><b><%=regname%></b> </font></td>
      <td rowspan="3">
        <table width="100%" border="0" cellspacing="0">
          <tr> 
            <td width="20%">��</td>
            <td colspan="4">
<%
for nn = 1 to 39
whohouse=Trim(application("house"&nn))
whohouse=mid(whohouse,1,1)
houseTop=mid(trim(application("house"&nn)),2,1)
if whohouse=cstr(i)  then 
nn=cstr(nn)
response.write "����"&nn&"~"&housetop&"��  <br>"	
end if
next

%>
</td>
          </tr>
        </table>
      </td>
    </tr>
 

    <tr> 
      <td width="20%"><%=formatcurrency(application("money"&i))%></td>
    </tr>


    <tr> 
      <td width="20%">��</td>
    </tr>
    <tr> 
      <td width="20%">��</td>
      <td width="80%">��</td>
    </tr>
  </table>
<%next%>
</div>
<p align="center"><span lang="zh-cn">��E�߽�����</span></p></body>
</html>