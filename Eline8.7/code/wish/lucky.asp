<%@ LANGUAGE=VBScript codepage ="936" %><%

' --- �����˳������趨
luck1 = "<img src=img/01.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFFF>��</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>�ƺ���������������ʢ!<br>��ʲô�������±�ȥ���԰�,<br>������гɹ�!<br>�ֻ��߸Ͽ�ȥ�����Ʊ��!"
luck2 = "<img src=img/02.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFFF>�м�</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>��ϲ��ѽ!<br>����Ӧ����������,<br>��������Ҳ��ת����!<br>�������¶��ǽ�̤ʵ�غ���,<br>��ֻ����С����!"
luck3 = "<img src=img/03.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFF>С��</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>����Ҳ�㲻���˰�!<br>��Ȼ��ʱ������ʧ��,<br>�����������!"
luck4 = "<img src=img/04.gif border=0></td></tr><tr><td align=center height=15 bgcolor#c00000><font color=#FFFFF>С��</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>Ҳ����첻���˸ɴ���,<br>˯�����ͺ���!"
luck5 = "<img src=img/05.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c0000><font color=#FFFFFF>����</font></td></tr><tr><td bgcolor#FFdbdb align=center valign=top height=130>�޲�Ҫ�ں���Щ��������!<br>������뻵�����������,<br>�����ѽ!"


randomize
x = int(rnd * 100)
' ȷ�� 20%
if x < 20 then 
 luck = luck1
' ȷ�� 20%
elseif x < 40 then
 luck = luck2
' ȷ�� 30%
elseif x < 70 then
 luck = luck3
' ȷ�� 20%
elseif x < 90 then
 luck = luck4
' ȷ�� 10%
else luck = luck5
end if
%>
<html>

<head>
<meta http-equiv="Content-Type"
content="text/html; charset=gb2312">
<title>�����˳̡�wWw.happyjh.com��</title>
<STYLE TYPE="text/css">
<!--
tr, td,body,th    {font-size: 9pt}
a:link    {font-size: 9pt; text-decoration:none; color:<%=link%> }
a:visited {font-size: 9pt; text-decoration:none; color:<%=vlink%> }
a:active  {font-size: 9pt; text-decoration:none; color:<%=alink%> }
a:hover   {font-size: 9pt; text-decoration:underline; color:<%=alink%> }
input,select,textarea     {font-size: 9pt; border: 1 solid black}
-->
</STYLE>
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<div align="left">

<table border="1" cellspacing="1" width="157"
bordercolor="#FFADAD">
    <tr>
        <td height=142 align=center valign=middle><%=luck%>
        </td>
    </tr>
</table>
</div>
</body>
</html>
