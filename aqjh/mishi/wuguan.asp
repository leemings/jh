<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const.asp"-->
<!--#include file="../config.asp"-->
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("axjh_usermdb")
rs.open "select ����,����,�书,�书��,�ȼ�,��Ա�ȼ�,��Ա from �û� where ����='"&axjh_name&"'",conn
hydj=rs("��Ա�ȼ�")
pdhy=rs("��Ա")
mytl=rs("����")
mynl=rs("����")
mywg=rs("�书")
wgsx=int(rs("�ȼ�")*axjh_wgsx+3800+rs("�书��"))
rs.close
set rs=nothing
conn.close
set conn=nothing
select case hydj
	case 0
		bf=1.2
	case 1
		bf=1.25
	case 2
		bf=1.3
	case 3
		bf=1.35
	case 4
		bf=1.4
end select
if pdhy<>false then
	bf=1.3
end if
maxwg=int(wgsx*bf)
wgc=maxwg-mywg
if wgc>0 and wgc=<1000 then
	xlwg="�������"
	xtl=10000
	xnl=5000
elseif wgc>1000 and wgc<=10000 then
	xlwg="����Ӳ��"
	xtl=100000
	xnl=50000
elseif wgc>10000 and wgc<=100000 then
	xlwg="��ü�ķ�"
	xtl=1000000
	xnl=500000
elseif wgc>100000 and wgc<=150000 then
	xlwg="��ң��"
	xtl=1500000
	xnl=750000
elseif wgc>150000 and wgc<=200000 then
	xlwg="��ҹƶ�"
	xtl=2000000
	xnl=1000000
elseif wgc>200000 then
	xlwg="��Ȼ����"
	xtl=10000000
	xnl=5000000
elseif wgc<=0 then
	xlwg="��������"
	xtl=0
	xnl=0
end if
if xlwg="��������" then
	mess="��Ŀǰ���书�Ѵ���߾��磬�����������������������ɡ�"
else
	mess="������"&xlwg&"��������"
	if xtl-mytl>0 then
		mess=mess&"<font color=red>"&xtl&"��</font>��������Ŀǰ���������㡣"
	else
		mess=mess&"<font color=red>"&xtl&"��</font>������"
	end if
	if xnl-mynl>0 then
		mess=mess&"<font color=red>"&xnl&"��</font>��������Ŀǰ���������㡣"
	else
		mess=mess&"<font color=red>"&xnl&"��</font>������"
	end if
end if
%>
<html>

<head>
<link rel="stylesheet" href="css.css">
<title><%=application("axjh_chatroomname")%></title>
<style type="text/css">
<!--
.style9 {font-size: 12px}
-->
</style>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../jhimg/bgcheetah.gif">
<br>
<br>
<table border="0" cellspacing="0" cellpadding="0" width="366" align="center">
  <tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=axjh_name%></font>��ӭ����ϰ�䳡</b><br>
        <br>
        �������㵱ǰ��״̬<br>
        ������<font color="#FF0000"><%=mytl%></font>��������<font color="#FF0000"><%=mynl%></font><br>
        �书��<font color="#FF0000"><%=mywg%></font>���书���ޣ�<font color="#FF0000"><%=wgsx%></font></font><br>
        �书��߿���������<font color="#008000"><%=maxwg%></font>�㡣����<font color="#008040"><%=maxwg-mywg%></font>�㡣 
        <%if xlwg<>"��������" then%>
        <br>
        ����������<font color="#FF0000"><%=xlwg%></font> 
        <%end if%>
        <br>
		<%=mess%></div>
<form method=POST action='wuguanok.asp'>
        <table width="414" align="center">
          <tr>
<td>
<tr>
<td align=center>
<select name=money size=1> 
<option value="1000" <%if xlwg="�������" or xlwg<>"��������" then%>selected<%end if%>>�������</option>
<option value="10000" <%if xlwg="����Ӳ��" then%>selected<%end if%>>����Ӳ��</option>
<option value="100000" <%if xlwg="�����ķ�" then%>selected<%end if%>>��ü�ķ�</option>
<option value="150000" <%if xlwg="��ң��" then%>selected<%end if%>>��ң��</option>
<option value="200000" <%if xlwg="��ҹƶ�" then%>selected<%end if%>>��ҹƶ�</option>
<option value="1000000" <%if xlwg="��Ȼ����" then%>selected<%end if%>>��Ȼ����</option>
</select>
</td>
</tr>
<tr>
<td align=center><input type=submit value=��ϰ�书 name="submit"></td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
�������</div>
</td>
</tr>
<tr>
            <td valign="top" > 
              <p><br>
                <font color="#660000">�������������䲻��Ҫ��ʦ���������ĵ�Ҳ��������������������������ÿ����1���书������<b><font color="#FF0000">10��</font></b>������<b><font color="#FF0000">5��</font></b>������</font><br>
                (1)��������� ���书��1000�� &nbsp(2)������Ӳ�� ���书��10000��<br> 
                (3)�������ķ� ���书��100000�� (4)����ң�� ���书��150000��<br> 
                (5)����ҹƶ� ���书��200000�� (6)����Ȼ���� ���书��1000000��<br> 
                <br>
              <p><%if hydj=0 and pdhy=false then%>
����������书�����޵�<%=bf%>����
<%else
if pdhy<>false then%>
�����ݵ��ƻ�Ա�����������书�������޵�<%=bf%>����
<%else%>
����<%=hydj%>����Ա�����������书�������޵�<%=bf%>����
<%end if
end if%></p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<div align="center">
</div>
</body>
</html>
