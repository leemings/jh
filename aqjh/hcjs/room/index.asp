<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="roomconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=210"
if sfkf=0 then
	response.write "���ɷ��䴴�������ѹرա�"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,���,grade from �û� where ����='"&aqjh_name&"'",conn,2,2
mymp=rs("����")
mysf=rs("���")
rs.close
if mysf<>"����" or aqjh_grade<5 or mymp="����" or mymp="����" or mymp="����" then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�㲻�����ţ���Ҫ���ܣ�"
	response.end
end if
rs.open "select count(����) as ���� from �û� where ����='"&mymp&"' and �ȼ�>="&dzdj,conn
zs=rs("����")
rs.close
rs.open "select * from r where a='"&mymp&"'",conn,1,2
if (rs.eof or rs.bof) and zs<dzrs then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�����ɵ�����"&dzdj&"�������߲���"&dzrs&"�ˣ����ܴ������䣡"
	response.end
end if
if not(rs.eof or rs.bof) then
	rooma=rs("a")	'������'
	roomid=rs("id")		'����ID
	roomb=rs("b")		'������߶�����
	roomc=rs("c")		'�Ƿ������ƣ�0Ϊ�����ƣ�1Ϊ������
	roomd=rs("d")		'���뷿������˵��
	roome=rs("e")		'���뷿������
	roomf=rs("f")		'�������Ƿ������ܣ�0Ϊ����1Ϊ��ֹ
	roomg=rs("g")		'�������Ƿ������������¼���0Ϊ����1Ϊ��ֹ
	roomh=rs("h")		'�������Ƿ���������0Ϊ����1Ϊ��ֹ
	roomi=rs("i")		'�������Ƿ�����ʹ�ÿ�Ƭ��0Ϊ����1Ϊ��ֹ
	roomj=rs("j")		'�������Ƿ�����Ĳ���0Ϊ����1Ϊ��ֹ
	tjasp="modiroomok.asp"
	an="�� ��"
	s="����"&mymp&" "&mysf&"�����ɵ�����<font color=blue>"&dzdj&"</font>�����ϵ��˹���"&zs&"��"
	if zs<100 then
    		s1="��<font color=red><b����"&dzrs&"��</b></font>�������µ�ʱ����"&dzdj&"�����ϵ���������"&dzrs&"���ϣ�վ����ɾ�������ɡ�"
	else
		s1="��"
    	end if
	s=s&s1
else
	rooma=mymp	'������'
	roomid="�������ɷ���"		'����ID
	roomb=200		'������߶�����
	roomc=0		'�Ƿ������ƣ�0Ϊ�����ƣ�1Ϊ������
	roomd="�����κ��˽���"		'���뷿������˵��
	roome=""		'���뷿������
	roomf=0		'�������Ƿ������ܣ�0Ϊ����1Ϊ��ֹ
	roomg=0		'�������Ƿ������������¼���0Ϊ����1Ϊ��ֹ
	roomh=0		'�������Ƿ���������0Ϊ����1Ϊ��ֹ
	roomi=0		'�������Ƿ�����ʹ�ÿ�Ƭ��0Ϊ����1Ϊ��ֹ
	roomj=0		'�������Ƿ�����Ĳ���0Ϊ����1Ϊ��ֹ
	tjasp="addroomok.asp"
	an="�� ��"
	s="����<font color=blue>"&mymp&"</font> <font color=red>"&mysf&"</font>�������ɵ�����"&dzdj&"�����ϵ��Ѵﵽ"&dzrs&"�����ϣ���������Դ���һ�����ɵķ��䡣"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>���ɷ��䴴������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#FFFFFF" background="../../JHIMG/bk_Hc3w.gif">
<div align="center">
  <p><font size="5"><b><font color="#0000FF">�������ɾ�����</font></b></font></p>
  <p><font size="2" color="#000000"><%=s%></font></p>
</div>
<form name="form1" method="post" action="<%=tjasp%>">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="82%">
    <tr> 
      <td width="102"> 
        <div align="center"><font color="#FFFFFF" size="2">ID</font></div>
      </td>
      <td width="75"><font color="#FFFFFF" size="2"> <%=roomid%> </font> 
        <input type="hidden" name="id" value="<%=roomid%>" size="15" maxlength="20">
      </td>
      <td width="81"> <font size="2" color="#FFFFFF">�� �� ��</font></td>
      <td width="353"><font color="#FFFFFF" size="2"> 
        <input type="text" name="name" readonly value="<%=rooma%>" size="15" maxlength="20">
        ������Ϊ�����������ɸ���</font> </td>
    </tr>
    <tr> 
      <td width="102"> 
        <div align="center"><font color="#FFFFFF" size="2">�Ƿ���Ա���</font></div>
      </td>
      <td width="75"><font color="#FFFFFF" size="2"> 
        <select name=bh size=1 >
          <option value='0' <%if roomh=0 then%>selected<%end if%>><font color="#FFFFFF" size="2">����</font></option>
          <option value='1' <%if roomh=1 then%>selected<%end if%>><font color="#FFFFFF" size="2">��ֹ</font></option>
        </select>
        </font></td>
      <td width="81" nowrap> <font color="#FFFFFF" size="2">�Ƿ����ɱ��</font></td>
      <td width="353"><font color="#FFFFFF" size="2"> 
        <select name=fzxz size=1 >
          <option value='0' <%if roomf=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if roomf=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="102" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">�Ƿ�������¼�</font></div>
      </td>
      <td width="75"><font color="#FFFFFF" size="2"> 
        <select name=sjxz size=1 >
          <option value='0' <%if roomg=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if roomg=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </font></td>
      <td width="81"> <font color="#FFFFFF" size="2">�Ƿ�����ÿ�</font></td>
      <td width="353"><font color="#FFFFFF" size="2"> 
        <select name=kp size=1 >
          <option value='0' <%if roomi=0 then%>selected<%end if%>><font color="#FFFFFF" size="2">����</font></option>
          <option value='1' <%if roomi=1 then%>selected<%end if%>><font color="#FFFFFF" size="2">��ֹ</font></option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="102"> 
        <div align="center"><font size="2" color="#FFFFFF">�Ƿ���ԶĲ�</font></div>
      </td>
      <td width="75"><font color="#FFFFFF" size="2"> 
        <select name=db size=1 >
          <option value='0' <%if roomj=0 then%>selected<%end if%>><font color="#FFFFFF" size="2">����</font></option>
          <option value='1' <%if roomj=1 then%>selected<%end if%>><font color="#FFFFFF" size="2">��ֹ</font></option>
        </select>
        </font></td>
      <td width="81"> 
        <div align="center"><font color="#FFFFFF" size="2">�Ƿ�ֻ�б��ŵ��Ӳſɽ���</font></div>
      </td>
      <td width="353"><font color="#FFFFFF" size="2"> 
        <select name=xz size=1 >
          <option value='0' <%if roomc=0 then%>selected<%end if%>><font color="#FFFFFF" size="2">���������˽���</font></option>
          <option value='1' <%if roomc=1 then%>selected<%end if%>><font color="#FFFFFF" size="2">ֻ�����Ž���</font></option>
        </select>
        �����ƶԹٸ���Ա��Ч </font></td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center">�������ɾ�������Ҫ����������<font color="#FF0000"><%=yinliang%></font>����ң�<font color="#FF0000"><%=jinbi%></font>����Ա�𿨣�<font color="#FF0000"><%=jinka%></font><br>
          <input type="submit" value="<%=an%>" name="submit">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.close()" value="�� ��" type=button name="close">
        </div>
      </td>
    </tr>
  </table>
</form>
</body>
</html>