<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"


Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select * from �û� where ����='"&sjjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
end if
if rs("�ݶ�����")<1000 then
	mess="�����ڻ�û�ж��ӣ�ֻ��:<font color=red>"&rs("�ݶ�����")&"</font>��"
else
	dousl=int(rs("�ݶ�����")/1000)
	conn.execute "update �û� set �ݶ�����=�ݶ�����-"& dousl*1000 &",��Ա��=��Ա��+"& 2*dousl &" where ����='"&sjjh_name&"'"
	mess="�����ӣ�"&dousl&"������Ա�����ӣ�"&dousl*2&"Ԫ����ʣ:"&rs("�ݶ�����")-dousl*1000&"�㣡"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<link rel="stylesheet" href="../../css.css">
<title>������</title>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bgcheetah.gif">
<table border="0" cellspacing="0" cellpadding="0" width="259" align="center">
  <tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=sjjh_name%></font>��ӭ���������г�����</b></font></div>
      <table width="280" align="center">
        <tr> 
            <td> 
          <tr> 
            
          <td valign="top" height="11" > 
            <div align="center">���Ӽ��</div>
            </td>
          </tr>
          <tr> 
            
          <td valign="top" height="150" > 
            <p>�����������ݵ�ʱ��������,���ݵ����㡣ÿ�ݹ�1000�㣬�����ϵͳ�Զ�����һ����������ѡ�񱩶���ɱ����Ϊƽʱ3�������Գ���1Сʱʱ�䡣����㲻��ʹ�ã��������������ˡ�һ������=2Ԫ��Ա���ã�������������ӵĻ�Ա�ѹ���ֻ�л�Ա���еĿ�Ƭ��</p>
<%=mess%><br><br><br>
            </td>
          </tr>
        </table>

</td>
</tr>
</table>
<div align="center"><font color="#00FF66"><b><font color="#0000FF"><%=Application("sjjh_chatroomname")%></font></b></font>
</div>
</body>
</html>



