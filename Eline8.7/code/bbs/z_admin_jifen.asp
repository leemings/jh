<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->

<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣"
		call dvbbs_error()
	else
		call main()
	end if

	sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
  <tr>
    <td>
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr >
        <th>��ӭ<b><%=membername%></b>�����������</th>
        </tr>
            <tr>
          <td width="100%" valign=top>
            <table width="90%" class=tb cellspacing="0" cellpadding="0" align="center" height="22">
              <tr>
                <td align="center" valign="middle">�������� V1.2 ��������</td>
              </tr>
            </table>
            <br>
            <table width="90%" class=tb cellspacing="0" cellpadding="0" align="center" height="22">
              <tr> 
                <td align="left" valign="middle"> 
                  <p> 1��ע����� �����棬�����������������ֵ��趨�������Ҫ�޸ģ�������Ӧ�ĵط�����������ã��Զ���������ǿ��<br>
                    <br>
                    2. ���õ�ֵ��Ҫ̫С��Ҳ��Ҫ̫��̫С��ʹ������������̫�̫࣬�����ӵ�����ֵ�к��٣�����Ϊ�㣻��Ϊ��������ȡ���ģ��������ʱ����<font color="#FF6633">С��</font>���õ������� 
                    ����������㡣</p>
                  <p>3. <font color="#9933CC">ʱ����</font>��ָ����ʱ�����<font color="#FF6600">����</font>�����ʱ��ſ��Ա�����֣��������֮������ʱ�佫���Զ�����Ϊ�㣬���¿�ʼ���㣻���ʱ���������Ϊ��Ҫ��Ҫ���أ�<font color="#9933CC">�Ʋ�������(W)</font>�������������ӵĲƲ�ֵ�����ֵ����Сһ�㣬һ�����<font color="#FF9933">С��</font>ʱ������ֵ��<font color="#9933CC">����������(C)</font>���Ծ������һ�㣬һ��Ϊ���ĸ�ֵ������һ����������������Ϊ�㣬���ӵ�����ֵ������Ϊ1���������ֵ��Ҳ���¡�</p>
                  <p>4. ����ֵ������0��99֮�䣬���򽫻��ȡǰ��λ��Ϊ��ֵ<br>
                    <br>
                    3.<font color="#FF0000"> �ر�ע�⣺</font>���趨��ֵ���Ḳ�ǵ�ԭ�������õ�ֵ����С��ʹ�ã�����</p></td>
              </tr>
            </table>
            <br>
            <form action="z_admin_jifen.asp" method="post">
              <table width="90%" class=tb cellspacing="0" cellpadding="0" align="center" height="249">
                <tr> 
                  <td align="center" valign="middle"> 
                    <%
	dim fsObj,txtObj
	dim jf_StartTime,jf_WealthRate,jf_CPRate,jf_EPRate     '���ּ��ʱ�䣬����Ʋ����ʣ������������ʣ����㾭�����
	dim SuanfaValue   '�㷨ѡ��
	 
if request.form("submit")<>"�ύ"  then
	Set fsObj=Server.CreateObject("Scripting.FileSystemObject")
	set txtObj = fsObj.OpenTextFile(Server.MapPath("./")&"\z_jifen.inc",1)
	txtObj.skipline()
'jf_StartTime	
	txtObj.skipline()
	txtObj.skip(13)
	jf_StartTime=txtObj.read(2)
'jf_WealthRate	
	txtObj.skipline()
	txtObj.skip(14)
	jf_WealthRate=txtObj.read(2)
'jf_CPRate	
	txtObj.skipline()
	txtObj.skip(10)
	jf_CPRate=txtObj.read(2)
'jf_EPRate	
	txtObj.skipline()
	txtObj.skip(10)
	jf_EPRate=txtObj.read(2)
'SuanfaValue			
	txtObj.skipline()
	txtObj.skip(12)
	SuanfaValue=txtObj.read(1)
	
	set txtObj=nothing
	set fsObj=nothing
%>
                    <p>&nbsp;</p>
                    <p>==================================== <font color="#0066FF">��������</font> 
                      ====================================</p>
                    <p> ʱ����(����)��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input type="text" name="jf_StartTime" size="20" value="<%=jf_StartTime%>">
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���Ա�����ֵ�ʱ����<br>
                      �Ʋ�������(W)��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input type="text" name="jf_WealthRate" size="20" value="<%=jf_WealthRate%>">
                      &nbsp;&nbsp;&nbsp;&nbsp; �Ʋ�ֵ = ����ʱ�� / W <br>
                      ����������(C)��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input type="text" name="jf_CPRate" size="20" value="<%=jf_CPRate%>">
                      &nbsp;&nbsp;&nbsp;&nbsp; ����ֵ = ����ʱ�� / C <br>
                      ����������(E)��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <input type="text" name="jf_EPRate" size="20" value="<%=jf_EPRate%> ">
                      &nbsp;&nbsp;&nbsp;&nbsp; ����ֵ = ����ʱ�� / E </p>
                    <p><font color="#FF99CC">�ο�ֵ��ʱ����:T=<font color="#FF0000">20</font> 
                      ����, �Ʋ�������: W=<font color="#FF0000">10</font>, ����������: C=<font color="#FF0000">30</font>, 
                      ����������: E=<font color="#FF0000">20</font></font></p>
                    <p> ==================================== <font color="#0066FF">�㷨ѡ��</font> 
                      ====================================<br>
                      <br>
                      <input type="radio" name="suanfa" value="1" <%if SuanfaValue="1" then response.write("checked") end if%>>
                      �㷨һ: һ���㷨<br>
                      <br>
                      �Ʋ�ֵ = ����ʱ�� / W <br>
                      ����ֵ = ����ʱ�� / C <br>
                      ����ֵ = ����ʱ�� / E <br>
                      <br>
                      <input type="radio" name="suanfa" value="2" <%if SuanfaValue="2" then response.write("checked") end if%>>
                      �㷨��: ���������㷨<br>
                      <br>
                      �Ʋ�ֵ = (����ʱ��+5) / W <br>
                      ����ֵ = (����ʱ��+5) / C <br>
                      ����ֵ = (����ʱ��+5) / E <br>
                      <br>
                      <input type="radio" name="suanfa" value="3" <%if SuanfaValue="3" then response.write("checked") end if%>>
                      �㷨��: �������롢��Сֵ��С��1�㷨<br>
                      <br>
                      �Ʋ�ֵ = MAX��(����ʱ��+5) / W ��1��<br>
                      ����ֵ = MAX��(����ʱ��+5) / C ��1��<br>
                      ����ֵ = MAX��(����ʱ��+5) / E ��1��</p>
                    <p><font color="#FF99CC">�Ƽ��㷨���㷨��</font><br>
                    </p>
                    <p> 
                      <input type="submit" name="Submit" value="�ύ">
                      ���������� 
                      <input type="reset" name="Submit2" value="����">
                    </p>
                    <p>&nbsp; </p>
                    <%
else
    jf_StartTime=trim(request.form("jf_StartTime"))
    jf_WealthRate=trim(request.form("jf_WealthRate"))
    jf_CPRate=trim(request.form("jf_CPRate"))
    jf_EPRate=trim(request.form("jf_EPRate"))
    SuanfaValue=trim(request.form("Suanfa"))
	
	Set fsObj=Server.CreateObject("Scripting.FileSystemObject")
	set txtObj = fsObj.OpenTextFile(Server.MapPath("./")&"\z_jifen.inc",2,true)
	txtObj.writeline("<%")
	txtObj.writeline("dim jf_StartTime,jf_WealthRate,jf_CPRate,jf_EPRate,SuanfaValue    '���ּ��ʱ�䣬����Ʋ����ʣ������������ʣ����㾭����ʣ��㷨ѡ��")
	txtObj.write("jf_StartTime=")
	txtObj.writeline(jf_StartTime)
	txtObj.write("jf_WealthRate=")
	txtObj.writeline(jf_WealthRate)
	txtObj.write("jf_CPRate=")
	txtObj.writeline(jf_CPRate)
	txtObj.write("jf_EPRate=")
	txtObj.writeline(jf_EPRate)
	txtObj.write("SuanfaValue=")
	txtObj.writeline(SuanfaValue)			
	txtObj.writeline("")
	txtObj.writeline("")
	txtObj.writeline("%\>")
	set txtObj=nothing
	set fsObj=nothing
	response.write("��ϲ�������óɹ�������<br>")
end if
%>
                  </td>
                </tr>
              </table>
            </form>
			<br><br>
          </td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<br>
<br>
<br>
<%
end sub
%>