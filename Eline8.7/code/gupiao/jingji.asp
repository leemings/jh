<html>
<head>
<title>��Ʊ�����˰칫��</title>
<style>
input,body,select,td{font-size:14;line-height:160%}
</style>
</head>
<body bgcolor="#000000">
<p align=center style='font-size:16;color:yellow'>��ӭ<%=sjjh_name%>���ٹ�Ʊ�����˵İ칫�ң�
<p align=center style='font-size:14;color:blue'><a href=index.asp>���ع�Ʊ�г�</a></p>
<!--#include file="jhb.asp"-->
<%
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
sql="select ������ from �ͻ� where �ʺ�='"&sjjh_name&"'"
set rs=conn.execute(sql)
if not rs.eof then
jjr=rs(0)
%>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p> 
</td></tr>  

<tr><td style="color:red;font-size:9pt">  
<br><p align=center>���Ĺ�Ʊ������<%=jjr%>�ڴ�Ϊ������</p><br>
<a href=cha.asp>��ѯ���ս���</a>
<br>  
<p align=center><a href=index.asp>�뿪</a></p>
</table></table>  
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing%>
<p align=center style='font-size:16;color:yellow'>�����Ʊ�ʻ�/���������� 
<form method=POST action='jingji2.asp' id=form2 name=form2>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
        <table width=100%>
          <tr> 
            <td width="47%">�ʺţ� 
              <input type=text name=name size=10 value="<%=sjjh_name%>" maxlength="0">
            </td>
            <td width="53%">�����ˣ� 
              <select name=jjr size=1>
                <option value="��Pip">��Pip</option>
                <option value="ǧ�����">ǧ�����</option>
                <option value="ӡ�Ȱ���">ӡ�Ȱ���</option>
                <option value="Ǯ����">Ǯ����</option>
                <option value="������">������</option>
                <option value="�����">�����</option>
                <option value="Լ����">Լ����</option>
                <option value="ɳ¡��˹">ɳ¡��˹</option>
                <option value="��̫��">��̫��</option>
                <option value="��Ǯ��">��Ǯ��</option>
                <option value="����˿">����˿</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td colspan=2 align=center> 
              <input type=submit value=ȷ�� id=submit2 name=submit2>
              <input type=reset value="ȡ��" id=reset2 name=reset2>
            </td>
          </tr>
          <tr> 
            <td colspan=2 style='font-size:9pt'> 
              <hr>
              1.����ʱ��ֻ�����һ�����룻���ǽ�ͬʱΪ������һλ��Ʊ�����ˣ���������ɹ�Ʊ���ף�ÿ����������ȡ���׽���1%��Ϊ���͡�<br>
              2.�޸�����ʱ����һ������Ϊ�����룬�ڶ�������Ϊ�����룡���ڱ�ϵͳ�����û����ܵİ취��������ֱ���͵Ǯ�����飬����������Բ����޸ģ� 
              <br>
              3.����ʱ�ʺſ���������д������ÿ��������ʿֻ��������һ����Ʊ�ʺţ�</td>
          </tr>
            <td width="47%"> 
          </table>
</td></tr>
</table>
</form>
</body>
</html>