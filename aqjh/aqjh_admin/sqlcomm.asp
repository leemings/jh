<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
%>
<html>
<head>
<title>sqlָ��ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<SCRIPT LANGUAGE="JavaScript">
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.sqlstr.value;
revisedTitle = addTitle;
document.PostTopic.sqlstr.value=revisedTitle;
document.PostTopic.sqlstr.focus();
return; }
</script>
<LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><%=Application("aqjh_chatroomname")%>���ݿ�<br>
  <br>
  <b><font color="#FF0000">ע�⣺�˴���������Ӱ�쵽���ݿ�����ϵͳ��<br>
  ���鲻���׵�վ����Ҫʹ�ã��ǳ�Σ�գ�</font></b> </p>
<p align="left">��������Ҫִ�е�ָ�<br>
  �磺update �û� set grade=1,���='����' where ���&lt;&gt;'����' and ����&lt;&gt;'�ٸ�'<br>
  �ǽ�������ݲ�Ϊ���ţ����ɲ�Ϊ�ٸ��Ĺ���ȼ����1��,������ó�Ϊ���ӣ��� <br>
  �磺delete * from w where i=0<br>
  �ǽ���Ʒ���ݿ��е���ƷΪ0�ļ�¼ɾ����<br>
  �磺delete * from �书 where ����='�ٸ�'<br>
  �ǽ��ٸ����书ɾ����
</p>
<p> �ϲ�����֮ǰ��ִ�� <br>
  update �û� set grade=1, ���='��' where ����='����A'���='����' <br>
  update �û� set grade=1, ���='��' where ����='����A'���='����' <br>
  update �û� set grade=1, ���='��' where ����='����A'���='����' <br>
  update �û� set grade=1, ���='��' where ����='����B'���='����' <br>
  update �û� set grade=1, ���='��' where ����='����B'���='����' <br>
  update �û� set grade=1, ���='��' where ����='����B'���='����' <br>
  update �û� set grade=1, ���='��' where ����='����A'���='����' <br>
  update �û� set grade=1, ���='��' where ����='����A'���='����' <br>
  �����ִ�л����ܶ೤������������ <br>
  �ٰѱ��ϲ������ɵ�����Ȩ���û�����OK�� </p>
</p>
<p align="left"><br>
  <br>
</p>
<%if aqjh_name=Application("aqjh_user") then%>
<form method="POST" action="sqlcommok.asp" name="PostTopic">
 <div align="center">
<SELECT name=fs onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">����ָ��</OPTION> 
<OPTION value="update b set o=1000 where a='��Ʒ��' and n<>'��'">[�޸ľ�������]</OPTION>
<OPTION value="update �û� set ����=0,���=0 where (����+���)>1000000000">[����ʲ�����10���û��ʲ�Ϊ0]</OPTION>
<OPTION value="update �û� set grade=1,���='����' where ����='����'">[��������ɹ�����Ա]</OPTION>
<OPTION value="delete * from y where b='�ٸ�'" >[ɾ���ٸ��书]</OPTION>
<OPTION value="delete * from l " >[ɾ�����в�����¼]</OPTION>
<OPTION value="delete * from l where d='����'" >[ɾ�����н���������¼]</OPTION>
<OPTION value="delete * from j " >[ɾ�������ռ�����]</OPTION>
<OPTION value="delete * from h where c='���'" >[ɾ���������ϼ�¼]</OPTION>
<OPTION value="delete * from h where c='���'" >[ɾ����������¼]</OPTION>
<OPTION value="delete * from t where b='����'" >[ɾ���������ĺ��弼]</OPTION>
<OPTION value="delete * from �û� where allvalue<=0" >[ɾ�����allvalueΪ0���û�]</OPTION>
<OPTION value="delete * from �û� where ��Ա��>500" >[ɾ����Ա�𿨴���500���û�]</OPTION>
<OPTION value="delete * from �û� where allvalue<=2 and month(��¼)<>month(now())" >[ɾ�����ݵ�С��2���û�]</OPTION>
<OPTION value="update �û� set W8='�±�|10000;' where ����='����'">[����һ��ָ����Ʒ]</OPTION>
<OPTION value="update �û� set allvalue=174050 where allvalue>174050">[�ѵȼ�����59�Ľ���59��]</OPTION>
<OPTION value="update �û� set ����=0,���=100000000 where ����<>'�ٸ�' and ����+���>100000000">[Ǯ����1�ڵĽ���1��]</OPTION>
<OPTION value="update �û� set W3=W3+'���齣|1;' where ����='�û���'">[��ĳ��һ�����齣]</OPTION>
<OPTION value="update �û� set zw='����|����|2002-7-22 21:05:06|75000|21000|14000|4|2002-9-20 4:13:49' where ����='�û���'">[��ĳ��һ������]</OPTION>
<OPTION value="update �û� set ״̬='����',�¼�ԭ��='��' where lastip='ip��ַ'">[���ĳ��IP]</OPTION>
<OPTION value="update �û� set ���=���+150000000 where grade=5">[������Ǯ��SQL]</OPTION>
<OPTION value="update �û� set allvalue=allvalue+1000">[��������1000�����]</OPTION>
<OPTION value="update �û� set ����='������A' where ����='������B'">[�ϲ�����]</OPTION>
<OPTION value="update �û� set ����='����' where ����='����'">[��ĳ�û�����]</OPTION>
<OPTION value="update �û� set �Ա�='�л�Ů' where ����='�û���'">[��ĳ�˸��Ա�]</OPTION>
<OPTION value="update �û� set ����=����+1000 where ����='�û���'">[����ĳ�˷���]</OPTION>
<OPTION value="update �û� set w1=' ',w2=' ',w3=' ',w4=' ',w5=' ',w6=' ',w7=' ',w8=' ',w9=' ',w10=' ',z1=' ' where ����='���׵���'">[�����û�������Ʒ]</OPTION>
</SELECT><br>
   <br>
    ������ָ� 
    <input type="text" name="sqlstr" size="50" maxlength="280">
    <br>
    <input type="submit" value="ִ��" name="B1" class="p9">
    <input type="reset" name="Submit" value="���">
  </div>
  </form>
 <%else%>
  <div align="center">������Ȩվ���������ֹʹ�ã�</div>
 <%end if%>
</body>
</html>