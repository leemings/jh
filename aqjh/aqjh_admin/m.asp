<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<html>
<head>
<title>�����µİ��ɣ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: ����}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: ����; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#000000" text="#000000" link="#000080" alink="#800000" vlink="#000080" background="../jhimg/bk_hc3w.gif">
<form action="newzt.asp" method=POST >
  <ul>
    <table border=1 cellspacing=0 cellpadding=0 align="center" width="560" height="104">
      <tr> 
        <td height="13"> 
          <div align="center"><font size="+1"><b>��������</b></font></div>
        </td>
      </tr>
      <tr> 
        <td height="27"> 
          <div align="center"><font color="#FFFFFF"><font color="#000000">��������:</font></font> 
            <input name="mp" value="" size=10 maxlength=10>
            ���ţ� 
            <input name="zm" value="" size=10 maxlength=10>
            ���ƣ� 
            <input name="xz" value="" size=3 maxlength=1>
          </div>
          <div align="center"></div>
          <div align="center"><font color="#FFFFFF"><font color="#000000">���ɼ��:</font></font> 
            <input name="sm" value="" size=27 maxlength=30>
            <font color="#FFFFFF"><font color="#000000"><br>
            </font><font color="#FFFFFF"><font color="#000000">����˵��:</font></font> 
            <input name="xzsm" value="" size=27 maxlength=30>
            <br>
            <font color="#000000">�� �� ʽ:</font></font> 
            <input name="bds" value="" size=27 maxlength=30>
            <font color="#FFFFFF"> <font color="#000000"> </font></font><br>
            ע�����ɼ�����Ϊ30���ַ����Ҳ���Ϊ�գ�</div>
        </td>
      </tr>
    </table>
    <div align="center"> <font size="3" class="c" color="#000000"><br>
      <br>
      <input type="HIDDEN" name="action" value="RegSubmit">
      <input type="SUBMIT" name="Submit" value="ע��">
      <input type="RESET" name="Reset" value="���">
      </font> </div>
  </ul>
</form>
<p align="center">�������֣�1-10���ַ� ���ţ�1-10���ַ����Զ����óɹ�����5�� ��������<br>
  ���ƣ�0 û�� 1���� ����Ϊ1ʱ����˵��������ʽ���������� <br>
  ���ɼ�飺�й����ɵ�˵����дʲô������<br>
  ����˵�������Ǽ���������磺ս���ȼ�&gt;10��<br>
  ����ʽ���ȼ�&gt;10<br>
  <br>
  ����Ӧ�ı���ʽһ��Ҫ����ʽ��д����д��ȷ�����򽫲�����ȷ�������ɣ�</p>
</body>
</html>