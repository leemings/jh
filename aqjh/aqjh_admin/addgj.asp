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
<html>
<head>
<title>����µĹ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css/css.css" type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<form action="addgjok.asp" method=POST >
  <ul>
    <table border="0" cellspacing="1" cellpadding="4" bgcolor="#B8AF86" cellspacing=0 cellpadding=0 align="center" width="560" height="104">
      <tr> 
        <td height="13" bgcolor=#336633> 
          <div align="center"><b><font color=ffffff>�������</b></div>
        </td>
      </tr>
      <tr bgcolor="f2f2ea"> 
        <td height="27"> 
          <div align="center"><font color="#FFFFFF"><font color="#000000">��������:</font></font> 
            <input name="gj" value="" size=10 maxlength=10 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
            ������ 
            <input name="jz" value="" size=10 maxlength=10 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
          </div>
          <div align="center"></div>
          <div align="center"><font color="#FFFFFF"><font color="#000000">���Ҽ��:</font></font> 
            <input name="sm" value="���ﻶӭ���ļ��룡" size=27 maxlength=30 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
            <font color="#FFFFFF"><font color="#000000"><br>
            </font><font color="#FFFFFF"><font color="#000000">����˵��:</font></font> 
            <input name="xzsm" value="����˹���������" size=27 maxlength=30 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
            <br>
            <font color="#000000">�� �� ʽ:</font></font> 
            <input name="bds" value="True" size=27 maxlength=30 class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
            <font color="#FFFFFF"> <font color="#000000"> </font></font><br>
            ע�����Ҽ�����Ϊ30���ַ����Ҳ���Ϊ�գ�</div>
        </td>
      </tr>
    </table>
    <div align="center"> <font size="3" class="c" color="#000000"><br>
      <br>
      <input type="HIDDEN" name="action" value="RegSubmit">
      <input type="SUBMIT" name="Submit" value="ע��" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
      <input type="RESET" name="Reset" value="���" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
      </font> </div>
  </ul>
</form>
<p align="center">�������֣�1-10���ַ� ������1-10���ַ�<br>
  ���Ҽ�飺�йع��ҵ�˵����дʲô������<br>
  ����˵�������Ǽ���������磺ս���ȼ�&gt;10��<br>
  ���ʽΪ���ȼ�&gt;10<br>
  ��ֻ����2������Ů��Ҽ��� <br>
  ���ʽΪ���ȼ�&gt;=2 and �Ա�='Ů'<br>
  ������˼���������룺True<br>
  <br>
  <br>
  ����Ӧ�ı��ʽһ��Ҫ����ʽ��д����д��ȷ�����򽫲�����ȷ������ң�</p>
</body>
</html>
