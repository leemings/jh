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
rs.open "Select  * from �û� where ����='"&sjjh_name&"'",conn
yin=rs("����")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML>
<HEAD>
<title>�Ĵ�С��һ�������wWw.happyjh.com��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
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
</HEAD>
<body text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../jhimg/bk_hc3w.gif">
<div align="left"></div>
<div align=center> 
  <p><font color="#000000" size="+3">�Ĵ�С</font><font size="2"><br>
    ������ע������ <b><font color="#CC0000">10 </font> ��<br>
    </b>�����ע������ <font color="#CC0000"><b>3000</b></font><b> ��</b> <br>
    ������һ�������� <b><font color="#CC0000"><%=yin%></font> ��</b> ������Ϊ��ע</font></p>
  <table width="545" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td> 
        <form method="POST" action="b&amp;spose.asp">
          <table border=1 cellspacing=0 cellpadding=0 align="center" width="350" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr align="center" bgcolor="#FFFFFF"> 
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
            </tr>
            <tr bgcolor="#336633"> 
              <td width="960" colspan="3" height="29"><font size="2" class="c" color="#000000"><b>&nbsp;&nbsp;<font color="#FFFFFF">��ѡ��</font></b></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td align=center colspan="3"> 
                <table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                  <tr align="center"> 
                    <td width="50%"><img src="../jhimg/bbs/big.gif" width="46" height="40"></td>
                    <td width="50%"><img src="../jhimg/bbs/small.gif" width="46" height="40"></td>
                  </tr>
                  <tr align="center"> 
                    <td width="50%"> 
                      <input type="radio" name="select" value="big" checked>
                    </td>
                    <td width="50%"> 
                      <input type="radio" name="select" value="small">
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr> 
              <td align=center colspan="3"><font size="2" color="#000000">��Ҫ��ע�� 
                <input type="text" name="splosh" size="4" value="0" maxlength="4">
                &nbsp;<b>��</b></font></td>
            </tr>
            <tr> 
              <td align=center colspan="3"> <font size="2" color="#000000"> 
                <input type="submit" value="��ע��������" name="B1" >
                <input type="reset" value="��Ҫ����һ�£���" name="B2" >
                </font></td>
            </tr>
          </table>
        </form>
        <p align="center"><font size="2">���棺ÿ�ζ�Ǯ���������ֵ��� 1 ��</font></p>
      </td>
    </tr>
  </table>
<p align="center">�����ֽ�����</p>
</div>
</BODY>
</HTML>
