<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<meta http-equiv="Content-Language" content="zh-cn">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="760" id="AutoNumber2">
    <tr>
      <td width="100%"><img border="0" src="����.gif" width="1" height="3"></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
<table border="0" width="760" cellspacing="0" cellpadding="0" style="border-collapse: collapse">
  <tr>
    <td valign=top width=150>
    <!--#include file="admin_left.asp"-->
    ��</td>
    <td valign=top width=10>��</td>
    <td valign=top width="600">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber3" bgcolor="#FF9900" height="186">
      <tr>
        <td width="100%" height="186">
        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#333333" width="100%" id="AutoNumber4" height="100%">
<form method="POST" action="Songadd3.asp">
          <tr>
            <td width="100%" align="center" bgcolor="#FFCB7D" colspan="2">�����������
            ��һ��(�趨)</td>
          </tr>
          <tr>
            <td width="28%" align="right">��Ӷ����ף�</td>
            <td width="72%">
              <select name="addnum" size="1">
                <%	for i=1 to 100
						Response.Write "<option>"&i&"</option>"
					next%>
              </select></td>
          </tr>
          <tr>
            <td width="28%" align="right">ѡ��ͳһ���ࣺ</td>
            <td width="72%">
              <select name="Sclassid" size="1">
                <option value="" <%if request("classid")="" then%> selected<%end if%>>ѡ����Ŀ</option>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from Sclass"
rs.open sql,conn,1,1
do while not rs.eof
%>
                <option<%if cstr(request("Sclassid"))=cstr(rs("Sclassid")) and request("Sclassid")<>"" then%> selected<%end if%> value="<%=CStr(rs("SclassID"))%>" name=Sclassid><%=rs("Sclass")%></option>
<%
rs.movenext
loop
rs.close
%>
              </select></td>
          </tr>
          <tr>
            <td width="28%" align="right">��дͳһ���ƣ�</td>
            <td width="72%">
            <input type="text" name="MusicName" size="26"></td>
          </tr>
          <tr>
            <td width="28%" align="right">��дͳһ�����ˣ�</td>
            <td width="72%">
            <input type="text" name="DJ Anjo" size="18"></td>
          </tr>
          <tr>
            <td width="28%" align="right">��дͳһ���ӵ�ַ��</td>
            <td width="72%">
            <input type="text" name="LF_Path" size="52"></td>
          </tr>
          <tr>
            <td width="27%" align="right">ѡ��ͳһ���ŷ�ʽ��<br>
            <font color="#FF0000">Ĭ����RealPlayer����</font></td>
            <td width="73%"><select size="1" name="MusicType">
            <option value="RealPlayer">RealPlayer</option>
            <option value="MediaPlayer">MediaPlayer</option>
            <option value="Flash">Flash</option>
            </select><font color="#FF0000">��һ��Ҫע���ļ��ĸ�ʽ����ѡ���Ӧ�Ĳ��ŷ�ʽ��</font></td>
          </tr>
          <tr>
            <td width="100%" align="center" colspan="2" bgcolor="#FFCB7D">
            <p align="center">
            <input type="hidden" value="add" name="act">
              <input type="submit" value=" �� �� �� һ ��" name="cmdok"></td>
          </tr>
          </form>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  </table>
  </center>
</div>