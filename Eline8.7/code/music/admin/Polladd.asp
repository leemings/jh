<%PageName="admin_ResearchAdd"%>
<!--#include file="function.asp"-->
<%CheckAdmin2%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<div align="center">
<center>
<table border="0" width="90%" cellspacing="1" cellpadding="1">
  <tr>
    <td align=center valign=top>
      <table border="1" width="100%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
        <form method="POST" action="PollSave.asp">
          <tr>
            <td width="100%" height="20" colspan=2 bgcolor="#FF1171" align=center><font color="white"><b>�� �� �� վ �� ��</b></font></td>
          </tr>
          <tr>
            <td width="15%" align="right">���⣺</td>
            <td width="85%"><input type="text" name="Title" size="20"></td>
          </tr>
<%for i=1 to 10%>
          <tr>
            <td align="right">ѡ��<%=i%>��</td>
            <td><input type="text" name="select<%=i%>" size="20"> ����:<input type="text" name="url<%=i%>" size="20"> Ʊ��:<input type="text" name="answer<%=i%>" size="5" value="0"></td>
          </tr>
<%next%>
          <tr>
            <td colspan=2 align=center>
              <input type="hidden" value="add" name="act">
              <input type="submit" value=" �� �� " name="cmdok">&nbsp; 
              <input type="reset" value=" ȡ �� "  name="cmdcancel">
            </td>
          </tr>
        </form>
      </table>
    </td>
  </tr>
</table>
<%
conn.close
set conn=nothing%></body></html>


