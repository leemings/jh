<%PageName="admin_AdminMana"%>
<!--#include file="function.asp"-->
<%CheckAdmin3%>
<!--#include file="conn.asp"-->
<!--#include file="top.asp"-->
<table border="0" width="100%" cellspacing="0" cellpadding="5">
  <tr>
    <td align=center valign=top>
      <table border="1" width="88%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
          <tr>
            <td width="100%" height="20" colspan=7 bgcolor="#FF1171" align=center><font color="white"><b>�� �� Ա �� ��</b></font></td>
        </tr>
        <tr>
          <td width="25%" align="center" bgcolor="#FFD2E4" height="22">����Ա��</td>
          <td width="20%" align="center" bgcolor="#FFD2E4">Ȩ��</td>
          <td width="10%" align="center" bgcolor="#FFD2E4">�޸�</td>
          <td width="10%" align="center" bgcolor="#FFD2E4">ɾ��</td>
        </tr>
<%
set rs=server.CreateObject("ADODB.RecordSet")
	  
sql="select * from admin"
rs.open sql,conn,1
%> 
<%
if rs.EOF then
%>
        <tr><td colspan=5 align=center>û���û�����</td></tr>
<%
else
	do while NOT rs.EOF
if rs("oskey")="super" then oskey="�߼�����Ա"
if rs("oskey")="check" then oskey="�м�����Ա"
if rs("oskey")="input" then oskey="��������Ա"
%> 
        <tr> 
          <td width="25%" align="center" height="19"><%=rs("Username")%></td>
          <td width="20%" align="center"><%=oskey%>��</td>
          <td width="10%" align="center"><a href="AdminModify.asp?id=<%=rs("id")%>">�޸�</a></td>
          <td width="10%" align="center"><a href="AdminDel.asp?id=<%=rs("id")%>">ɾ��</a></td>
        </tr>
<%
	rs.MoveNext
	loop
end if
rs.close
%> 
      </table>
      <FORM METHOD=POST ACTION="AdminSave.asp" id=form1 name=form1>
        <table border="1" width="40%" cellspacing="0" cellpadding="0" class="TableLine" bordercolor="#FF1171" bordercolordark="#FFFFFF">
          <tr> 
            <td align="center" bgcolor="#FF1171" height=20 colspan=2><font color="white"><b>�� �� �� �� Ա</b></font></td>
          </tr>
          <tr>
            <td align="right">�� �� Ա ����</td>
            <td><input type=text name=UserName size="15"  value="" onfocus=this.select() onmouseover=this.focus() name=keyword size=14 maxlength="30"></td>
          </tr>
          <tr> 
            <td align="right">�� �� Ȩ �ޣ�</td>
            <td>
              <select name="oskey">
                <option value=super selected>�߼�����Ա</option>
                <option value=check>�м�����Ա</option>
                <option value=input>��������Ա</option>
              </select>
            </td>
          </tr>
          <tr> 
            <td align="right">�� �� �� �룺</td>
            <td><input type=text name=Password size="15"  value="" onfocus=this.select() onmouseover=this.focus() name=keyword size=14 maxlength="30"></td>
          </tr>
          <tr> 
            <td align="center" colspan=2> 
              <input type=hidden value="add" name="act">
              <input type=submit value=���� name="submit">
              <input type=reset name="Submit" value="ȡ��">
            </td>
          </tr>
        </table>
      </FORM>
    </td>
  </tr>
</table>
</div>
<%
set rs=nothing
conn.close
set conn=nothing
%></body></html>




