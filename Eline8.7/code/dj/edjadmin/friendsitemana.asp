<!--#include file="admin_top.asp"-->
<!--#include file="checkadmin.inc"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from FriendSite order by id desc"
rs.open sql,conn,1,1
%>
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
    <td valign=top width="600" bgcolor="#FF9900">
      <table border="0" width="100%" cellspacing="3" cellpadding="3" style="border-collapse: collapse">
        <form method="POST" action="FriendSiteSave.asp?act=set">
          <tr>
            <td width="100%" height="20" colspan=9 align=center bgcolor="#CC6600"><b>�� �� �� �� վ ��</b></td>
          </tr>
          <tr>
            <td width="5%" align="center">ѡ��</td>
            <td width="5%" align="center">ID</td>
            <td width="10%" align="center">վ��</td>
            <td width="10%" align="center">��ַ</td>
            <td width="10%" align="center">���</td>
            <td width="20%" align="center">��վLogo</td>
            <td width="10%" align="center">վ��</td>
            <td width="5%" align="center">�޸�</td>
            <td width="5%" align="center">ɾ��</td>
          </tr>
<%
do while not rs.eof
%>
          <tr>
            <td width="5%"><input type="checkbox" name="checked" value="<%=rs("ID")%>"<%if rs("IsOK")=true then%> checked<%end if%>></td>
            <td width="5%" align="center"><%=rs("ID")%>��</td>
            <td width="10%" align="center"><%=rs("SiteName")%>��</td>
            <td width="10%" align=center><a href="<%=rs("SiteUrl")%>" target="_blank" title="<%=rs("SiteUrl")%>">��վ��ַ</a></td>
            <td width="10%" align=center><a style="cursor:hand" title="<%if rs("SiteIntro")="" or isnull(rs("SiteIntro")) then%>��<%else%><%=rs("SiteIntro")%><%end if%>">��վ���</td>
            <td width="20%" align=center><%if not isNull(rs("LogoUrl")) then%><img src="<%=rs("LogoUrl")%>" width=88 height=31 border=0 alt="<%=rs("LogoUrl")%>"><%end if%></td>
            <td width="10%" align="center"><%=rs("SiteAdmin")%>��</td>
            <td width="5%" align="center"><input onclick="javascript:window.open('FriendSiteModify.asp?id=<%=rs("ID")%>','_self','')" type="button" value="�޸�" name="button2"></td>
            <td width="5%" align="center"><input onclick="javascript:window.open('FriendSiteSave.asp?act=del&id=<%=rs("ID")%>','_self','')" type="button" value="ɾ��" name="button3"></td>
          </tr>
<%
rs.movenext
loop
rs.close
%>
          <tr>
            <td colspan=9 align=center bgcolor="#CC6600"><input type="submit" value="ѡ��" name="submit"> <input onclick="javascript:window.open('FriendSiteAdd.asp','_self','')" type="button" value="�������վ" name="button"></td>
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
<%
set rs=nothing
conn.close
set conn=nothing%>