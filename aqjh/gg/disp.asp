<%
id=request("id")
if not isnumeric(id) then
	Response.Write "<script language=JavaScript>{alert('��ʾ������');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select * from ���� where id="&id,conn,1,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language=JavaScript>{alert('�����Ҳ�����Ҫ���Ĺ��棡��');window.close();}</script>"
	response.end
end if
bt=rs("����")
fbsj=rs("ʱ��")
nr=rs("����")
rs.close
conn.execute "update ���� set �鿴����=�鿴����+1 where id="&id
set rs=nothing
conn.close
set conn=nothing
if nr="" or nr=null then
	nr="������"
end if
%>
<html>
<STYLE type=text/css>BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 18px
}
TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 18px
}
A:visited {
	COLOR: black; TEXT-DECORATION: none
}
A:link {
	COLOR: black; TEXT-DECORATION: none
}
A:hover {
	COLOR: red; TEXT-DECORATION: none
}
</STYLE>

<head>
<title>�����ֽ���������www.happyjh.net���쳱����������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<META NAME="Keywords" CONTENT="���ֽ������������������ֽ��������潭����������Ϸ���������У����ѣ����죬���֣���Ӱ��Ц����SHOW�������㣬���������������ֽ�����������Ϸ����������">
<META NAME="Description" CONTENT="���ֽ�������������������Ϸ���쳱����������Ǵ�ҽ�������ĺ�ȥ�������ܶ�࣬��ӭ������飡">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<BODY background=../jhimg/back002.gif bgColor=#736745 text=#000000 topMargin=0 marginheight="0" marginwidth="0" leftmargin="0">
<div align=center><table width="770" border="0" align="center" cellpadding="0" cellspacing="0">
  </table><div align=center>
<TABLE WIDTH=770 BORDER=0 CELLPADDING=0 CELLSPACING=0 bgcolor="#B56F00">
    <TR> 
      <TD width="10" bgcolor="#B56F00"></TD>
      <TD> <div align=center><table border="0" cellspacing="0" cellpadding="0" width="648" height="298">
<tr>
<td width="22"><img src="../jhimg/jiao_z_s.gif" width="21" height="21"></td>
<td background="../jhimg/line_s.gif" width="622"><font color=#e2d1c0></font></td>
<td width="10"><img src="../jhimg/jiao_y_s.gif" width="22" height="21"></td>
</tr>
<tr>
<td background="../jhimg/line_z.gif" width="22"></td>
<td width="622">
<table align=center bgcolor=#644e49 border=1 cellpadding=5 cellspacing=3 width="100%">
<tbody>
<tr>
<td height="245" valign="top"><font color=#e2d1c0>
<p><br>
<br>
</p>
<p align="center"><span class=title><font size="4" face="����"><b><%=bt%></b></font></span>
</p>
<p align="center">�����ڣ�<%=fbsj%><br>
</p>
</font>
<div align="right">
<p style="margin-right: 61">
<input type="button" name="Submit" onclick="javascript:window.close()" value="�رմ���" style="background-color: #644E49; color: #FCCF49; border: 1 solid #8D6D69">
<font color=#e2d1c0></font></div>
<hr>
<font color=#e2d1c0>
<p style="text-indent: 19; line-height: 150%; margin-left: 38; margin-right: 38"><%=nr%></p>
</font>
<div align="center"><br>
<br>
<hr>
<br>
<input type="button" onclick="javascript:window.close()" name="close" value="�رմ���" class="umberbutton">
<br>
<br>
</div>
</td>
</tr>
</tbody>
</table>
</td>
<td background="../jhimg/line_y.gif" width="10"></td>
</tr>
<tr>
<td width="22"><img src="../jhimg/jiao_z_x.gif" width="22" height="21"></td>
<td background="../jhimg/line_x.gif" width="622"></td>
<td width="10"><img src="../jhimg/jiao_y_x.gif" width="21" height="21"></td>
</tr>
</table></div></TD>
      <TD width="10" bgcolor="#B56F00"></TD>
    </TR>
  </TABLE></div>
<TABLE WIDTH=770 BORDER=0 CELLPADDING=0 CELLSPACING=0 bgcolor="#B56F00">
    <TR> 
      <TD width="10" bgcolor="#B56F00"></TD>
      <TD> <TABLE WIDTH=750 BORDER=0 CELLPADDING=0 CELLSPACING=0>
          <TR> 
            <TD> <IMG SRC="../../newjh/newjh_hb.gif" WIDTH=750 HEIGHT=10 ALT=""></TD>
          </TR>
          <TR> 
            <TD height="58" align="center" bgcolor="#412800"><font color="#FF9900">���ʹ�÷ֱ��ʣ�800��600��Windows��Internet 
              Explorer V6.0 or higher��<br>
              Copyright 2003-2004 ��www.happyjh.com�� All Rights Reserved ���ֽ�����վ 
              ����Ȩ���� ��ֹ��Ϯ<br>
              վ��ά���������ֽ���������QQ��865240608</font></TD>
          </TR>
          <TR> 
            <TD> <IMG SRC="../../newjh/newjh_hb.gif" WIDTH=750 HEIGHT=10 ALT=""></TD>
          </TR>
        </TABLE></TD>
      <TD width="10" bgcolor="#B56F00"></TD>
    </TR>
  </TABLE>
<br>
</div>
</BODY></HTML>
