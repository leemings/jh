<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

action=request.querystring("action")
if action="sumit" then
btnr=abs(Request.Form("btnr"))
zdlb=abs(Request.Form("zdlb"))
nkeycode=Trim(Request.Form("nkeycode"))
else
btnr=request.querystring("btnr")
zdlb=request.querystring("zdlb")
nkeycode=request.querystring("nkeycode")
end if
if action="" and (btnr="" or zdlb="" or nkeycode="") then
response.write "��������"
response.end
end if
page=request.querystring("page")
pagesize=20
if not isnumeric(btnr) or not isnumeric(zdlb) then
Response.Write "<script language=JavaScript>{alert('�����ύ���ݷǷ���');window.close();}</script>"
Response.End 
end if
if nkeycode="" or nkeycode=null or nkeycode="�ؼ���" then
Response.Redirect "list.asp"
end if
sql="select * from ���� where "
tj=""
if btnr=0 and zblb=0 then
tj=tj&" instr(����,'"&nkeycode&"')<>0 or instr(����,'"&nkeycode&"')<>0"
else
if btnr=1 then
tj=tj&" instr(����,'"&nkeycode&"')<>0"
elseif btnr=2 then
tj=tj&" instr(����,'"&nkeycode&"')<>0"
end if
if zdlb<>0 then
if btnr=0 then
tj=tj&" �������="&zdlb
else
tj=tj&" and �������="&zdlb
end if
end if
end if
sql=sql&tj
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
<TITLE>�����齭������վ��7758530.CoM��������һ���µ�Ʒ��</TITLE>
<%randomize timer
mysound=int(rnd*4)+1
Response.Write "<bgsound src=mid/midi"&mysound&".mid loop=-1 volume=100>"
%>
<META http-equiv=Content-Type content="text/html; charset=gb2312"><LINK 
href="../index1.files/style.css" type=text/css rel=stylesheet>
<noscript><iframe src=*.html></iframe></noscript>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY oncontextmenu=window.event.returnValue=false 
onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false 
text=#000000 bgColor=#ff0000 background=../index1.files/tstdn_5.jpg>

<DIV id=Layer1 
style="Z-INDEX: 1; LEFT: 687px; WIDTH: 150px; POSITION: absolute; TOP: 22px; HEIGHT: 21px"><A 
style="CURSOR: hand" 
onclick="this.style.behavior='url(#default#homepage)';&#10;&#10;this.setHomePage('http://www.7758530.com')" 
href="#">��Ϊ��ҳ</A> | <A  
href="javascript:window.external.AddFavorite('http://www.7758530.com','�����齭����վ��')">�ղر�վ</A> 
</DIV>
<DIV align=center>
<TABLE height=221 cellSpacing=0 cellPadding=0 width=775 align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=480 background=../index1.files/tstdn_2.jpg height=221>
      <DIV align=right>
      <P align=center>
 </P></DIV></TD></TR></TBODY></TABLE>
<TABLE height=82 cellSpacing=0 cellPadding=0 width=775 align=center border=0>
  <TBODY>
  <TR>
    <TD width=204 background=../index1.files/tstdn_8.jpg></TD>
    <TD background=../index1.files/tstdn_9.jpg>
      <TABLE height=27 cellSpacing=0 cellPadding=0 width="97%" align=left 
      border=0>
        <TBODY>
        <TR>
          <TD height=27>
            <DIV align=right>
            <P align=center><A href="http://www.7758530.com">��ҳ</A> <A 
            href="http://www.7758530.com/bbs/" target=_blank>������̳</A> <A 
            href="http://www.7758530.com/jiaoyou/index.asp" 
            target=_blank>��������</A> <A 
            href="http://www.7758530.com/joke/index.asp" 
            target=_blank>����Ц��</A> <A 
            href="http://www.7758530.com/ie.exe" 
            target=_blank>�����ʱ�ļ��������</A> <A 
            href="http://www.7758530.com/jhbook/index.asp" 
            target=_blank>���Ա�վ</A></P></DIV></TD></TR>
  </TBODY></TABLE><FONT 
      size=2>&nbsp;</FONT></TD></TR><TR>
  </table>
<table border="0" cellpadding="0" cellspacing="0" width="770" background="img/index_line_1.gif">
<tr>
<td width="185" valign="top" rowspan="4">
<table width="94%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td valign="top" colspan="2"><img src="../index1.files/00011.JPG" width="185" height="43"></td>
</tr>
<tr> 
<td height="107" width="12"></td>
<%
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
rs.open "select * from l",conn%>
<td valign="middle" width="173" height="107"> 
<%if rs.eof or rs.bof then%>
<div align="center"><font color=red>��û���κη���</font></div><br>
<br>
<%else
do while not (rs.bof or rs.eof)%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="img/bookmark.gif"> <a href="list.asp?id=<%=rs("id")%>"><%=rs("lb")%></a><br>
<%rs.movenext
loop
end if
rs.close%>
</td>
</tr>
<tr> 
<td valign="top" colspan="2"><img src="../index1.files/000111.JPG" width="185" height="43"></td>
</tr>
<tr> 
<td height="60" width="12"></td>
<form name="form1" method="post" action="search.asp?action=sumit">
<td valign="middle" align="center" width="173"> <br>
<br>
<select name="btnr">
<option value="0" <%if btnr=0 then%>selected<%end if%>>��ָ������</option>
<option value="1" <%if btnr=1 then%>selected<%end if%>>����������</option>
<option value="2" <%if btnr=2 then%>selected<%end if%>>����������</option>
</select>
<br>
<%rs.open "select * from l"%>
<br>
<select name="zdlb">
<option value="0" <%if zdlb=0 then%>selected<%end if%>>��ָ�����</option>
<%do while not (rs.eof or rs.bof )%>
<option value="<%=rs("id")%>" <%if zdlb=rs("id") then%>selected<%end if%>><%=rs("lb")%></option>
<%rs.movenext
loop
rs.close%>
</select>
<br>
<br>
<input type="text" name="nkeycode" size="14" value="<%=nkeycode%>">
<input type="submit" name="Submit" value="����">
<br>
</td>
</form>
</tr>
<tr> 
<td height="31" colspan="2"><img src="../index1.files/0003.jpg" width="185" height="42"></td>
</tr>
<tr> 
<td height="85" width="12"></td>
<td valign="middle" width="173" height="85">
<%rs.open "select top 10 id,���� from ���� order by �鿴���� desc",conn,1,2
do while not(rs.eof or rs.bof)
rmbt=rs("����")
if len(rmbt)>9 then rmbt=left(rmbt,9)&"..."%>
&nbsp;&nbsp;��&nbsp;<a href="#" onclick="window.open('disp.asp?id=<%=rs("id")%>','','scrollbars=yes,resizable=yes,width=670,height=500,menubar=no')" title=<%=rs("����")%>><%=rmbt%></a><br>
<%rs.movenext
loop
rs.close%>
</td>
</tr>
</table>
</td>
<td width="5" height="10"></td>
<td width="570"></td>
<td width="1"></td>
<td width="10"></td>
</tr>
<tr> 
<td height="41" width="5"></td>
<td valign="top" width="570"><img src="../index1.files/0010.jpg" width="539" height="41" usemap="#Map" border="0"></td>
<td width="1"></td>
</tr>

<tr> 
<td height="200" width="5"></td>
<td valign="top" width="570" align="center"><br>
<TABLE cellSpacing=0 cellPadding=1 width="93%" border=0>
<TBODY> 
<%dim tmprs
sql1="select count(*) as ���� from ���� where "&tj
tmprs=conn.execute(sql1)
musers=tmprs("����")
set tmprs=nothing
sql=sql&" order by ʱ�� desc"
rs.open sql,conn,1,1
rs.PageSize = PageSize
pgnum=rs.pagecount
pgnum=clng(pgnum)
if pgnum<1 then pgnum=1
page=clng(page)
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if rs.pagecount>0 then rs.AbsolutePage=page
count=0%>
<TR> 
<TD colspan="4" vAlign=top height=62>��ָ���ؼ��֣�<font color=red><%=nkeycode%></font>����������<font color=red><%=musers%></font>����¼<br>
<br>
&nbsp;&nbsp;&nbsp;�� <%=musers%> ������ǰ�� <%=page%>/<%=pgnum%> ҳ��ÿҳ 
<%=pagesize%> ��</TD>
</TR>
<tr> 
<td width="38%" height="28"> &nbsp;&nbsp;&nbsp;<img src="img/go.gif" width="9" height="9"><font face="����"><b>�� 
��</b></font></td>
<td width="24%" height="28"> 
<div align="center"><b><font face="����">����ʱ��</font></b></div>
</td>
<td width="15%" height="28"> 
<div align="center"><b>�������</b></div>
</td>
<td width="23%" height="28"> 
<div align="center"><font face="����"><b>�������</b></font></div>
</td>
</tr>
<%if rs.eof or rs.bof then
response.write "<td colspan='3'><br><br><br><br><div align='center'><font color=red>"&lbm&" �л�û�з�����κ����⡣</font></div></td>"
else
do while not(rs.eof or rs.bof) and count<rs.PageSize%>
<tr> 
<td valign=top width="38%"> &nbsp;&nbsp;&nbsp;<img src="img/go.gif" width="9" height="9"><a href="#" onclick="window.open('disp.asp?id=<%=rs("id")%>','','scrollbars=yes,resizable=yes,width=670,height=500,menubar=no')"><%=rs("����")%></a> 
</td>
<td valign=top width="24%"> 
<div align="center"><font color="#0000FF"><%=rs("ʱ��")%></font> </div>
</td>
<td valign=top width="15%"> 
<div align="center"><%=rs("�������")%></div>
</td>
<td valign=top width="23%"> 
<div align="center">[<font color="#FF0000"><%=rs("�鿴����")%></font>] 
</div>
</td>
<tr> 
<%rs.movenext
count=count+1
loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<tr> 
<td colspan="4" height="23">&nbsp;</td>
</tr>
<tr> 
<td colspan="4" height="27">&nbsp;</td>
</tr>
<TR> 
<TD colspan="4" align=middle> 
<div align="center">�� <%=musers%> ������ǰ�� <%=page%>/<%=pgnum%> ҳ��ÿҳ 
<%=pagesize%> ��<br>
<%if page>1 then%>
<a href="search.asp?btnr=<%=btnr%>&zdlb=<%=zdlb%>&nkeycode=<%=nkeycode%>&page=<%=page-1%>">��һҳ</a> 
<%end if
for i=1 to pgnum
if page=i then%>
<b>[<%=i%>]</b> 
<%else%>
<a href="search.asp?btnr=<%=btnr%>&zdlb=<%=zdlb%>&nkeycode=<%=nkeycode%>&page=<%=i%>"><b>[<font color=red><%=i%></font>]</b></a> 
<%end if
next
if page<pgnum then%>
<a href="search.asp?btnr=<%=btnr%>&zdlb=<%=zdlb%>&nkeycode=<%=nkeycode%>&page=<%=page+1%>">��һҳ</a> 
<%end if%>
</div>
</TD>
</TR>
<TR> 
<TD colspan="4" align=middle>&nbsp;</TD>
</TR>
</TBODY> 
</TABLE>
</td>
<td width="1"></td>
</tr>
</table>
<TABLE WIDTH=770 BORDER=0 CELLPADDING=0 CELLSPACING=0 bgcolor="#B56F00">
    <TR> 
      <TD width="10" bgcolor="#B56F00"></TD>
      </TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE><FONT 
      size=2></FONT></TABLE></DIV>
<DIV align=center>
<TABLE height=70 cellSpacing=0 cellPadding=0 width=775 align=center 
background=index1.files/bottom_t.jpg border=0>
  <TBODY>
  <TR>
    <TD height=70>
      <CENTER>Copyright&copy;2004 7758530.CoM&#8482; All Rights Reserved  
      ���齭����վ<BR>ȫվ����޸�:���齭�����ͷ�QQ:51726805  
TEL:13635716119</CENTER></TD></TR></TBODY></TABLE></DIV></FORM></BODY>
</html>
