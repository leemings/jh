<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<TITLE>�����������</TITLE>
<link rel="stylesheet" href="../../css.css">
</HEAD>
<BODY bgcolor="#FFFFFF" background="../jhimg/bk_hc3w.gif" topmargin=0>
<!--#include file="head.asp"-->
<table width=500 border="0" align=center cellspacing="0" bordercolor="#FF0000">
        <tr bordercolor="#FFFFFF">
          <td align=center><font color=red><b>�� �� �� �� �� �� �� �� �� �� ��</b></font><BR><BR></td>
        </tr>
      </table>
<!--#include file="data.asp"-->
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "Select * from ��Ȧ����",connt,3,3
jz=rs("ͶƱ����")
if cdate(jz)>now() then
 Response.Write "<center>��ʾ��ͶƱû�������в��ܿ�����"
 Response.End
end if
%>
      <table width=650 border="1" align=center cellspacing="0" bordercolorlight="#800080"
bordercolordark="#FFFFFF">
        <tr bgcolor="#DFEFFF">
          <td width=90 align=center>������</td>
          <td width=90 align=center>��������</td>
          <td width=270 align=center>�� �� ˵ ��</td>
          <td width=100 align=center>��������</td>
          <td width=40 align=center>����</td><td width=40 align=center>�� Ʒ</td>
<%
rs.close
sql="select * from ������ư� where ����>0 order by ����"
rs.open sql,connt,1,1
if rs.recordcount<>0 then
  rs.pagesize = 25
  page = cint(request("pageno"))
     if page <= 1 then 
        page = 1
     end if
     if page >= rs.pagecount then
        page = rs.pagecount
     end if
     rs.absolutepage = page
  for i=1 to rs.pagesize
%>
        <tr bgcolor="" onmouseout="this.bgColor='';"onmouseover="this.bgColor='#DFEFFF';">
          <td width=90 align=center><font color=red><%=rs("������")%></font></td>
          <td width=90 align=center><a href=list.asp?myid=<%=rs("������")%> target=_blank><font color=blue><%=rs("��������")%></a></td>
          <td width=270 align=center><%=rs("����˵��")%></td>
          <td width=100 align=center><%=rs("��������")%></td>
          <td width=40 align=center>��<font color=red><%=rs("����")%></font>��</td><td width=40 align=center><%if rs("�콱")=true then%>���콱<%else%><a href=hfybok.asp?myid=<%=rs("������")%> target=_blank>�콱</a><%end if%></td>
        </tr>
<% 
  rs.movenext
  if rs.eof then exit for
  next
 %>   
��<font color=2020dd><% =rs.pagecount %></font>ҳ<font color=2020dd><% =rs.recordcount %></font>�� ��ǰ:<font color=2020dd><%=page%></font>/<%=rs.pagecount%>&nbsp;&nbsp;<a href=hfyb.asp?pageno=1>��ҳ</a>&nbsp;<a href=hfyb.asp?pageno=<% =page-1 %>>��ҳ</a>&nbsp;<a href=hfyb.asp?pageno=<% =page+1 %>>��ҳ</a>&nbsp;<a href=hfyb.asp?pageno=<% =rs.pagecount %>>ĩҳ</a>
<%end if%>
</table><br>
<center>
<font color=red>ע�⣺�콱ǰ���Ƚ���ɱ������ϵͳ���Զ�����������̣��Ա���һ����������ľ��У�</font><br><br>
<FONT color=#0000ff>&copy; ��Ȩ���� 2015-2015 </FONT><A href="http://www.happyjh.com/" target=_blank><FONT color=#0000ff>���ֽ�����</FONT></A>
</body></html>