<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if aqjh_grade<>10 or aqjh_name<>application("aqjh_user") then Response.Redirect "../../error.asp?id=439"
page=request.querystring("page")
sfsz=request.querystring("action")
fs=""
jg=""
if sfsz="L" then
	fs=lcase(trim(request.form("fs")))
	jg=trim(request.form("jg"))
end if
PageSize=30
mdbsj="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../gg.asa")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open mdbsj
if sfsz="L" and jg<>"" then
	if fs="id" then
		cx=fs&"="&jg
	else
		cx="instr("&fs&",'"&jg&"')<>0"
	end if
	sql="select * from ���� where "&cx
else
	sql="select * from ���� order by ʱ�� desc"
end if
'response.write sql
'response.end
set rs=conn.execute(sql)
rs.PageSize = PageSize
pgnum=rs.pagecount
if pgnum<1 then pgnum=1
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if rs.pagecount>0 then rs.AbsolutePage=page
count=0
%>
<html>
<head>
<title>��ҳվ���������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!-- 
a { text-decoration: none} 

--> 
A:link {COLOR: #0000FF;FONT-FAMILY: ����; }
A:visited {COLOR: #0000FF; FONT-FAMILY: ����; }
A:active {FONT-FAMILY: ����; }
A:hover {FONT-FAMILY: ����;COLOR: #FF0000; }
BODY {FONT-FAMILY: ����; FONT-SIZE: 9pt;COLOR: #000000;}
TABLE {FONT-FAMILY: ����; FONT-SIZE: 9pt;COLOR: #000000;}
</style>
</head>
<body bgcolor="#FFFFFF" background="../../jhimg/bk_hc3w.gif" style="font-size: 12px">
<p align="center"><font color="#0000FF">��ҳվ���������</font></p>
<p align="center"><a href="gggl.asp">������</a><font color="#0000FF">&nbsp;</font></p>
<p></p>
<table border="1" width="648" bordercolorlight="#000000" cellspacing="1" cellpadding="1" bordercolordark="#85C2E0" height="25" align="center">
  <tr> 
    <td align="center" nowrap height="25" width="5%">ID��</td>
    <td align="center" nowrap height="25" width="44%">����</td>
    <td align="center" nowrap height="25" width="14%">����ʱ��</td>
    <td align="center" nowrap height="25" width="14%">�������</td>
    <td align="center" nowrap height="25" width="11%">�������</td>
    <td align="center" nowrap height="25" width="12%"><a href=addgg.asp>����¹���</a></td>
  </tr>
  <%
if rs.eof or rs.bof then
%>
  <tr> 
    <td bgcolor="#85C2E0" onMouseOut="this.bgColor='#85C2E0';"onMouseOver="this.bgColor='#DFEFFF';" align="center" nowrap height="25" colspan="6"><font color="ff0000" size="2">��û������κι���!</font></td>
  </tr>
  <%response.end
end if
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="#85C2E0" onMouseOut="this.bgColor='#85C2E0';"onMouseOver="this.bgColor='#DFEFFF';" font size="2";> 
    <td  align="center" nowrap height="25" width="5%"><%=rs("id")%></td>
    <td nowrap height="25" width="44%"><a href="../disp.asp?id=<%=rs("id")%>" target=_blank><%=rs("����")%></a></td>
    <td  align="center" nowrap height="25" width="28%"><%=rs("ʱ��")%></td>
	<td  align="center" nowrap height="25" width="28%"><%=rs("�������")%></td>
    <td  align="center" nowrap height="25" width="11%"><%=rs("�鿴����")%></td>
    <td  align="center" nowrap height="25" width="12%"><a href="modigg.asp?id=<%=rs("id")%>">�޸�</a> 
      <a href="delgg.asp?id=<%=rs("id")%>" title='ɾ���˹���'>ɾ��</a></td>
  </tr>
  <%
	rs.movenext
loop 
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
  <tr>
    <form name="form1" method="post" action="ggadmin.asp?action=L">
      <td colspan="6"> 
        <div align="right"> 
          <select name="fs">
            <option value="����" selected>����</option>
            <option value="id">�ɣ�</option>
            <option value="�鿴����">���</option>
            <option value="�������">���</option>
          </select>
          <input type="text" name="jg">
          <input type="submit" name="Submit" value="����">
        </div>
      </td>
    </form>
  </tr>
</table>
<p align="center">��ǰ�� <%=page%>/<%=pgnum%> ҳ��ÿҳ <%=pagesize%> �� 
  <%ys=1
		  if page>pgnum then%>
  <a href="ggadmin.asp?page=<%=page-1%>"><br>
  ��һҳ</a> 
  <%end if
		  do while ys<pgnum
		  	if page=ys then%>
  <B>[<%=ys%>]</B> 
  <%else%>
  <A class=black href="ggadmin.asp?page=<%=ys%>">[<%=ys%>]</A> 
  <%end if
			ys=ys+1
		  loop
		  if page<pgnum then%>
  <a href="list.asp?page=<%=page+1%>">��һҳ</A> 
  <%end if%>
</p>
<p align="center">������������������д</p>
</body>
</html>