<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then 
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻����վ�����㲻�ܲ�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set Conn=Server.CreateObject("ADODB.Connection") 
conn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../chat/questions.asp")
dim rootRs
Set rootRs=Server.CreateObject("ADODB.RecordSet")
sql="select * from questlib"
pagesize=50

rootRs.Open sql,conn,1,1%>
<META content="text/html; charset=gb2312" http-equiv=content-type>
<style type='text/css'>");
body{font-family:"����";font-size:12pt;}td{font-family:"����";font-size:10pt;}
</style>
<body bgcolor="#CCCCCC" background="../jhimg/bk_hc3w.gif" link="#000000" vlink="#000000" alink="#000000">
<table cellspacing=0 cellpadding=0 align=center width=100% height=100%><tr><td valign=center>
<table align=center width=90% cellspacing=0 cellpadding=0 border=1 BORDERCOLORLIGHT=#000000 BORDERCOLORDARK=#ffffff>
 <td colspan=7 class=navy2><center><b>������ȫ</b></td></tr>
 <%If rootRs.Bof OR rootRs.Eof Then%><tr><td colspan=4><center><font color=red>���ݿ���Ŀǰû��׼���κδ��⡣
 <%else
   rootRs.pagesize=pagesize
   page=Request("page")
   if not IsNumeric(page) or page="" then page=1
   if (page-1)<0 then
     page=rootRs.pagecount
   elseif (page-rootRs.pagecount)>0 then
     page=1
   end if
   rootRs.AbsolutePage=page
   RowCount=rootRs.pagesize
   If Not rootRs.Eof Then%>
     <form action="questionlist.asp" method="post">
     <tr><td colspan=7 nowrap><center>
     <font color=990000>��<%=page%>ҳ ��<%=rootRs.pagecount%>ҳ</font>
     <a href="questionlist.asp?page=<%=(page-1)%>">��һҳ</a>
     <a href="questionlist.asp?page=<%=(page+1)%>">��һҳ</a>
     <a href='questionmod.asp'>���</a>
     ת�� <input type="text" name="page" size=6 maxlength=20> ҳ
     <input class=navy1 type="submit" name="cmdGo" value="GO!" title="ת������ҳ">
     </td></tr></form>
     <tr><td class=navy1 nowrap><center>���</td><td class=navy1 nowrap><center>���</td><td class=navy1 nowrap><center>����</td><td class=navy1 nowrap><center>��</td><td class=navy1 nowrap><center>����</td>
     <td class=navy1 nowrap>�޸�</td>
     <%Do While Not rootRs.Eof AND RowCount>0%>
       <tr><td><%=rootRs("ID")%></td><td><%=rootRS("����")%></td><td><%=rootRS("����")%></td><td><%=rootRS("�ش�")%></td><td><%=rootRS("����")%></td>
       <td><a href="questionmod.asp?ID=<%=rootRs("ID")%>">�޸�</a></td>
       <%rootRs.MoveNext
       RowCount=RowCount-1
     Loop%><tr><td class=navy1 nowrap><center>���</td><td class=navy1 nowrap><center>���</td><td class=navy1 nowrap><center>����</td><td class=navy1 nowrap><center>��</td><td class=navy1 nowrap><center>����</td>
     <td class=navy1 nowrap>�޸�</td><%End If
   If rootRs.pagecount>1 then%>
     <tr><td colspan=7><center><b>��ҳ��</b>
     <%For i=1 to rootRs.pagecount
       if (i-page)<>0 then%><a href='questionlist.asp?page=<%=i%>'>[<%=i%>] </a>
       <%else%>(<%=i%>) 
       <%end if
     Next
   end if
 End if
 rootRs.close
 Set rootRs=nothing
 conn.close
 Set conn=nothing%> 
</table><br><center>[<a href='questionmod.asp'>���</a>]  [<a href='javascript:history.go(-1)'>����</a>]</table></body></html>

