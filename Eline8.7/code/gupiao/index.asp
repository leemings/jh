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
%>
<title>��Ʊ�������ġ�wWw.51eline.com��</title>
<script language=javascript>function selectstock(stock){parent.stockfrm.location.replace('showstock.asp?stock='+stock);parent.infofrm.location.replace('info.asp?stock='+stock);}setTimeout('location.reload();',180000);</script>
<link rel="stylesheet" href="new.CSS" type="text/css">
<body link="#0000FF" vlink="#0000FF" alink="#0000FF" bgcolor="#339966">
<a href="../welcome.asp" target="_parent">��E�߽�����</a>-��Ʊ���� 
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr> 
<td> 
<table width="100%" border="0" cellspacing="0" bgcolor="#000000" cellpadding="0">
<tr bgcolor="whitesmoke"> 
<td width="10%" height="25"> 
<div align="center">��Ʊ����</div>
</td>
<td width="20%" height="25"> 
<div align="center">��Ʊ����</div>
</td>
<td width="20%" height="25"> 
<div align="center">��ͨ������</div>
</td>
<td width="10%" height="25"> 
<div align="center">���м�</div>
</td>
<td width="10%" height="25"> 
<div align="center">��ǰ�۸�</div>
</td>
<td width="10%" height="25"> 
<div align="center">��ʷ�߼�</div>
</td>
<td width="10%" height="25"> 
<div align="center">��ʷ�ͼ�</div>
</td>
<td width="10%" height="25"> 
<div align="center">����</div>
</td>
</tr>
</table>
<%
set conn=server.CreateObject("adodb.connection")
set rs=server.CreateObject("adodb.recordset")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("stock.mdb")
rs.Open "select * from ��Ʊ",conn,2,2
do while not(rs.EOF or rs.BOF)
increa =(rs("�ּ�")/rs("����ɽ���")-1)*100
if increa>0 then 
increa ="<font color=0000FF>"&formatnumber(increa,2,-1)&"</font>"
elseif increa<0 then
increa ="<font color=FF0000>"&formatnumber(increa,2,-1)&"</font>"
else
increa ="<font color=000000>"&formatnumber(increa,2,-1)&"</font>"
end if 
%>
      <table width="100%" border="0" cellpadding="2" cellspacing="1">
        <tr bgcolor="#339966" onmouseout="this.bgColor='#339966';"onmouseover="this.bgColor='#DFEFFF';"> 
<td width="10%"><div align="center"><%=rs("id")%> </div> 
</td>
<td width="20%"> 
            <div align="center"> <a href="#" onClick="selectstock('<%=rs("��Ʊ����")%>');"><%=rs("��Ʊ����")%></a></div>
</td>
<td width="20%"><div align="center"><%=rs("��ͨ������")%> 
</div>
</td>
<td width="10%"><div align="center"><%=rs("���м�")%> 
</div>
</td>
<td width="10%"><div align="center"><%=rs("�ּ�")%> 
</div>
</td>
<td width="10%"><div align="center"><%=rs("��ͨ������")%> 
</div>
</td>
<td width="10%"><div align="center"><%=rs("��ʷ�ͼ�")%> 
</div>
</td>
<td width="10%"><div align="center"><%=increa%> 
</div>
</td>
</tr>
</table>
<%
rs.MoveNext
loop
%>
</td>
</tr>
</table>
<% 
rs.close
conn.close
set rs=nothing
set conn=nothing
%>