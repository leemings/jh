<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"%>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
<title>��E�߽��������ӹ�Ʊ</title></head>
<body text="#000000" background="../../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"> 
<p><font size="6" face="����">������Ʊ����</font><br>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../gupiao/stock.mdb")
sql="select * from ��Ʊ"
set rs=conn.execute(sql)
do while Not rs.Eof
%>
<form name="form1" method="post" action="gupiaook.asp">
    <div align="left"><font size="-1">��Ʊ<input type="text" name="gupiao" readonly size="10" maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=rs("��Ʊ����")%>">
      ��ͨ����<input type="text" name="zl" size="5" maxlength="5" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=rs("��ͨ������")%>">
      ʣ���<input type="text" name="sy" size="5" maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=rs("ʣ��ɷ�")%>">
      ���м�<input type="text" name="fx" size="5" maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=rs("���м�")%>">
      ��ʷ�߼�<input type="text" name="gj" size="5" maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=rs("��ʷ�߼�")%>">
      ��ʷ�ͼ�<input type="text" name="dj" size="5" maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=rs("��ʷ�ͼ�")%>">
      ����ɽ���<input type="text" name="cjj" size="5" maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=rs("����ɽ���")%>">
      �ּ�<input type="text" name="xj" size="5" maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="<%=rs("�ּ�")%>">
      <input type=submit value='����' name="submit">
      <input type=submit value='ɾ��' name="submit1">
      </font> </div>
  </form>
    <%
rs.MoveNext
loop
rs.close 
conn.close
set rs=nothing
set conn=nothing 
%>
