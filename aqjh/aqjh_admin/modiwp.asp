<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
wupinid=clng(Request("wupinid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sqlstr="select * from b where id="&wupinid
rs.open sqlstr,conn
if rs.EOF or rs.BOF then
Response.Redirect "wpadmin.asp"
Response.End
else
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href="css/css.css" type=text/css rel=stylesheet>
<title>��Ʒ�޸�</title></head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><font color="#000000"><br>
  ������Ʒ�޸ģ���Ҫע�ⲻҪ���޸ģ�������������ݿ�ĳ���<br>
  �������޸�֮ǰ�����ݣ� </font></p>
<form method="post" action="modiwpok.asp">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="386">
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">ID</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinid" readonly
value="<%=rs("ID")%>" size="5">
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">��Ʒ��</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="a"
value="<%=rs("a")%>" size="20">
        <%if rs("i")<>"��" then%>
        <img src="../hcjs/jhjs/images/<%=rs("i")%>"> 
        <%end if%>
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">ͼƬ</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="i"
value="<%=rs("i")%>" size="10">
        (hcjs/jhjs/images)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">�;ö�</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="j"
value="<%=rs("j")%>" size="10">
        (ֻ��װ����Ч��)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="k"
value="<%=rs("k")%>" size="8">
        (ֻ��<font color="#FFFF00">����</font>��<font color="#FFFF00">����</font>��Ч����������д:<font color="#FFFF00">��</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="l"
value="<%=rs("l")%>" size="20">
        (���⹺����д:<b><font color="#FFFF00">True</font></b>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="b"
value="<%=rs("b")%>" size="10">
        (���������д)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="d"
value="<%=rs("d")%>" size="10">
        (����<font color="#FFFF00">ҩƷ</font>��<font color="#FFFF00">����</font>��<font color="#FFFF00">��ҩ</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="e"
value="<%=rs("e")%>" size="10">
        (����<font color="#FFFF00">ҩƷ</font>��<font color="#FFFF00">����</font>��<font color="#FFFF00">��ҩ</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="f"
value="<%=rs("f")%>" size="10">
        (��������װ����Ч)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="g"
value="<%=rs("g")%>" size="10">
        (��������װ����Ч)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">˵��</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="c"
value="<%=rs("c")%>" size="40">
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="h"
value="<%=rs("h")%>" size="10">
        (<font color="#FFFF00"><b>��Ƭ</b></font>ʱ��д:0 ���ش˿�)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF">���</font></div>
      </td>
      <td width="324"><font color="#FFFFFF"> 
        <input type="text" name="m"
value="<%=rs("m")%>" size="10">
        (<font color="#FFFF00"><b>ֻ�Գ�����Ч</b></font>)</font></td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" value="ȷ ��">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='wpadmin.asp?lb=<%=rs("b")%>'" value="�� ��" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<%
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<div align="center">���ݿ���³�����Ϊʱ�����ޣ�û�м���Ϊ��ֵʱ�ļ�⣬���Ҹ���ʱע��û��ֵ�ĵط���0</div>