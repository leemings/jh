<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
lb=LCase(trim(Request.QueryString("lb")))
if lb="" then lb="ҩƷ"
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href=css/css.css type=text/css rel=stylesheet>
<title>�����Ʒ</title></head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<p align="center"><font size="2" color="#000000">�����Ʒ</font></p>
<form method="post" action="addwpok.asp">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="386">
    <tr> 
      <td colspan="2" height="18"><font color="#FFFF00" size="-1">��ʾ:������ӿ�Ƭ����Ҫ��chat/sjfunc/44.asp�ļ�!</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">��Ʒ��</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="a" size="20"
value="<%=lb%>">
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">ͼƬ</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="i"
value="��" size="10">
        (hcjs/jhjs/images)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">�;ö�</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="j"
value="0" size="10">
        (ֻ��װ����Ч��)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="k"
value="��" size="8">
        (ֻ��<font color="#FFFF00">����</font>��<font color="#FFFF00">����</font>��Ч����������д:<font color="#FFFF00">��</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font size="2" color="#FFFFFF">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="l" size="20"
value="True">
        (���⹺����д:<b><font color="#FFFF00">True</font></b>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">���</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="b"
value="<%=lb%>" size="10">
        (���������д)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="d"
value="0" size="10">
        (����<font color="#FFFF00">ҩƷ</font>��<font color="#FFFF00">����</font>��<font color="#FFFF00">��ҩ</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="e"
value="0" size="10">
        (����<font color="#FFFF00">ҩƷ</font>��<font color="#FFFF00">����</font>��<font color="#FFFF00">��ҩ</font>)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="f"
value="0" size="10">
        (��������װ����Ч)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="g"
value="0" size="10">
        (��������װ����Ч)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font size="2" color="#FFFFFF">˵��</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="c" size="40"
value="<%=lb%>">
        </font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="h"
value="0" size="10">
        (<font color="#FFFF00"><b>��Ƭ</b></font>ʱ��д:0 ���ش˿�)</font></td>
    </tr>
    <tr> 
      <td width="62"> 
        <div align="center"><font size="2" color="#FFFFFF">���</font></div>
      </td>
      <td width="324"><font color="#FFFFFF" size="2"> 
        <input type="text" name="m"
value="0" size="10">
        (<font color="#FFFF00"><b>ֻ�Գ�����Ч</b></font>)</font></td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" value="ȷ ��">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='wpadmin.asp?lb=<%=lb%>'" value="�� ��" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
