<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../chat/readonly/bomb.htm"
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href=css/css.css type=text/css rel=stylesheet>
<title>��������</title></head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center"><font color="#000000">������������</div>
<form method="post" action="addroomok.asp"><center>
  <table border="0" cellspacing="1" cellpadding="4" bgcolor="f2f2ea" cellspacing=0 cellpadding=0  width="75%">
    <tr> 
      <td width="67"> 
        <div align="center">����</div>
      </td>
      <td width="197">������������</td>
      <td width="81"> 
        <div align="center">������</div>
      </td>
      <td width="224">
        <input type="text" name="name" size="15" maxlength="20" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">�������</div>
      </td>
      <td width="197">
        <input type="text" name="zdrs" size="15" maxlength="3" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
      <td width="81" nowrap> 
        <div align="center">��������</div>
      </td>
      <td width="224">
        <select name=fzxz size=1 >
          <option value='0' selected >����</option>
          <option value='1' >��ֹ</option>
        </select>
        </td>
    </tr>
 <tr> 
      <td width="67"> 
        <div align="center">���ʱ��</div>
      </td>
      <td width="197"> 
        <input type="text" name="djsj"
value="60" size="2" maxlength="2" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid"> [����60Ϊ������]
        </td>
      <td width="81" nowrap> 
        <div align="center">�ݵ�</div>
      </td>
      <td width="224"> 
        <select name=pd size=1 >
          <option value='1' selected>����</option>
          <option value='0'>��ֹ</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center">�¼�����</div>
      </td>
      <td width="197">
        <select name=sjxz size=1 >
          <option value='0' selected >����</option>
          <option value='1' >��ֹ</option>
        </select>
        </td>
      <td width="81"> 
        <div align="center">����</div>
      </td>
      <td width="224">
        <select name=bh size=1 >
          <option value='0' selected>����</option>
          <option value='1'>��ֹ</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">��Ƭ</div>
      </td>
      <td width="197">
        <select name=kp size=1 >
          <option value='0' selected >����</option>
          <option value='1' >��ֹ</option>
        </select>
        </td>
      <td width="81"> 
        <div align="center">�Ĳ�</div>
      </td>
      <td width="224">
        <select name=db size=1 >
          <option value='0' selected>����</option>
          <option value='1' >��ֹ</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">����</div>
      </td>
      <td colspan="3">
        <select name=xz size=1>
          <option value='0' selected >����[����˷���������]</option>
          <option value='1' >��ֹ[����˷���������]</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">����˵��</div>
      </td>
      <td colspan="3">
        <input type="text" name="xzsm" size="50" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">���ʽ</div>
      </td>
      <td colspan="3">
        <input type="text" name="bds" size="50" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        <br>
        ����ʹ�ã�and��or��not��+��-��=�ȱ��ʽ���Ӧ�������ӣ�<br>
        ���ڱ�̲��˽��߲��������޸ģ�(IDΪ1�����첻Ҫ�����ƣ�) </td>
    </tr>
    <tr> 
      <td width="67" height="2"> 
        <div align="center">���ջ���</div>
      </td>
      <td colspan="3" height="2">
        <input type="text" name="jrht"
value="����" size="50" maxlength="150" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        (�ڽ��뽭��ʱ�ǻ���ʾ��)</td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"> 
          <input type="submit" value="ȷ ��">
          <font color="#CCCCCC">-------  
          <input  onClick="javascript:window.document.location.href='roomlist.asp'" value="�� ��" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<div align="center">���ݿ���³�����Ϊʱ�����ޣ�û�м���Ϊ��ֵʱ�ļ�⣬���Ҹ���ʱע��û��ֵ�ĵط���0</div>