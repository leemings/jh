<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
if sjjh_name<>Application("sjjh_user") then Response.Redirect "../chat/readonly/bomb.htm"
wupinid=clng(Request("wupinid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
sqlstr="select * from r where id="&wupinid
rs.open sqlstr,conn
if rs.EOF or rs.BOF then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��û����Ҫ�����ķ��䣬��ˢ�£�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
<title>�޸ķ������ϡ�wWw.happyjh.com��</title></head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font color="#000000" size="2">��н��������޸ĳ���</font> </div>
<form method="post" action="modiroomok.asp">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="75%">
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">ID</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> <%=rs("id")%> </font> 
        <input type="hidden" name="id" value="<%=rs("id")%>" size="15" maxlength="20">
      </td>
      <td width="81"> 
        <div align="center"><font size="2" color="#FFFFFF">������</font></div>
      </td>
      <td width="224"><font color="#FFFFFF" size="2"> 
        <input type="text" name="name"
value="<%=rs("a")%>" size="15" maxlength="20">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">�������</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <input type="text" name="zdrs"
value="<%=rs("b")%>" size="15" maxlength="3">
        </font></td>
      <td width="81" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">��������</font></div>
      </td>
      <td width="224"><font color="#FFFFFF" size="2"> 
        <select name=fzxz size=1 >
          <option value='0' <%if rs("f")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("f")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">�¼�����</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <select name=sjxz size=1 >
          <option value='0' <%if rs("g")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("g")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </font></td>
      <td width="81"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td width="224"><font color="#FFFFFF" size="2"> 
        <select name=bh size=1 >
          <option value='0' <%if rs("h")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("h")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">��Ƭ</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <select name=kp size=1 >
          <option value='0' <%if rs("i")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("i")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </font></td>
      <td width="81"> 
        <div align="center"><font size="2" color="#FFFFFF">�Ĳ�</font></div>
      </td>
      <td width="224"><font color="#FFFFFF" size="2"> 
        <select name=db size=1 >
          <option value='0' <%if rs("j")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("j")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">����</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <select name=xz size=1 >
          <option value='0' <%if rs("c")=0 then%>selected<%end if%>>����[����˷���������]</option>
          <option value='1' <%if rs("c")=1 then%>selected<%end if%>>��ֹ[����˷���������]</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">����˵��</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <input type="text" name="xzsm"
value="<%=rs("d")%>" size="50">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font size="2" color="#FFFFFF">���ʽ</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <input type="text" name="bds"
value="<%=rs("e")%>" size="50">
        <br>
        ����ʹ�ã�and��or��not��+��-��=�ȱ��ʽ���Ӧ�������ӣ�<br>
        ���ڱ�̲��˽��߲��������޸ģ�(IDΪ1�����첻Ҫ�����ƣ�) </font></td>
    </tr>
    <tr> 
      <td width="67" height="2"> 
        <div align="center"><font size="2" color="#FFFFFF">���ջ���</font></div>
      </td>
      <td colspan="3" height="2"><font color="#FFFFFF" size="2"> 
        <input type="text" name="jrht"
value="<%=rs("k")%>" size="50" maxlength="150">
        (�ڽ��뽭��ʱ�ǻ���ʾ��)</font></td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"> 
          <input type="submit" value="ȷ ��" name="submit">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='roomlist.asp'" value="�� ��" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<%
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<div align="center"> <font color="#FFFFFF" size="2">���ݿ���³�����Ϊʱ�����ޣ�û�м���Ϊ��ֵʱ�ļ�⣬���Ҹ���ʱע��û��ֵ�ĵط���0</font></div>