<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
wupinid=clng(Request("wupinid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
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
<LINK href=css/css.css type=text/css rel=stylesheet>
<title>�޸ķ�������</title></head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center"><font color="#000000">�����޸ĳ��� </div>
<form method="post" action="modiroomok.asp"><center>
  <table border="0" cellspacing="1" cellpadding="4" bgcolor="f2f2ea" cellspacing=0 cellpadding=0  width="75%">
    <tr> 
      <td width="67"> 
        <div align="center">ID</div>
      </td>
      <td width="197"> <%=rs("id")%>  
        <input type="hidden" name="id" value="<%=rs("id")%>" size="15" maxlength="20">
      </td>
      <td width="81"> 
        <div align="center">������</div>
      </td>
      <td width="224"> 
        <input type="text" name="name"
value="<%=rs("a")%>" size="15" maxlength="20" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">�������</div>
      </td>
      <td width="197"> 
        <input type="text" name="zdrs"
value="<%=rs("b")%>" size="15" maxlength="3" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
      <td width="81" nowrap> 
        <div align="center">��������</div>
      </td>
      <td width="224"> 
        <select name=fzxz size=1 >
          <option value='0' <%if rs("f")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("f")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </td>
    </tr>
 <tr> 
      <td width="67"> 
        <div align="center">���ʱ��</div>
      </td>
      <td width="197"> 
        <input type="text" name="djsj"
value="<%=rs("l")%>" size="2" maxlength="2" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid"> [����60Ϊ������]
        </td>
      <td width="81" nowrap> 
        <div align="center">�ݵ�</div>
      </td>
      <td width="224"> 
        <select name=pd size=1 >
          <option value='1' <%if rs("m")=1 then%>selected<%end if%>>����</option>
          <option value='0' <%if rs("m")=0 then%>selected<%end if%>>��ֹ</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center">�¼�����</div>
      </td>
      <td width="197"> 
        <select name=sjxz size=1 >
          <option value='0' <%if rs("g")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("g")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </td>
      <td width="81"> 
        <div align="center">����</div>
      </td>
      <td width="224"> 
        <select name=bh size=1 >
          <option value='0' <%if rs("h")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("h")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">��Ƭ</div>
      </td>
      <td width="197"> 
        <select name=kp size=1 >
          <option value='0' <%if rs("i")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("i")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </td>
      <td width="81"> 
        <div align="center">�Ĳ�</div>
      </td>
      <td width="224"> 
        <select name=db size=1 >
          <option value='0' <%if rs("j")=0 then%>selected<%end if%>>����</option>
          <option value='1' <%if rs("j")=1 then%>selected<%end if%>>��ֹ</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">����</div>
      </td>
      <td colspan="3"> 
        <select name=xz size=1 >
          <option value='0' <%if rs("c")=0 then%>selected<%end if%>>����[����˷���������]</option>
          <option value='1' <%if rs("c")=1 then%>selected<%end if%>>��ֹ[����˷���������]</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">����˵��</div>
      </td>
      <td colspan="3"> 
        <input type="text" name="xzsm"
value="<%=rs("d")%>" size="50" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">���ʽ</div>
      </td>
      <td colspan="3"> 
        <input type="text" name="bds"
value="<%=rs("e")%>" size="50" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
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
value="<%=rs("k")%>" size="50" maxlength="150" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        (�ڽ��뽭��ʱ�ǻ���ʾ��)</td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"> 
          <input type="submit" value="ȷ ��" name="submit">
          <font color="#CCCCCC">-------  
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
<div align="center"> ���ݿ���³�����Ϊʱ�����ޣ�û�м���Ϊ��ֵʱ�ļ�⣬���Ҹ���ʱע��û��ֵ�ĵط���0</div>