<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
chatimage=Application("aqjh_chatimage")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
%><head>
<style type="text/css">
td           { font-family: ����; font-size: 12px }
body         { font-family: ����; font-size: 12px;}
</style>
</head>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<link rel="stylesheet" href="css.css"><meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body bgcolor='#006699' background="<%=chatimage%>" leftmargin="0" topmargin="0" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" background="<%=chatimage%>">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="red" size="2"><strong>���ֶĳ�</strong></font></div></td>
      </tr>
    </table>
	
<table width="140" border="1" align="center" cellpadding="1" cellspacing="0" bordercolorlight="#000000"
bordercolordark="#FFFFFF" bgcolor="#006699" background="<%=chatimage%>">
  <tr>
<td>
	<p align="center"><input type=button value='��������' onClick="window.open('horse.asp','horse','scrollbars=yes,resizable=yes')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button129"></p>
      <p align="center"> 
        <input type=button value='�ľ���Ϣ' onClick="javascript:s('/�Ĳ�$ ��Ϣ')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button129">
      </p>
      <p align="center"> 
        <input type=button value='�ĳ�ׯ����Ϊ'  style="border-style:double; border-width:1.0; background-color:#669900;color:#FFFFFF" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button22">
        <br>
        <input type=button value='�����ľ���ׯ' onClick="javascript:s('/�Ĳ�$ ����ׯ')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
        <input type=button value='���ľ���ׯ' onClick="javascript:s('/�Ĳ�$ ����ׯ')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='�����ľ���ׯ' onClick="javascript:s('/�Ĳ�$ ����ׯ')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='�Ṧ�ľ���ׯ' onClick="javascript:s('/�Ĳ�$ ����ׯ')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center"> 
        <input type=button value='��Ҷľ�ߺ��' onClick="javascript:s('/�Զ�$10')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='���Ҷľ�ߺ��' onClick="javascript:s('/������$10')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='�����ľ�ߺ��' onClick="javascript:s('/�ķ���$100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='����ľ�ߺ��' onClick="javascript:s('/�Ķ�$100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
      </p>
      <p align="center">
        <input type=button value='�ĳ�ɢ����Ϊ'  style="background-color:669900;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button222">
        <br>
        &nbsp;<br>
        <input type=button value='�����ľ�ѺС' onClick="javascript:s('/��ע$ ��&С&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='�����ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&10000')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='�����ľ�Ѻ˫' onClick="javascript:s('/��ע$ ��&˫&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='�����ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      <p align="center">
        <input type=button value='�Ľ����Զľ�' onClick="javascript:s('/�Ľ�����$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button124">
        <input type=button value='��ľ���Զľ�' onClick="javascript:s('/��ľ����$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button123">
      </p>
      <p align="center"> 
        <input type=button value='��ˮ���Զľ�' onClick="javascript:s('/��ˮ����$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button125">
        <input type=button value='�Ļ����Զľ�' onClick="javascript:s('/�Ļ�����$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button126">
      </p>
      <p align="center"> 
        <input type=button value='�������Զľ�' onClick="javascript:s('/��������$100')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button127">
        <input type=button value='�����������' onClick="javascript:s('/����$ ��������1-1000')" style="background-color:#99CE40;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button128">
      <p align="center">
        <input type=button value='�����ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='�����ľ�ѺС' onClick="javascript:s('/��ע$ ��&С&10000')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='�����ľ�Ѻ˫' onClick="javascript:s('/��ע$ ��&˫&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='�����ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&10000')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      </p>
      <p align="center">
        <input type=button value='���ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='���ľ�ѺС' onClick="javascript:s('/��ע$ ��&С&100')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='���ľ�Ѻ˫' onClick="javascript:s('/��ע$ ��&˫&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='���ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      </p>
      <p align="center">
        <input type=button value='�����ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='�����ľ�ѺС' onClick="javascript:s('/��ע$ ��&С&100')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='�����ľ�Ѻ˫' onClick="javascript:s('/��ע$ ��&˫&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='�����ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      </p>
      <p align="center">
        <input type=button value='�Ṧ�ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button23">
        <input type=button value='�Ṧ�ľ�ѺС' onClick="javascript:s('/��ע$ ��&С&100')" style="background-color:#A938A0;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2233">
      </p>
      <p align="center">
        <input type=button value='�Ṧ�ľ�Ѻ˫' onClick="javascript:s('/��ע$ ��&˫&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2204">
        <input type=button value='�Ṧ�ľ�Ѻ��' onClick="javascript:s('/��ע$ ��&��&100')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214">
      </p>
      <p align="center">
        <input type=button value='˫�˶Ĳ�' style="background-color:red;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button22">
      </p>
      <p align="center"> 
        <input type=button value='ʯͷ' onClick="javascript:s('/˫�˶Ĳ�$ 1')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214" width=30>
        <input type=button value='����' onClick="javascript:s('/˫�˶Ĳ�$ 2')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214" width=30>
        <input type=button value=' �� ' onClick="javascript:s('/˫�˶Ĳ�$ 3')" style="background-color:#A938A0;color:#FFffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2214" width=30>
		<br>
      </p>
</td>
</tr>
</table>
</body>