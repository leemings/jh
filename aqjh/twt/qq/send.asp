<html><head>
<title>������������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../css.css" rel=stylesheet>
<SCRIPT language=JavaScript>
<!--
function FormCheck(){
        if (document.fm1.oicq.value == "")
        {
        alert("����дQQ���룡");
		document.fm1.oicq.focus();
        return false;
        }
//�ж����������Ƿ�λ����
        var filter=/[.0-9]/;
        if (!filter.test(document.fm1.oicq.value)) { 
                alert("QQ����ֻ��ʹ�����֣�"); 
                document.fm1.oicq.focus();
                document.fm1.oicq.select();
                return false; 
                } 
		return true;
}
//-->
</SCRIPT>
</head>
<body bgcolor="#FFFFFF">
<form action="sendsave.asp" onSubmit="return FormCheck();" name="fm1" method="post">
  <table width="100%" border="0" cellspacing="0" cellpadding="5">
    <tr bgcolor="#CCCCCC"> 
      <td colspan="2">�� <a href="qqlist.asp">�������б�</a> �� <a href="qqlist.asp?id=1">��������</a> �� <a href="qqlist.asp?id=2">��������</a> �� <a href="qqlist.asp?id=3">��δ��������</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%if request("id")=2 then%><font color="#FF0000">����ȷ��дQQ���룡</font><%end if%></td>
    </tr>
    <tr> 
      <td colspan="2" align="center">���齭��QQ�����������뵥</td>
    </tr>
    <tr> 
      <td colspan="2" align="center"> 
        <hr size="1" noshade>
      </td>
    </tr>
    <tr> 
      <td align="right">����QQ�ţ�</td>
      <td> 
        <input type="text" name="oicq" maxlength="12">
        <font color="#FF0000">*</font> ע���������Ѿ��޸������Ϻ��QQ��</td>
    </tr>
    <tr> 
      <td colspan="2" align="center"> 
        <input type="submit" name="Submit" value="�ύ">
        <input type="reset" name="Submit2" value="����">
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <hr size="1" noshade>
      </td>
    </tr>
    <tr> 
      <td colspan="2"> 
        <p><b><font color="#FF0000">�޸ķ�����</font></b></p>
        <p><span class="qq">1��������QQ��<b>������ϸ����</b>�ĸ�����ҳ��Ŀ�޸�Ϊ��<font color="#FF0000"><b><%=Application("aqjh_homepage")%></b></font><br>
          2��������<b>����˵��</b>��Ŀ����:<%=Application("aqjh_chatroomname")%> <font color="#FF0000"><b><%=Application("aqjh_homepage")%></b></font> 
          ��Ȼ�˼�һЩ�����Ե�˵���͸����ˣ����磺�ƽ�Ľ���,���������˵� <br>
          3�� ������ʩ�������޸���QQ���ϲ�����һ�������ϵģ��ͽ�������2000������200���书2�� </span></p>
        <p><span class="qq">������ȷ�Ϻ�ϵͳ���Զ����ڽ����� �����ϵͳ���Զ���¼����ҿ��Ի���ල������ĳЩ���������뼰ʱ��٣��һ������Ӧ�Ĵ�����</span></p>
        <p><span class="qq"><font color="#FF0000">ע��</font>����ʵ��д����QQ�ţ��������������������<font color="#FF0033"><b>ɾ���˺�</b></font>�����ָ�������֧�ֽ�����չ���������������������ǻ���轱���������޸���QQ���ϲ�����һ�������ϵģ��ͽ�������2000������200���书2�� 
          ����2���������߿������Ը�վ�����������Ľ�����������м���QQ�Ż��������ѵĵ�QQҲ��������[������5������]��״̬�����ظ����ӣ���������������޿�����������״ֵ̬���潱���� 
          </span></p>
        <p><span class="qq"><font color="#FF0000">����</font>����������д�����ǲ�ȡ����ȴ�ʩ��������ɾ���˺Ż����IP���������뽱���󱣳ֲ���һ����������Ҳ���ȡ��Ӧ�Ĵ�����ʩ�ջ������ӵ�״̬��</span></p>
      </td></tr></table></form></body></html>