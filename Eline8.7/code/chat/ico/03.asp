<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if chatbgcolor="" then chatbgcolor="008888"%>

<html>

<head>
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:ffffff;text-decoration:none;}
a:hover{color:ffffff;text-decoration:none;}
</style>
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�¹���˵��</title>
</head>

<body bgcolor="#000000" text="#FFFFCC">

<table border="1" width="100%" bordercolor="#000000" bordercolorlight="#FFFFFF" bordercolordark="#000000">
  <tr>
    <td width="100%"><img border="0" src="ki17.gif" align="absmiddle" width="18" height="18">�¹���˵��-������ƽ�����Ϸ��-����ħ��</td>
  </tr>
  <tr>
    <td width="100%">���ӹ�����Ŀ</td>
  </tr>
  <tr>
    <td width="100%">����----���ͷ���.��ŷ���.ȡ�ط���.Ѱ��ˮ��.Ѱ�ҷ���.��������.�Ȼ��˼�.��ʩ����.ħ������.û�շ���.�Ի��.��������.ħ��ˮ��.������.ʥ����.������.��ڤ��.����ȭ.��ȡ��.��������</td>
  </tr>
  <tr>
    <td width="100%">�Ṧ----�����Ṧ.�Ṧ�ݴ�.��ȡ�Ṧ.���߹���.Ѱ������.��ȡ�Ṧ.</td>
  </tr>
  <tr>
    <td width="100%">����----������.�ߵ�����.������.��ť��.������.������ť.���°�ť.</td>
  </tr>
  <tr>
    <td width="100%">ְҵ----ħ��ʦ.��ʦ.���.����.����ʦ.����ʦ.</td>
  </tr>
  <tr>
    <td width="100%">����սʿ</td>
  </tr>
  <tr>
    <td width="100%">��ʬ��</td>
  </tr>
  <tr>
    <td width="100%">ħ��ʦ��ִ�����¹���</td>
  </tr>
  <tr>
    <td width="100%">Ѱ�ұ���.�����.�ⶾ��.ƽ����.͵����.�����.˧����.������.ҡǮ��.���黷.������.</td>
  </tr>
  <tr>
    <td width="100%" align="center">����᲻���ṩ��ѵĹ��ܸ���λ</td>
  </tr>
  <tr>
    <td width="100%" align="center"><a href="http://zhzx.jjedu.org/eline/" target="_blank">���ֽ���</a></td>
  </tr>
</table>

</body>

</html>
<Script >
parent.msgfrm.rows = '*,*,*';
parent.tbclu=true;
</script>