<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
%><html>
<head>
<title>������ʾ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="chat/readonly/style.css">
<style type="text/css">
body{
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff}
</style>
<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<body background=images/bg.gif leftmargin="0" topmargin="0" onLoad="MM_preloadImages('jhimg/err_close2.gif','jhimg/err_but2.gif')">
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<table align=center background=jhimg/err_bg1.gif border=0 
      cellpadding=0 cellspacing=0 class=body width=400>
  <tbody> 
  <tr> 
    <td height=23 width=25><img border=0 height=23 
            src="jhimg/err1.gif" width=24></td>
    <td height=23 width=340>&nbsp;<font color=#000000 
            face="Arial, Helvetica, sans-serif">ERROR - ��������</font></td>
    <td align=right height=23 valign=baseline width="35"><a 
            href="javascript:window.close()" onMouseOut=MM_swapImgRestore() 
            onMouseOver="MM_swapImage('close','','jhimg/err_close2.gif',1)"><img 
            border=0 height=18 name=close src="jhimg/err_close1.gif" 
            width=15></a><img border=0 height=23 name=errr1_c4 
            src="jhimg/err2.gif" width=5></td>
  </tr>
  </tbody> 
</table>
<table border="0" width="400" cellpadding="0" cellspacing="0" align="center" background="JHIMG/err_bg.gif">
  <tr> 
    <td width="98" align="center" valign="top"><img src="Error.gif" width="100" height="92"></td>
    <td width="302"> 
      <%Select Case id
Case "000"%>
      <p>�����Բ��𣬳�������Ŀ¼��������Ŀ¼���������д���Global.asa û�б�ִ�С���������Ҫ����Ŀ¼��֧�֣����������Ŀ¼�뽫global.asa���Ƶ���Ŀ¼��ȥ��</p>
      <%Case "002"%>
      <p>�����Բ��𣬱�����������̨�����������У�������Ǳ�����ĺϷ��û�����������������ϵ��<br>
        ���䣺<a href="mailto:yxt_email@etang.com">yxt_email@etang.com</a><br>
      </p>
      <%Case "003"%>
      <p>�������ݿ���δ�򿪣�������վ����������ѹ�����ݿ⣬���Ժ�������</p>
      <%Case "010"%>
      <p>�����Բ��𣬱�����ר����� Internet Explorer 4.0 ���ϰ汾���������������ʹ�� IE ����������ʡ�</p>
      <%Case "011"%>
      <p>�����Բ�������ͨ�����������������ʱ�վ��</p>
      <%Case "50"%>
      <p>��������oicq��ʽ����ȷ��oicqΪ��5-10λ�����֣�</p>
      <%Case "51"%>
      <p>��������Email��ʽ����ȷ��EmailΪ�һ������ã��ǳ���Ҫ��</p>
      <%Case "52"%>
      <p>����Ϊ�����İ�ȫ��������ʹ��6λ���ϵ��ַ�����������ĸ�����֡�</p>
      <%Case "53"%>
      <p>�������������������ݣ�</p>
      <%Case "54"%>
      <p>����ϵͳ��ֹ�ˡ�'��,��,��,��"��,��or���ȶ�ϵͳ���ƻ��Ե��ַ���</p>
      <%Case "55"%>
      <p>����ϵͳΪ�˷�ֹ�ڿ�ʹ��̽������������60����ٲ������û�����</p>
      <%Case "56"%>
      <p>�����û�����oicq�ţ�Email�����룬��ʾ������Ϊ�գ�</p>
      <%Case "57"%>
      <p>���������˵��������������Լ���������ͬ��</p>
      <%Case "58"%>
      <p>���������˲������ڣ��������</p>
      <%Case "59"%>
      <p>����ע�����������Ϊͬһip,ϵͳ��ֹע�ᣡ</p>
      <%Case "60"%>
      <p>���������˺��зǷ��ַ���ֻ��ʹ�����ģ����ܴ��ո��������ţ���</p>
      <%Case "61"%>
      <p>�����㲻�ܲ��������˳������������ٽ��У�</p>
      <%Case "62"%>
      <p>�������û����Ѿ����ڣ�</p>
      <%Case "63"%>
      <p>����ԭ�û����Ѿ������ڻ����벻��ȷ��</p>
      <%Case "64"%>
      <p>�������ǹٸ����˻��������������ţ���ֹ������</p>
      <%Case "65"%>
      <p>�������ǽ�����Ա��Ϊ�˷������������ֹ������</p>
      <%Case "66"%>
      <p>������ʾ��𰸲�����ͬ��</p>
      <%Case "67"%>
      <p>������Ϊ���ע����������IP��ͬ��Ϊ��ֹ������������120����ע�ᣡ(ͬһ����ע��ʱҲ�����ȴ�!)</p>
      <%Case "68"%>
      <p>����������id�������ڣ�</p>
      <%Case "69"%>
      <p>�������������ղ���ȷ�������޷����ң�</p>
      <%Case "70"%>
      <p>����������𰸲���ȷ������ȡ�أ�</p>
      <%Case "71"%>
      <p>�����㲻�ܲ��������˳������������ٽ��У�</p>
      <%Case "72"%>
      <p>����������ʲô���ǡ������ϴ�</p>
      <%Case "100"%>
      <p>������ӭ���Ĺ��٣�ֻ��վ���Ѿ��ر��������ҵĵ�¼���ܣ����ڲ��ܽ��е�¼�����Ժ�������</p>
      <%Case "101"%>
      <p>������ӭ���Ĺ��٣�����������������վ���������ͬʱ��������Ϊ <font color=red><%=Application("sjjh_chat_maxpeople")%></font> 
        �ˣ����Ժ�������</p>
      <%Case "102"%>
      <p>����վ����ֹ���û�����¼�����Ժ�������</p>
      <%Case "110"%>
      <p>���������ڵ� IP��<font color=red><%=Request.ServerVariables("REMOTE_ADDR")%></font> 
        ������ <%=Application("sjjh_iplocktime")%> ���ӣ����ܽ��������ҡ�<br>
        �������Զ�����ʱ�仹�У�<font color=red><%=ABS(int(Application("sjjh_iplocktime"))-int(Datediff("s",Request.QueryString("lockdate"),now())/60))%></font> 
        ���ӡ�</p>
      <%Case "111"%>
      <p>���������ڵ� IP��<font color=red><%=Request.ServerVariables("REMOTE_ADDR")%></font> 
        ��<font color=red>�����á�</font>���������ܽ��������ҡ�����վ����ϵ��</p>
      <%Case "120"%>
      <p>�����û������зǷ��ַ���ֻ��ʹ�����ģ����ܴ��ո���������,���������</p>
      <%Case "121"%>
      <p>�������뺬�зǷ��ַ���ֻ��ʹ��Ӣ����ĸ�����֣����ܴ��ո񣩡�</p>
      <%Case "122"%>
      <p>������ν���зǷ��ַ���ֻ��ʹ�����ġ�Ӣ����ĸ�����֣����ܴ��ո񣩡�</p>
      <%Case "123"%>
      <p>���������뺬�зǷ��ַ���ֻ��ʹ��Ӣ����ĸ�����֣����ܴ��ո񣩡�</p>
      <%Case "124"%>
      <p>����ȷ�����뺬�зǷ��ַ���ֻ��ʹ��Ӣ����ĸ�����֣����ܴ��ո񣩡�</p>
      <%Case "125"%>
      <p>�����û����ĳ��ȳ��� 10 ���ַ���һ������ռ�����ַ�����</p>
      <%Case "126"%>
      <p>������ν�ĳ��ȳ��� 4 ���ַ���һ������ռ�����ַ�����</p>
      <%Case "127"%>
      <p>�����û�������Ϊ�ա�</p>
      <%Case "128"%>
      <p>�������벻��Ϊ�ա�</p>
      <%Case "129"%>
      <p>�������벻�ܺ��û�����ͬ��</p>
      <%Case "130"%>
      <p>�����Բ��𣬸��û���Ϊϵͳ��������������ʹ��������ֵ�¼��</p>
      <%Case "131"%>
      <p>�����Բ��𣬸��û������в������ۣ�������ʹ��������ֵ�¼��</p>
      <%Case "132"%>
      <p>�����Բ��𣬳�ν���в������ۣ�������ʹ�������ν��¼��</p>
      <%Case "140"%>
      <p>�����Բ��𣬸��û�������ʹ���У��������������Ҫ�Ĳ�����<br>
        ������������״�ʹ�ø��û�����¼��������Ǹ��û����Ѿ�����������ע�ˣ���ֻ�ܻ����������ֵ�¼��<br>
        �����������ǰ����ʹ�ù����û�����¼�ɹ���������������û����������˵��á�<br>
        ��һ�ֿ����ǣ���û�������˳������ң��磺���ߡ���ʱ��������Ͽ����ӵȣ�����ʹ�õ��ߴ����������û����������������У��������ر�����������������´򿪣�����ʮ���Ӻ�������¼�����ʵ�ڲ��У��뵽���Բ����������ԣ������Ϊ�������</p>
      <%Case "141"%>
      <p>�����Բ����������ע�⣺�������ִ�Сд�����������������Ҫ�Ĳ�����<br>
        ������������״�ʹ�ø��û�����¼��������Ǹ��û����Ѿ�����������ע�ˣ���ֻ�ܻ����������ֵ�¼��<br>
        �����������ǰ����ʹ�ù����û�����¼�ɹ���������������û����������˵��ã����ҵ����߸��������롣���������������뵽���Բ����������ԣ������Ϊ�������</p>
      <%Case "142"%>
      <p>�����Բ��𣬸��û��������ã��뵽���Բ����������ԣ������Ϊ�������</p>
      <%Case "143"%>
      <p>�����Բ��𣬸��û������߳������ң�ԭ��:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>,5 
        �����ڲ���ʹ�ø��û�����¼������ <font color=red><%=ABS(300-Datediff("s",Request.QueryString("lastkick"),now()))%></font> 
        �롣</p>
      <%Case "150"%>
      <p>�����Բ��𣬸��û�����δ��ע�ᣬ��Ȼ�����޸������ˡ�</p>
      <%Case "151"%>
      <p>�����Բ�����û��ָ���������롱����ô�޸������أ�</p>
      <%Case "152"%>
      <p>�����������롱���������ȫ��ͬ���ò������޸���������</p>
      <%Case "153"%>
      <p>�����������롱�������û�����ͬ��</p>
      <%Case "160"%>
      <p>�����Բ��𣬸��û�����δ��ע�ᣬ��Ȼ���ܡ���ɱ���ˡ�</p>
      <%Case "161"%>
      <p>������ȷ�����롱Ϊ�գ�����ִ����ɱ������</p>
      <%Case "162"%>
      <p>������ȷ�����롱�롰���롱��һ�£�����ִ����ɱ������</p>
      <%Case "163"%>
      <p>�������û��������ñ���������ִ����ɱ������</p>
      <%Case "164"%>
      <p>�����������󣬲���ִ����ɱ������</p>
      <%Case "165"%>
      <p>������û�и���������ıɱ����</p>
      <%Case "166"%>
      <p>�����������������ͬ����ע�ᣡ</p>
      <%Case "167"%>
      <p>�����û���������ע�ᣬ��ѡ�������û�����</p>
      <%Case "001"%>
      <p>�����ó���ִ���˷Ƿ�������������ֹͣʹ�ã�����������𲻿�Ԥ��ĺ����������������ȡ����ϵ��<br>
        ���䣺<a href="mailto:yxt_email@etang.com?subject=%D3%D0%B9%D8E%CF%DF%BD%AD%BA%FE%B5%C4%CE%CA%CC%E2">yxt_email@etang.com</a><br>
      </p>
      <%Case "200"%>
      <p>����û�д���ɹ���ʾ���������ݣ����ܲ鿴��</p>
      <%Case "201"%>
      <p>����û�д���ɹ���ʾ�����ݣ����ܲ鿴��</p>
      <%Case "210"%>
      <p>��������δ��¼�����Ѿ���ʱ�Ͽ������ӣ������µ�¼��</p>
      <%Case "220"%>
      <p>���������Ժη������зǷ��ַ���Ҳ����ʹ��HTML�﷨��</p>
      <%Case "221"%>
      <p>������E-Mail����ַ���зǷ��ַ���Ҳ����ʹ��HTML�﷨��</p>
      <%Case "222"%>
      <p>��������ҳ��ַ�����зǷ��ַ���Ҳ����ʹ��HTML�﷨��</p>
      <%Case "223"%>
      <p>���������Ժη������ȳ���30�ַ���1������ռ2�ַ���</p>
      <%Case "224"%>
      <p>������E-mail�����ȳ���50�ַ���</p>
      <%Case "225"%>
      <p>��������ҳ��ַ�����ȳ���50�ַ� ��</p>
      <%Case "226"%>
      <p>���������˼�顱���ȳ���200�ַ� ��</p>
      <%Case "227"%>
      <p>������E-mail����ַ��ʽ����</p>
      <%Case "228"%>
      <p>��������ҳ��ַ����ʽ���� ��</p>
      <%Case "229"%>
      <p>�����û��������ڣ������޸ĸ�����Ϣ��</p>
      <%Case "230"%>
      <p>����������Ҫ��ѯ���û�����</p>
      <%Case "231"%>
      <p>�����û�����<font color=red><%=Request.QueryString("name")%></font> �����ڡ�</p>
      <%Case "240"%>
      <p>�����ؼ���Ϊ�գ�����������</p>
      <%Case "250"%>
      <p>�������С���ɫ������Ŀ��Ϊ������Ŀ������д������</p>
      <%Case "251"%>
      <p>�����û�����<font color=red><%=Request.QueryString("name")%></font> �����ڣ�����ʹ�á����Ļ�����ʽ��</p>
      <%Case "252"%>
      <p>��������Ϊ�գ�����ʹ���û�����<font color=red><%=Request.QueryString("name")%></font> 
        �������ԡ�</p>
      <%Case "253"%>
      <p>����������󣬲���ʹ���û�����<font color=red><%=Request.QueryString("name")%></font> 
        �������ԡ�</p>
      <%Case "254"%>
      <p>������д�����ĳ��ȳ���20���ַ���1������=2���ַ�����</p>
      <%Case "255"%>
      <p>���������⡱�ĳ��ȳ���40���ַ���1������=2���ַ�����</p>
      <%Case "256"%>
      <p>���������ݡ��ĳ��ȳ���1024���֡�</p>
      <%Case "257"%>
      <p>�������������ĳ��ȳ���20���ַ���1������=2���ַ�����</p>
      <%Case "258"%>
      <p>��������ַ���ĳ��ȳ���20���ַ���1������=2���ַ�����</p>
      <%Case "259"%>
      <p>���������䡱�ĳ��ȳ���50���ַ���1������=2���ַ�����</p>
      <%Case "260"%>
      <p>��������ҳ���ơ��ĳ��ȳ���24���ַ���1������=2���ַ�����</p>
      <%Case "261"%>
      <p>��������ҳ��ַ���ĳ��ȳ���50���ַ���</p>
      <%Case "262"%>
      <p>���������䡱�����������������롣</p>
      <%Case "263"%>
      <p>��������ҳ��ַ���ĸ�ʽ����</p>
      <%Case "264"%>
      <p>�������ԡ����ݡ��в��ܳ��������Ŀհ��С�</p>
      <%Case "265"%>
      <p>���������ظ�ճ����ͬ�����ԡ�</p>
      <%Case "266"%>
      <p>���������⡱���ܰ�����ǵġ�"����'�����š�</p>
      <%Case "300"%>
      <p>�������û����Ѿ����������У������ظ�����,��ʹ�õ��߹����ٽ��뽭������</p>
      <%Case "301"%>
      <p>���������ԡ���������Ա�������ƽ����������У����ػص�¼ҳ�滻����¼��</p>
      <%Case "400"%>
      <p>��������<font color=red><%=Request.QueryString("name")%></font> �����������У����ܷ�����Ϣ��</p>
      <%Case "401"%>
      <p>������Ϣ�ĳ��ȱ���С�� 1024 ���ַ���</p>
      <%Case "402"%>
      <p>������Ϣ�в��ܳ��������Ŀհ��С�</p>
      <%Case "403"%>
      <p>������Ϣ����Ϊ�ա�</p>
      <%Case "410"%>
      <p>��������<font color=red><%=Request.QueryString("name")%></font> �����������У�����Ϊ���衣</p>
      <%Case "420"%>
      <p>���������㱻ץ�����ԭ��:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>,���ű�����ɣ���Ҫ�ٸɻ����ˣ�</p>
      <%Case "421"%>
      <p>�����㱻�˴����ˣ�ԭ��:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>�����ǵ�<a href="yamen/disp.asp">������</a>���ɣ�</p>
      <%Case "422"%>
      <p>���������㱻ץ�����ԭ��:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>,��10���Ӻ��ͷŰɣ���Ҫ�ٸɻ����ˣ�</p>
      <%Case "423"%>
      <p>����������δע�ᡢ���˺ű��������׻���û�м�ʱ�����ɱ����������<a href='yamen/joinjh.asp'>ע��</a>����˺Űɣ�</p>
      <%Case "424"%>
      <p>���������㱻�˵�Ѩ��һʱ�仹û���������ˣ�</p>
      <%Case "425"%>
      <p>�����㲻�ǹٸ����ˣ���û��Ȩ���������</p>
      <%Case "426"%>
      <p>�����Ҳ��ǳ��ϣ���û��Ȩ�����������</p>
      <%Case "427"%>
      <p>�����㲻�ǽ������ˣ����ܽ�飡</p>
      <%Case "428"%>
      <p>������ô���£�������ͬ�����ɣ�</p>
      <%Case "429"%>
      <p>����������һ���ˡ���</p>
      <%Case "430"%>
      <p>�������ܵǼǣ��Է��Ľ����ȼ�����5������ż��Ϊ�ջ��������Ա���ͬ����</p>
      <%Case "431"%>
      <p>�������ܵǼǣ���Ϊ���Ѿ��Ǽ�����3����ٵǼǣ�</p>
      <%Case "432"%>
      <p>����������벻��Ϊ�գ�</p>
      <%Case "433"%>
      <p>�������µ����֡�Ƹ�����10����Ϣ����Ϊ�գ�</p>
      <%Case "4333"%>
      <p>������ĵȼ�������������ʲô��������ż��Ҫ3�����ϣ�</p>
      <%Case "4334"%>
      <p>������(��)�ĵȼ�������������ʲô��������ż��Ҫ3�����ϣ�</p>
      <%Case "4335"%>
      <p>�����������Ҳ�������ȡ���ˣ�</p>
      <%Case "4336"%>
      <p>����call!������ʲô�����ﲻ��ӭ�㣬����</p>
      <%Case "434"%>
      <p>���������Լ�ȡ�Լ���������������û������£�</p>
      <%Case "435"%>
      <p>������ʲô��Ц���������Ա�һ�������������ǲ�����ͬ�����ġ�</p>
      <%Case "436"%>
      <p>����ֻ�����洫��һƬ��У���Ż����ŵ����˳�����</p>
      <%Case "437"%>
      <p>���������û�д���ô��Ǯ�˰ɣ�</p>
      <%Case "438"%>
      <p>�����Բ���������Ѿ��ܸɾ��ˣ���Ȫԡÿ��ֻ����ϴһ�Ρ�</p>
      <%Case "439"%>
      <p>������û�е�¼���㲻�ǹٸ����ˣ��㲻�ܽ����������</p>
      <%Case "440"%>
      <p>�����Բ�������û��<a href=go.asp target=_blank>��¼</a></p>
      <%Case "441"%>
      <p>����δ�ҵ����ŵ����ϣ�����δ����</p>
      <%Case "442"%>
      <p>��������������Ŀ�����ɡ��ʺϱ�����д���ʺϱ���Ϊ�С�Ů�����У�</p>
      <%Case "443"%>
      <p>�������󣺹ۿ����ʳ�����Χ�����ϣ����������</p>
      <%Case "444"%>
      <p>�������󣺰��������Ѿ����ڣ�</p>
      <%Case "445"%>
      <p>�������󣺰������Ʊ�����д����</p>
      <%Case "446"%>
      <p>����δ�ҵ����ŵ����ϣ������Ѿ��������ˣ�����.....����ɣ���ô�˵����ţ�</p>
      <%Case "447"%>
      <p>�����������Ѳ����ڣ���ˢ��ҳ�棡</p>
      <%Case "448"%>
      <p>�����ٸ��������������գ�</p>
      <%Case "449"%>
      <p>���������㲻�ǽ�����,�벻Ҫ�󴳽�����</p>
      <%Case "449"%>
      <p>����δ�ҵ����ŵ����ϣ�����δ����</p>
      <%Case "450"%>
      <p>���������ռ�����һ�������ˣ���ⲻ��֧����Ŀ�����ʡ�ŵ�ɣ�</p>
      <%Case "451"%>
      <p>���������㲻������,�벻Ҫ�󴳽�����</p>
      <%Case "452"%>
      <p>�����������ǹٸ����ˣ������Ը�����</p>
      <%Case "453"%>
      <p>���������㲻�ǽ������ˣ��벻Ҫ����Ҵ���</p>
      <%Case "454"%>
      <p>���������㲻�ǽ������ˣ����ǲ��������ִ�̹���С���</p>
      <%Case "455"%>
      <p>�����������������ɣ��������뱾����أ�</p>
      <%Case "456"%>
      <p>�����������©�������ϣ��벻Ҫ�����ˣ�</p>
      <%Case "457"%>
      <p>��������ûǮҲ�����������׬���������������ɣ�</p>
      <%Case "458"%>
      <p>������������������������ǲ�����������Ǽǣ�</p>
      <%Case "459"%>
      <p>�������󣺲������ɹ�!!�뷵�أ�</p>
      <%Case "460"%>
      <p>����������������û���أ���ô���봴���ˣ�����ͷ���濴�����������ɣ�</p>
      <%Case "461"%>
      <p>��������������������˰ɣ��ж�����Σ����뽨���ٸ���ѽ��</p>
      <%Case "462"%>
      <p>������������������������Ʒ��</p>
      <%Case "463"%>
      <p>����������û�м��۵���Ʒ�ڱ����У�</p>
      <%Case "464"%>
      <p>����������������ݳ�������ȷ��������ĺ���Ϊ���֣�</p>
      <%Case "465"%>
      <p>�����Բ�����������Ϊ���ˣ�����̫�ͣ����㵽<a href="xg.asp">˼����</a>�����˼������</p>
      <%Case "466"%>
      <p>�����Բ�����ǮׯС����Ӫ�������ֽ���Ŀ̫�࣬���´�������</p>
      <%Case "467"%>
      <p>�����������֣����书����150000���ޣ�ϵͳ�Ѿ��Զ����併��1499999���벻Ҫ�������书�ˣ����Ѿ��ǽ���һ��һ�����ˣ���<a href="index.htm">���µ�¼</a>���ɣ�</p>
      <%Case "468"%>
      <p>�����������֣�����������1�����ޣ�ϵͳ�Ѿ��Զ����併��99000000���벻Ҫ�����������ˣ����Ѿ��ǽ���һ��һ�����ˣ���<a href="index.htm">���µ�¼</a>���ɣ�</p>
      <%Case "469"%>
      <p>������û���������͵�װ�����뵽���й������װ��</p>
      <%Case "470"%>
      <p>������û���κ���Ʒ��</p>
      <%Case "471"%>
      <p>�����������ø���������</p>
      <%Case "472"%>
      <p>��������������Ϣ�У����ڻ�û�����������һ��ʱ�������ɣ�</p>
      <%Case "473"%>
      <p>����������Υ�����������ѱ�������������˺Ų���ʹ���ˣ�</p>
      <%Case "474"%>
      <p>������������Ʒ̫�����С��-10000����������ӭ�㣬��ȥ�ɱ����ɿ��Ǯȥ�¶�Ժ�����ȥ�ɣ���</p>
      <%Case "475"%>
      <p>�����Բ����������������һ���ǿ�ֵ��</p>
      <%Case "476"%>
      <p>�����Բ���,�ٸ�Ŀǰ��δ��������</p>
      <%Case "477"%>
      <p>�����Բ���,Ϊ�˹�ƽ,�ٸ����˲��ܲ��ڴ˻��</p>
      <%Case "478"%>
      <p>������...�ڿ��ϴ�,���ﲻ��ӭ�����,����ȥ���˼Ұ�!</p>
      <%Case "479"%>
      <p>�����Բ���,�����ṩ�Ĵ𰸲��ԣ���ò���������</p>
      <%Case "480"%>
      <p>�����Բ���,������Ѩ�ˣ�ԭ��:<font color=blue>[<%=Request.QueryString("sjyy")%>]</font>,�������ӻ�Ѩ�����Զ��⿪,��һ���ٵ�¼��</p>
      <%Case "481"%>
      <p>����������ȡϱ��ѽ���ú�������㣡</p>
      <%Case "482"%>
      <p>������ĵȼ�̫���ˣ���������Щ�</p>
      <%Case "483"%>
      <p>�������Ѿ��������ˣ��벻Ҫ�������ˣ�</p>
      <%Case "483"%>
      <p>������Ϊ������̫�࣬��������-100���Բ��ܽ���������������ã�</p>
      <%Case "485"%>
      <p>��������ٸ�һ��Ҫս���ȼ�����65��,�����͵��³���10��,����10�ڣ�</p>
      <%Case "490"%>
      <p>�����㲻��������������Ϣһ�����㾫���ټ����ɣ�</p>
      <%Case "491"%>
      <p>������Ľ�����Ϊ�գ��뵽������̳ר���Ǽǣ���վ�������</p>
      <%Case "492"%>
      <p>һ����෢���Σ���̫������</p>
      <%Case else%>
      <p>�����Բ��𣬸ó�������δ���Ǽǣ�����ϵվ�������</p>
      <%End Select%>
    </td>
  </tr>
  <tr> 
    <td colspan="2" align="center" valign="top" height="17"> <a href="index.asp" onMouseOut=MM_swapImgRestore() onMouseOver="MM_swapImage('back','','jhimg/err_but2.gif',1)">
<img border=0 height=20 name=back src="jhimg/err_but1.gif" width=73></a><br>
      <img height=5 src="jhimg/err_bot.gif" 
        width=400></td>
  </tr>
</table>
</body>
</html>