<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
call nav()
stats="���߲���"
call head_var(2,0,"","")
if Cint(GroupSetting(14))=0 then
Errmsg=Errmsg+"<br>"+"<li>��û���ڱ��������߲��Ե�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
call dvbbs_error()
else
response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25><b>���߲���</b></td></tr><tr><td valign=middle class=tablebody1 height=100><CENTER>"
'�������Ȩ���ڳ���������վ����
%>

<html>

<head>
<title>��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
</head>
<style>
<!--
BODY {SCROLLBAR-FACE-COLOR: #D6D6D6; SCROLLBAR-HIGHLIGHT-COLOR: #3A6EA5; SCROLLBAR-SHADOW-COLOR: #ffffff; SCROLLBAR-3DLIGHT-COLOR: #FFFFFF; SCROLLBAR-ARROW-COLOR:  #8FA5B6; SCROLLBAR-TRACK-COLOR: #f3f3f3; SCROLLBAR-DARKSHADOW-COLOR: #3A6EA5; }
-->
</style>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="2" topmargin="2">
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="600" height="23" background="ontest/bg1.gif">
    <table border="0" cellspacing="0" width="600" id="AutoNumber2" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="25">
        <img border="0" src="ontest/err1.gif" width="24" height="23"></td>
        <td width="536" align="left">
        <table border="0" cellspacing="0" width="40%" id="AutoNumber5" cellpadding="0" height="15">
          <tr>
            <td width="100%" valign="bottom" height="15">&nbsp;&nbsp;
            <font color="#FF0000">&nbsp;���顢�˼ʡ�����</font><font color="#0000FF"> 
            ���߲��ԣ�</font></td>
          </tr>
        </table>
        </td>
        <td width="39" valign="bottom" style="border-right: 1px solid #000000">
        <img border="0" src="ontest/help.gif" start="fileopen" width="15" height="18" align="absbottom">
        <a href="javascript:window.close()">
        <img border="0" src="ontest/close.gif" width="15" height="18" align="absbottom" alt="�رմ���"></a></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td height="562" valign="top" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-bottom: 1px solid #000000">
    <div align="center">
      <center>
      <table border="1" cellspacing="1" height="570" style="border-left: 1px solid #000000; border-right: 1px solid #000000; border-bottom: 1px solid #000000" width="596" id="AutoNumber3" bgcolor="#ECEDEF">
        <tr>
          <td width="100%"> 
          <div align="center">
            <center>
            <table border="1" width="580" id="AutoNumber6" height="549" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" cellspacing="1" background="ontest/bg2.gif">
              <tr>
                <td width="80" height="145" valign="top" align="center">
                <p style="margin-top: 5">
                <img border="0" src="ontest/love.GIF" width="80" height="60"></p>
                <p><font color="#0000FF">[</font> <font color="#FF0000">�������</font>
                <font color="#0000FF">]</font>��</td>
                <td width="500" height="145" valign="top">
                <div align="center">
                  <center>
                  <table border="0" cellspacing="3" width="95%" id="AutoNumber7" cellpadding="0" height="1" background="ontest/bg3.gif">
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=001>���������ɹ���</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=002>���Ƿ��������</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=003>�Գ���̶Ȳ���(����)</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=004>�Գ���̶Ȳ���(Ů��)</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=005>���鷽���Բ�_����ʱ��</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=006>�ϸ��ɷ����</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=007>��ԼΣ��</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=008>�������Խ���������</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=009>�����ʧ��</td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=010>���ʺ����ʾ����ʽ</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=011>˭����һ������?</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=012>���������ȡ��</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=013>����ѡ��������</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=014>���ܼ�Ԧ������?</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=015>���Ƿ�ֵ�����и�����</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=016>������ѹ������</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=017>�����������͵��ɷ�</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=018>������Ḷ������</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=019>������˺�ɫ���</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=020>ʮ��ֽҪ�㶪��ʱ</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=021>���˽���ô</td>
                    </tr>
                    <tr>
                      <td width="33%" height="17"><a href=z_answer.asp?id=022>��˧������Щ������</a></td>
                      <td width="33%" height="17"><a href=z_answer.asp?id=023>������������һ�ֽ�ָ��</a></td>
                      <td width="34%" height="17"><a href=z_answer.asp?id=024>����Σ��</a></td>
                    </tr>
                  </table>
                  </center>
                </div>
                </td>
              </tr>
              <tr>
                <td width="80" height="215" valign="top" align="center">
                <p style="margin-top: 5">
                <img border="0" src="ontest/zonghe.GIF"></p>
                <p><font color="#0000FF">[</font> <font color="#FF0000">�ۺϲ���</font>
                <font color="#0000FF">]</font><p>��</td>
                <td width="500" height="215" valign="top">
                <div align="center">
                  <center>
                  <table border="0" cellspacing="3" width="95%" id="AutoNumber7" cellpadding="0" height="47" background="ontest/bg3.gif">
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=101>���Ǹ��ֹ۵�����</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=102>���Ĳ���</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=103>���б����ʶ��</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=104>���ʺϵĹ���</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=105>�Ƿ����ڴ������˴򽻵�</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=106>���ػ���־��</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=107>���Ǹ����ȵ�����</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=108>������Ĺ�ͨ����</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=109>���ڻ�������˾��ì����</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=110>������ļ�ͥ�Ƿ�����</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=111>Ϊɶ������Ӯ������������</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=112>����Ͷ���˵�&quot;���EQ&quot;</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=113>��ǵ�...��</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=114>ѡ�����ĸ���֦��</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=115>�㽫�������Ƽƻ�</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=116>������ļ�ͥ�Ƿ�������</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=117>��־���Բ�</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=118>���Ǹ���Ľ���ٵ�����</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=119>���Ǻ�ǿ��ִ������</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=120>������Ӧ�Բ���</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=121>���θв���</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=122>����ɫ�б���԰���</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=123>�������</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=124>�����Բ���</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=125>���������˷ܵ�����</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=126>��־������</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=127>���н���������</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=128>���ر�������</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=129>����ʱͨ���ǣ�</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=130>����ʲô���������أ�</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=131>����ʲô��?</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=132>ϰ�߿���֪��һ����</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=133>ϲ��ʲô������������ԣ�</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=134>����ʲô��������</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=135>�����ػ���־������</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=136>���������˷ܵ�����</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="16"><a href=z_answer.asp?id=137>���ر�������</a></td>
                      <td width="33%" height="16"><a href=z_answer.asp?id=138>��ĳɹ�ָ���ж��٣�</a></td>
                      <td width="34%" height="16"><a href=z_answer.asp?id=139>������Ĺ�ͨ����</a></td>
                    </tr>
                    </table>
                  </center>
                </div>
                </td>
              </tr>
              <tr>
                <td width="80" height="189" valign="top" align="center">
                <p style="margin-top: 5">
                <img border="0" src="ontest/gexing.GIF"></p>
                <p><font color="#0000FF">[</font> <font color="#FF0000">���Բ���</font>
                <font color="#0000FF">]</font><p>��</td>
                <td width="500" height="189" valign="top">
                <div align="center">
                  <center>
                  <table border="0" cellspacing="3" width="95%" id="AutoNumber7" cellpadding="0" height="17" background="ontest/bg3.gif">
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=201>�������˵ĵ�һӡ����Σ�</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=202>���������Ǵ��������</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=203>��ܾ����¹���</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=204>������������</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=205>�����ܻ�ӭ������</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=206>��������ϵˮƽ����</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=207>��ȴ����</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=208>���Ƿ���ƽ���ͣ�</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=209>Ů���������飿</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=210>������һ���ۺϲ���</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=211>�罻�ۺ�</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=212>�弸������</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=213>����Ů��͹���</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=214>�����������</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=215>����Ϊ���Ǻ��ֽ����</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=216>Ů��������˵ʲô��</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=217>��Ǯҡ����</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=218>���켦��������</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=219>������</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=220>�»���һ����</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=221>��ҹ������</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=222>����˭��</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=223>�����󸻼�</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=224>�ĸ����͵��к���������</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=225>���������</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=226>��������</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=227>�˼ʹ�ϵ�Ĳ���</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=228>�ĸ����͵�Ů����������</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=229>�����������</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=230>�������Ӻ���</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=231>�㴦�����������ǿ��</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=232>��������һ��������</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=233>��������������밮������</a></td>
                    </tr>
                    <tr>
                      <td width="33%" height="15"><a href=z_answer.asp?id=234>������������е�������</a></td>
                      <td width="33%" height="15"><a href=z_answer.asp?id=235>�����԰��Ƿ��кܺõĹ�ͨ</a></td>
                      <td width="34%" height="15"><a href=z_answer.asp?id=236>�������������أ�</a></td>
                    </tr>
                  </table>
                  </center>
                </div>
                </td>
              </tr>
            </table>
            </center>
          </div>
          </td>
        </tr>
      </table>
      </center>
    </div>
    </td>
  </tr>
</table>
</body>

</html>
<%
response.write "</CENTER></td></tr></table>"
end if
call footer()
%>