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
<script language=javascript>
function gBon() {
if (document.af.gbbox.checked)
{
document.all["el"].innerHTML ="<strong><object classid='clsid:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA' width='152' height='50' align='top' id='RealAudio1' controls='StatusPanel' console='cons'><param name='_ExtentX' value='7938'><param name='_ExtentY' value='2646'><param name='AUTOSTART' value='1'><param name='SHUFFLE' value='0'><param name='PREFETCH' value='0'><param name='NOLABELS' value='0'><param name='LOOP' value='0'>><param name='SRC' value='rtsp://www.szr.com.cn/encoder/szr2#39;><param name='NUMLOOP' value='0'><param name='CENTER' value='0'><param name='MAINTAINASPECT' value='0'><param name='BACKGROUNDCOLOR' value='#000000'><param name='CONTROLS' value='ControlPanel,StatusPanel'></object></strong>"
}
else
{
document.all["el"].innerHTML ="<iframe name=live border=0 marginwidth=0 marginheight=0 src='' frameborder=no width=0 scrolling=no height=0 target='_blank'></iframe>"
}
}
</script>
<html>
<head><META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();

if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}</script>

<style type='text/css'>
.aq {  font-size: 9pt; line-height: 150%; color: #000000; filter: DropShadow(Color=#000000, OffX=1, OffY=1, Positive=1)}
.webstyle   {font-family: Webdings; font-size: 9pt}
td{font-size:9pt;color:83a8d5}
.yy4{filter:dropshadow(color=#ffffff,offx=1,offy=1); position: relative; width: 5% }
.yy3{filter:dropshadow(color=#000000,offx=1,offy=1); position: relative; width: 2% }
A:hover {CURSOR:url('1.cur');}
body{font-size:9pt;CURSOR: url('56.ani');}input{font-size:9pt;}a{font-size:9pt;color:83a8d5;text-decoration:none;}
</style>
<title></title>
</head>
<body background="f2.gif" text="#990000" link="#ffffff" vlink="#ffffff" alink="#ffffff" leftmargin="0" topmargin="0" bgproperties="fixed" marginwidth="0" marginheight="0" oncontextmenu=window.event.returnValue=false onselectstart=event.returnValue=false ondragstart=window.event.returnValue=false>
<form name=af method=POST action='say.asp'  target='d' onsubmit='return(parent.checksays());'>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<input type=hidden name='fs' value='10'>
<input type=hidden name='lh' value='125'>
<input type=hidden name='sy' value=''>
<input type=hidden name='oldsays' value>
<input type=hidden name='oldact' value>
<input type=hidden name='oldtowho' value>
<input type=hidden name='username' readonly  style="text-align:center;font-size:12px;color:008888" size=5 maxlength=10>

    <tr> 
      <td height="20"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
            
          <td width="10" height="29"><div ID="el" style="VISIBILITY: hidden;POSITION: absolute;"></div></td>
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  
                <td height="29">
<div align="center"> 
                    <select name='zt' onChange="rc(this.value);"  style='font-size:12px; background-color:#b7d4f1'>
                      <option selected>��ϯ 
                      <option style=background-color:"#FFEEFF" value="/��ϯ$ ����">���� 
                      <option style=background-color:"#FFEEFF" value="/��ϯ$ ˯��">˯�� 
                      <option style=background-color:"#FFEEEE" value="/��ϯ$ �Է�">�Է� 
                      <option style=background-color:"#EEFFEE" value="/��ϯ$ ����">���� 
                      <option style=background-color:"#FFFFEE" value="/��ϯ$ ����">���� 
                      <option style=background-color:"#EEEEFF" value="/��ϯ$ ����">���� 
                      <option style=background-color:"#EEFFFF" value="/��ϯ$ ���">��� 
                      <option style=background-color:"#EEFFFF" value="/��ϯ$ Լ��">Լ�� 
                      <option style=background-color:"#EEFFFF" value="/��ϯ$ ��˼">��˼ 
                      <option style=background-color:"#EEFFFF" value="/��ϯ$ ϴ��">ϴ�� 
                    </select>
                    �� 
                    <select name='addsays' onchange="document.af.sytemp.focus();" style='font-size:12px; background-color:#b7d4f1'>
                      <option value="��" selected>����</option>
                      <option value="΢΢Ц��">΢Ц</option>
                      <option value="����ض�">����</option>
                      <option value="��������">����</option>
                      <option value="ҡͷ���Ե���ض�">����</option>
                      <option value="������������Ц�Ŷ�">��Ц</option>
                      <option value="��������ض�">����</option>
                      <option value="սս�����ض�">ս��</option>
                      <option value="ë��ë�ŵض�">ë��</option>
                      <option value="�����ض�">���</option>
                      <option value="����˹��ض�">����</option>
                      <option value="ͬ��ض�">ͬ��</option>
                      <option value="�����ֻ��ض�">�ֻ�</option>
                      <option value="��Ҫ�޵ض�">���</option>
                      <option value="���Ŷ�">��</option>
                      <option value="ȭ����ߵض�">ȭ��</option>
                      <option value="��������ض�">����</option>
                      <option value="�ź��ض�">�ź�</option>
                      <option value="�ɴ����۾����ܲ���ض�">����</option>
                      <option value="�Ҹ��ض�">�Ҹ�</option>
                      <option value="���䵹��ض�">����</option>
                      <option value="��ʹ�ض�">��ʹ</option>
                      <option value="������Ȼ�ض�">����</option>
                      <option value="����ض�">����</option>
                      <option value="�����ض�">����</option>
                      <option value="�����ض�">����</option>
                      <option value="ɵ�����ض�">ɵ</option>
                      <option value="������ض�">����</option>
                      <option value="�����޴�ض�">�޴�</option>
                      <option value="���޹��ض�">�޹�</option>
                      <option value="������ض�">����</option>
                      <option value="��ݺݵص����۶�">����</option>
                      <option value="��Ҫ�µض�">����</option>
                      <option value="�޾���ɵض�">�޲�</option>
                      <option value="��������ض�">����</option>
                      <option value="���°�ĭ��">��ĭ</option>
                      <option value="���������ض�">����</option>
                      <option value="�޿��κεض�">����</option>
                      <option value="��������ض�">����</option>
                    </select>
                    �� 
                    <input type="text" name="towho" readonly size="8" onClick="dj()" value="���" style="height: 18; text-align: center; background-image: url('ico/f2.jpg'); border-style: groove; border-width: 1">
                    <font class=yy3>˵</font> 
                    <input type=text name='sytemp'  size=35 maxlength=250 onpaste="return false" style="height: 18; background-image: url('ico/f2.jpg'); border-style: groove; border-width: 1">
                    <a href="#" onClick="gop()"><img src="ico/qian.gif" alt="��һ��" width="32" height="32" border="0" align="absmiddle"></a><a href="#" onClick="gon()"><img src="ico/hou.gif" alt="��һ��" width="32" height="32" border="0" align="absmiddle"></a> 
                    <input type=submit name='subsay' value='����' style="height: 18; background-color:006699;color:b7d4f1;border: 1 double" onmouseover="this.style.color='ffffff'" onmouseout="this.style.color='b7d4f1'">                                                                         
                    </div></td> 
                </tr> 
              </table></td> 
            
          <td width="10" height="29">��</td> 
          </tr> 
        </table></td> 
    </tr> 
    <tr>  
      <td height="20" class="aq"><div align="center">
<select name="bgc" onchange="parent.f1.document.bgColor=parent.f0.document.bgColor=this.options[selectedIndex].value;this.value='#eeeeee';document.af.bgc.options[0].selected=true;"  style="font-size:12px; background-color:#b7d4f1">
          <option value="#eeeeee" selected>���� 
          <option value="FBE7DB" STYLE="background-color:#FBE7DB">���� 
          <option value="ffeaea" STYLE="background-color:#ffeaea">�ۺ� 
          <option value="FCF8E2" STYLE="background-color:#FCF8E2">��� 
          <option value="eaffea" STYLE="background-color:#eaffea">ǳ�� 
          <option value="effaff" STYLE="background-color:#effaff">���� 
          <option value="f2f2ff" STYLE="background-color:#f2f2ff">ǳ�� 
          <option value="eaeaff" STYLE="background-color:#eaeaff">���� 
          <option value="000000" STYLE="background-color:#eaeaff">��ɫ 
          <option value="ffffff" STYLE="background-color:#ffffff">��ɫ 
          <option value="f7f7f7" STYLE="background-color:#f7f7f7">Ĭ�� 
        </select>  
<select name='addwordcolor'  style='font-size:12px; background-color:#b7d4f1'> 
<option style="color:0099FF" value="0099FF" selected>��</option>
<option style="color:8800FF" value="8800FF">��</option> 
<option style="color:000000" value="000000">��</option> 
<option style="color:0088FF" value="0088FF">��</option> 
<option style="color:0000FF" value="0000FF">��</option> 
<option style="color:000088" value="000088">��</option> 
<option style="color:888800" value="888800">��</option> 
<option style="color:008888" value="008888">��</option> 
<option style="color:008800" value="008800">��</option> 
<option style="color:8888FF" value="8888FF">��</option> 
<option style="color:AA00CC" value="AA00CC">��</option> 
<option style="color:8800FF" value="8800FF">��</option> 
<option style="color:888888" value="888888">��</option> 
<option style="color:CCAA00" value="CCAA00">��</option> 
<option style="color:FF8800" value="FF8800">��</option> 
<option style="color:CC3366" value="CC3366">��</option> 
<option style="color:FF00FF" value="FF00FF">��</option> 
<option style="color:3366CC" value="3366CC">��</option>

<option style="color:003300" value="003300">��</option>
<option style="color:006600" value="006600">��</option>
<option style="color:009900" value="009900">��</option>
<option style="color:003333" value="003333">��</option>
<option style="color:006666" value="006666">��</option>
<option style="color:003399" value="003399">��</option>
<option style="color:0033cc" value="0033cc">��</option>
<option style="color:333300" value="333300">��</option>
<option style="color:333333" value="333333">��</option>
<option style="color:660000" value="660000">��</option>
<option style="color:663300" value="663300">��</option>
<option style="color:666600" value="666600">��</option>
<option style="color:669900" value="669900">��</option>
<option style="color:666633" value="666633">��</option>
<option style="color:666666" value="666666">��</option>
<option style="color:666699" value="666699">��</option>
<option style="color:6666cc" value="6666cc">��</option>
<option style="color:336666" value="336666">��</option>
<option style="color:993300" value="993300">��</option>
<option style="color:996600" value="996600">��</option>
<option style="color:993333" value="993333">��</option>
<option style="color:999900" value="999900">��</option>
<option style="color:99cc00" value="99cc00">��</option>
<option style="color:996666" value="996666">��</option>
<option style="color:999966" value="999966">��</option>
<option style="color:999999" value="999999">��</option>
<option style="color:993399" value="993399">��</option>
<option style="color:9999cc" value="9999cc">��</option>
<option style="color:99ccff" value="99ccff">��</option>
<option style="color:cc3300" value="cc3300">��</option>
<option style="color:cc6600" value="cc6600">��</option>
<option style="color:cc9900" value="cc9900">��</option>
<option style="color:cc0033" value="cc0033">��</option>
<option style="color:cc3333" value="cc3333">��</option>
<option style="color:cc6666" value="cc6666">��</option>
<option style="color:cc0099" value="cc0099">��</option>
<option style="color:cc3399" value="cc3399">��</option>
<option style="color:cc6699" value="cc6699">��</option>
<option style="color:cc66cc" value="cc66cc">��</option>
<option style="color:cc99ff" value="cc99ff">��</option>
<option style="color:ff6600" value="ff6600">��</option>
<option style="color:ff6633" value="ff6633">��</option>
<option style="color:ff6699" value="ff6699">��</option>
<option style="color:ffcc00" value="ffcc00">��</option>
<option style="color:ff6633" value="ff6633">��</option>
 
</select> 
<select name='saycolor'  style='font-size:12px; background-color:#b7d4f1'>                                                                          
<option style="color:666699" value="666699" selected>��</option>
<option style="color:008888" value="008888">��</option> 
<option style="color:000000" value="000000">��</option> 
<option style="bcolor:0088FF" value="0088FF">��</option> 
<option style="color:0000FF" value="0000FF">��</option> 
<option style="color:000088" value="000088">��</option> 
<option style="color:888800" value="888800">��</option> 
<option style="color:008888" value="008888">��</option> 
<option style="color:008800" value="008800">��</option> 
<option style="color:8888FF" value="8888FF">��</option> 
<option style="color:AA00CC" value="AA00CC">��</option> 
<option style="color:8800FF" value="8800FF">��</option> 
<option style="color:888888" value="888888">��</option> 
<option style="color:CCAA00" value="CCAA00">��</option> 
<option style="color:FF8800" value="FF8800">��</option> 
<option style="color:CC3366" value="CC3366">��</option> 
<option style="color:FF00FF" value="FF00FF">��</option> 
<option style="color:3366CC" value="3366CC">��</option>
<option style="color:003300" value="003300">��</option>
<option style="color:006600" value="006600">��</option>
<option style="color:009900" value="009900">��</option>
<option style="color:003333" value="003333">��</option>
<option style="color:006666" value="006666">��</option>
<option style="color:003399" value="003399">��</option>
<option style="color:0033cc" value="0033cc">��</option>
<option style="color:333300" value="333300">��</option>
<option style="color:333333" value="333333">��</option>
<option style="color:660000" value="660000">��</option>
<option style="color:663300" value="663300">��</option>
<option style="color:666600" value="666600">��</option>
<option style="color:669900" value="669900">��</option>
<option style="color:666633" value="666633">��</option>
<option style="color:666666" value="666666">��</option>
<option style="color:666699" value="666699">��</option>
<option style="color:6666cc" value="6666cc">��</option>
<option style="color:336666" value="336666">��</option>
<option style="color:993300" value="993300">��</option>
<option style="color:996600" value="996600">��</option>
<option style="color:993333" value="993333">��</option>
<option style="color:999900" value="999900">��</option>
<option style="color:99cc00" value="99cc00">��</option>
<option style="color:996666" value="996666">��</option>
<option style="color:999966" value="999966">��</option>
<option style="color:999999" value="999999">��</option>
<option style="color:993399" value="993399">��</option>
<option style="color:9999cc" value="9999cc">��</option>
<option style="color:99ccff" value="99ccff">��</option>
<option style="color:cc3300" value="cc3300">��</option>
<option style="color:cc6600" value="cc6600">��</option>
<option style="color:cc9900" value="cc9900">��</option>
<option style="color:cc0033" value="cc0033">��</option>
<option style="color:cc3333" value="cc3333">��</option>
<option style="color:cc6666" value="cc6666">��</option>
<option style="color:cc0099" value="cc0099">��</option>
<option style="color:cc3399" value="cc3399">��</option>
<option style="color:cc6699" value="cc6699">��</option>
<option style="color:cc66cc" value="cc66cc">��</option>
<option style="color:cc99ff" value="cc99ff">��</option>
<option style="color:ff6600" value="ff6600">��</option>
<option style="color:ff6633" value="ff6633">��</option>
<option style="color:ff6699" value="ff6699">��</option>
<option style="color:ffcc00" value="ffcc00">��</option>
<option style="color:ff6633" value="ff6633">��</option> 
</select> 
         
<%sjjh_name=sjjh_name%>                                                                          
<select name='command' onchange="rc(this.value);document.af.command.options[0].selected=true;" style='font-size:12px; background-color:#b7d4f1'>                                                                          
<%                                                                          
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("sjjh_usermdb") 
rs.open "select ��Ա�ȼ�,����,���,ְҵ from �û� where ����='" & sjjh_name &"'",conn,2,2 
hydj=rs("��Ա�ȼ�") 
jhmp=rs("����") 
jhzy=rs("ְҵ") 
jhsf=rs("���") 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
%> 
<option value="" selected>��Ҫ����
<option style="color:#FF6633" value="">==����==

<option style="color:#FF6633" value="/����$ ��Һã��������������뿪������һ������� ">��ʱ�뿪
<option style="color:#FF6633" value="/����$ ">�һ�����
<option style="color:#FF6633" value="/�����ͻ�$ ">�����ͻ� <%if chatinfo(5)=0 then%>
<option style="color:#FF6633" value="/����$ ">�������� <%end if%>
<option style="color:#FF6633" value="/����$ ����">�����ֵ�
<option style="color:#FF6633" value="/����$ �鿴">�鿴�ֵ�
<option style="color:#FF6633" value="/����$ ɾ��">���۶���
<option style="color:#FF6633" value="/���$ Ҫ����ĺ�������">����ֵ�
<option style="color:#FF6633" value="/����$ ����">��������
<option style="color:#FF6633" value="/����$ �鿴">�鿴����
<option style="color:#FF6633" value="/����$ ɾ��">ȥ������
<option style="color:#FF6633" value="/��Ϣ$">�鿴��Ϣ
<option style="color:#FF6633" value="/����$ ����Һã����ҹ�֧ͬ�ֺͰ������ֽ������">�� 
�� �� <% if Application("sjjh_baowu")>0 then%>
<option style="color:#FF6633" value="/����$ ">�������� <%end if%>
<option style="color:#FF6633" value="/�ᱦʤ��$ ">�ᱦʤ��
<option style="color:#FF6633" value="/�Ͻ��«$ ">�Ͻ��«
<option style="color:#B000D0" value="">==����==
<option style="color:#B000D0" value="/��Ǯ$ 0">���д�Ǯ
<option style="color:#B000D0" value="/ȡǮ$ 0">����ȡǮ
<option style="color:#B000D0" value="/ת��$ 10000">����ת��
<option style="color:#FF0000" value="">==����==
<option style="color:#FF0000" value="/����$ 100000">��������
<option style="color:#FF0000" value="/��Ǯ$ 1000">�����ֽ�
<option style="color:#FF0000" value="/�Ͷ���$ 10">���Ͷ���
<option style="color:#FF0000" value="/������$ 10">��������
<option style="color:#FF0000" value="/�ͽ��$ 10">���ͽ��
<option style="color:#FF0000" value="/�ͽ�����$ 10">�ͽ�����
<option style="color:#FF0000" value="/��ľ����$ 10">��ľ����
<option style="color:#FF0000" value="/��ˮ����$ 10">��ˮ����
<option style="color:#FF0000" value="/�ͻ�����$ 10">�ͻ�����
<option style="color:#FF0000" value="/��������$ 10">��������
<option style="color:#FF0000" value="/����$ ���|����|����">������Ʒ
<option style="color:#FF0000" value="/�ÿ�$ ��Ƭ��">ʹ�ÿ�Ƭ
<option style="color:#0808CF" value="">==����==
<option style="color:#0808CF" value="/��������$">��������
<option style="color:#0808CF" value="/���ŷ�Ǯ$">���ŷ�Ǯ
<option style="color:#0808CF" value="/ָ������$">ָ������
<option style="color:#0808CF" value="/���һ��$">���һ��
<option style="color:#0808CF" value="/����$">���Ż���
<option style="color:#0808CF" value="/��ͽ$ ">����ͽ��
<option style="color:#0808CF" value="/���$ ͽ����">���ʦ��
<option style="color:#0808CF" value="/��ʦ$ ">��ʦϰ��
<option style="color:#339933" value="">==����==
<option style="color:#339933" value="/����$ ">��������
<option style="color:#339933" value="/��Ŀ$ ">��Ŀ����
<option style="color:#339933" value="/����$ ">���ϰ��
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/�����Ṧ$ ">�����Ṧ
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/��ľ����$ ">��ľ����
<option style="color:#339933" value="/��ˮ����$ ">��ˮ����
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/�����书$ 200">�����书
<option style="color:#339933" value="/��������$ 1000">��������
<option style="color:#339933" value="/����$ 100">���ھ���
<option style="color:#339933" value="/���ڵ���$ 100">���ڵ���
<option style="color:#339933" value="/��������$ 100">��������
<option style="color:#339933" value="/����$ 100">��������
<option style="color:#FF6699" value="">==����==
<option style="color:#FF6699" value="/���$ �氮���">ʾ�����
<option style="color:#FF6699" value="/����$ ���ҵ����˺���">��������
<option style="color:#FF6699" value="/��������$ ���ҵĶ������˺���">��������
<option style="color:#FF6699" value="/��������$ ���ҵ��������˺���">��������
<option style="color:#ff6699" value="/����$ С������">����С��
<option style="color:#FF6699" value="/����$ �㲻�����ҵ����ˡ�">������ɫ
<option style="color:#FF6699" value="/���ŷ���$ �㲻�����ҵĶ������ˡ�">��ǵ���
<option style="color:#FF6699" value="/���ŷ���$ �㲻�����ҵ��������ˡ�">���Կ�η
<option style="color:#FF6699" value="/�鿴����$ ">�鿴����
<option style="color:#ff6699" value="/�������$ ">�������
<option style="color:#ff6699" value="/��������$ ">��������
<option style="color:#FF6699" value="/����$ ">ǧ�ﴫ��
<option style="color:#FF6699" value="/�Ķ�$ ">�Ķ��о�
<option style="color:#FF6699" value="/ŭ��$ ">��ʨŭ��
<option style="color:#FF6699" value="/����$ ">��������
<option style="color:#333333" value="">==͵Ϯ==
<option style="color:#333333" value="/����$ ����ק������ק">�������� 
<%if chatinfo(5)=0 then%>
<option style="color:#333333" value="/���Ǵ�$ ">��ȡ���� <%end if%>
<option style="color:#333333" value="/����$ 100">���ھ��� <%if chatinfo(5)=0 then%>
<option style="color:#333333" value="/͵Ǯ$ ">͵ȡǮ��
<option style="color:#333333" value="/Ǭ��һ��$ 100000">Ǭ��һ��
<option style="color:#333333" value="/�����ӵ�$ ">�����ӵ� 
<option style="color:#333333" value="/�¶�$ ��ҩ��">͵͵�¶�
<option style="color:#333333" value="/Ͷ��$ ������">Ͷ������
<option style="color:#333333" value="/����$ ��ʽ��">���й���
<option style="color:#333333" value="/ͬ���ھ�$ �ֵ��Ǳ��أ���������һ����">ͬ���ھ�  <%end if%>
<option style="color:#996633" value="">==ְҵ== <%if jhzy="С͵" then%>
<option style="color:#996633" value="/͵���$ ">͵��� <%end if%> <%if jhzy="�з�" then%>
<option style="color:#996633" value="/��ƭ��Ů">��ƭ��Ů <%end if%> <%if jhzy="Ů��" then%>
<option style="color:#996633" value="/��ƭ����">��ƭ���� <%end if%> <%if jhzy="���" then%>
<option style="color:#996633" value="/�β�$ ">���ֻش� <%end if%> <%if jhzy="��ʦ" then%>
<option style="color:#996633" value="/����$ ">�̵��书 <%end if%> <%if jhzy="����ʦ" then%>
<option style="color:#996633" value="/��ǩ$ ">�����ʲ� <%end if%><%if jhzy="��ҹ��" then%>
<option style="color:#996633" value="/��ҹ��$ ">��ҹ��ʦ <%end if%>
<option style="color:#996633" value="/����$ ���Һ�˭Ҳ�����Դ��㣬��Ҳ���ܴ��ˡ�">��������
<option style="color:#996633" value="/����$">�޳ɻ���
<option style="color:#996633" value="/����$">�ն�����
<option style="color:#996633" value="/���崺��$">���崺��
<option style="color:#996633" value="/��������$">��������
<option style="color:#996633" value="/�뿪����$">��������
<option style="color:#993333" value="/�뿪����$">==����== <%if jhsf="����" then%>
<option value="/���յ���$">���յ���
<option value="/�������$">�������
<option value="/���$ �����">������
<option value="/���$ ����">��ⳤ��
<option value="/���$ ����">��⻤��
<option value="/���$ ����">�������
<option value="/���$ ����">ȡ�����
<option value="/���ŷ���$">���ŷ��� <%end if
if (jhsf="����" or jhsf="����") and sjjh_grade>=4 then%>
<option value="/���յ���$">���յ���
<option style="color:blue" value="/��Ѩ$ ��д��ԭ�� ">��ѵ����
<option style="color:blue" value="/����$ 100 ">��������
<option style="color:red" value="/������$ ���� ">�� �� ��
<option style="color:red" value="/������$ ���� ">�� �� ��
<option style="color:red" value="/���ɷ���$ 1000 ">���ɷ���
<option style="color:red" value="/�����̷�$ 1000 ">�����̷� <%end if
if sjjh_grade>=6 then%>
<option style="color:blue" value="/���˷�$ 5000000">�� �� ��
<option style="color:blue" value="/��Ѩ$ ��д��ԭ�� ">��Ѩ����
<option style="color:blue" value="/����$ ��д��ԭ�� ">�������
<option style="color:blue" value="/����$ ��д��ԭ�� ">��������
<option style="color:blue" value="/����$ ��д��ԭ�� ">���β���
<option style="color:blue" value="/����$ ��д��ԭ�� ">���㵷�� <%end if
if sjjh_grade>=9 then%>
<option style="color:blue" value="/��ip$ ">�鿴IP
<option style="color:blue" value="/��Ѩ$ �û���">��Ѩ����
<option style="color:red" value="/״̬$ ">�鿴״̬
<option style="color:red" value="/�Ŵ�$ ">�Ŵ�����
<option style="color:red" value="/����$ ">��������
<option style="color:red" value="/���ջ���$ ��д�ϻ������ݡ���">���ջ���
<option style="color:red" value="/����$ ����">վ������ <%end if
if sjjh_grade>=10 then%>
<option style="color:red" value="/վ������$">վ������
<option style="color:red" value="/���չٸ�$">���չٸ�
<option style="color:red" value="/�鿴��Ʒ$">�鿴��Ʒ
<option style="color:red" value="/ը��$">����ը��
<option style="color:red" value="/��ը����$">��ը����
<option style="color:red" value="/��ըӲ��$">��ըӲ��
<option style="color:red" value="/ԭ�ӵ�$ ԭ��">ԭ�ӵ�
<option style="color:red" value="/���$ ��д��ԭ�� ">�������
<option style="color:red" value="/�ͷ�$ �û���">������
<option style="color:red" value="/ͨ��$ ͨ��������">ͨ���˷�
<option style="color:red" value="/���$ ͨ��������">ȡ��ͨ��
<option style="color:red" value="/���ip$ ĳһ�˵��ң��������ip��ͬ�û���">&nbsp;���ip&nbsp;
<option style="color:red" value="/����ip$ ԭ�����������ж�ε��ң�">&nbsp;����ip&nbsp;
<option style="color:red" value="<script>parent.f3.location.href='savevalue.asp';</script>�������ȫ�����!">ȫ�����
<option style="color:red" value="/ն��$ ԭ�� ">ն���˷�
<%end if%>
</select>
 <select name='addsign' onChange="rc(this.value);" style='font-size:12px; background-color:#b7d4f1'>                                                                          
        <option value="" selected>�����Ṧ</option>
        <option style="color:#B000D0" value="">==ħ��==</option>
        <option style="color:#B000D0" value="/��ʯ�ɽ�$ ">��ʯ�ɽ�</option>
        <option style="color:#B000D0" value="/�����Ϸ�$ ">�����Ϸ�</option>
        <option style="color:#B000D0" value="/������ɽ$ ">������ɽ</option>
        <option style="color:#B000D0" value="/���Ĵ�$ ">���Ĵ�</option>
        <option style="color:#B000D0" value="/�вƽ���$ ">�вƽ���</option>
        <option style="color:#B000D0" value="/���ֻش�$ ">���ֻش�</option>
        <option style="color:#B000D0" value="/�ػ�ھ�$ ">�ػ�ھ�</option>
        <option style="color:green" value="">==˽��==</option>
        <option style="color:green" value="/���ͷ���$ 100">���ͷ���</option>
        <option style="color:green" value="/�淨��$ 100">��ŷ���</option>
        <option style="color:green" value="/ȡ����$ 100">ȡ�ط���</option>
        <option style="color:green" value="/Ѱˮ����$ ">Ѱ��ˮ��</option>
        <option style="color:green" value="/Ѱ�ҷ���$ ">Ѱ�ҷ���</option>
        <option style="color:green" value="/Ѱ��ħ��$ ">Ѱ��ħ��</option>
        <option style="color:green" value="/Ѱ������$ ">Ѱ������</option>
        <option style="color:green" value="/Ѱ�ҽ��$ ">Ѱ�ҽ��</option>
        <option style="color:green" value="/ħ����ʯ$ ">ħ����ʯ</option>
        <option style="color:green" value="/���յ���$ ">���յ���</option>
        <option style="color:green" value="/�ٱ���ͨ$ ">�ٱ���ͨ</option>
        <option style="color:green" value="/������$ ">Ѱ�ұ���</option>
        <option style="color:green" value="/�����$ 100">�����</option>
        <option style="color:green" value="/�ⶾ��$ 100">�ⶾ��</option>
        <option style="color:green" value="/ƽ����$ 100">ƽ����</option>
       <option style="color:green" value="/͵����$ 100">͵����</option>
       <option style="color:green" value="/ҡǮ��$ 100">ҡǮ��</option>
       <option style="color:green" value="/�����$ 100">�����</option>
        <option style="color:green" value="/˧����$ ">˧����</option>
        <option style="color:green" value="/������$ ">������</option>
        <option style="color:green" value="/���黷$ ">���黷</option>
        <option style="color:green" value="/������$ ">������</option>
        <option style="color:red" value="">==ʹչ==</option>
        <option style="color:red" value="/��������$ ">��������</option>
        <option style="color:red" value="/��ʩ��$ ">��ʩ����</option>
        <option style="color:red" value="/�Ȼ��˼�$ ">�Ȼ��˼�</option>
        <option style="color:red" value="/ħ������$ ">ħ������</option>
        <option style="color:red" value="/û�շ���$ ">û�շ���</option>
        <option style="color:red" value="/û��ħ��$ ">û��ħ��</option>
        <option style="color:red" value="/�Ի��$ ">�Ի��</option>
        <option style="color:red" value="/��������$ ">��������</option>
        <option style="color:red" value="/���Ʊ�ʯ$ ">���Ʊ�ʯ</option>
        <option style="color:blue" value="">==����==</option>
        <option style="color:blue" value="/ħ��ˮ��$ ">ħ��ˮ��</option>
        <option style="color:blue" value="/�����ӵ�$ ">�����ӵ�</option>
        <option style="color:blue" value="/������$ ">������ </option>
        <option style="color:blue" value="/����׶$ ">����׶ </option>
        <option style="color:blue" value="/Ѫ����$ ">Ѫ���� </option>
        <option style="color:blue" value="/���鵶$ ">���鵶 </option>	
        <option style="color:blue" value="/������$ ">������ </option>
        <option style="color:blue" value="/ʹ��$ ������ ">������</option>
        <option style="color:blue" value="/ʹ��$ ������ ">������ </option>
        <option style="color:blue" value="/ʹ��$ ��ڤ�� ">��ڤ�� </option>
        <option style="color:blue" value="/ʹ��$ ��ȡ�� ">��ȡ�� </option>
        <option style="color:blue" value="/ʹ��$ ����ȭ ">����ȭ </option>
        <option style="color:blue" value="/ʹ��$ ʥ���� ">ʥ���� </option>
        <option style="color:blue" value="/ʹ��$ �������� ">�������� </option>
        <option style="color:#FF6699" value="">==����==</option>
        <option style="color:#FF6699" value="/������$ ">������</option>
        <option style="color:#FF6699" value="/������$ ">������</option>
        <option style="color:#FF6699" value="/��ť��$ ">��ť��</option>
        <option style="color:#FF6699" value="/������ť$ ">������ť</option>
        <option style="color:#FF6699" value="/���°�ť$ ">���°�ť</option>
         <option style="color:#FF6699" value="/����$ ">������</option>
         <option style="color:#FF6699" value="/�ߵ�$ ">�ߵ�����</option>
	<option style="color:#FF6699" value="/����ħ��$ ����|��ɫ|��С|����">����ħ��</option>
        <option style="color:#FF6699" value="/�ƶ�ħ��$ ����|��ɫ|����|��ʽ|�ٶ�|��ʱ">����ħ��</option>
        <option style="color:#FF6699" value="/��ťħ��$ ����|��ɫ|��������">��ťħ��</option>
        <option style="color:#996633" value="">==˽��==</option>
        <option style="color:#996633" value="/�����Ṧ$ 100">�����Ṧ</option>
        <option style="color:#996633" value="/�Ṧ�ݴ�$ 100">�Ṧ�ݴ�</option>
        <option style="color:#996633" value="/����Ṧ$ 100">����Ṧ</option>
        <option style="color:#000000" value="">==͵��==</option>
        <option style="color:#000000" value="/��ȡ�Ṧ$ ">��ȡ�Ṧ</option>
        <option style="color:#000000" value="/Ѱ������$ ">Ѱ������</option>

      </select>
     <select name='dubo' onchange="rc(this.value);document.af.dubo.options[0].selected=true;" style='font-size:12px; background-color:#b7d4f1'>                                                                            
<option value="" selected>&nbsp;�Ĳ�&nbsp;
<option value="/���˿�$ 1000|����[�򣺽�ң��������������������Ṧ]" style=color:#000000>���˿� 
<option value="/����$ ��֤|��[��ʤ]" style=color:#000000>�˹�֤
<option value="/���齫$ 1000|����[�򣺽�ң�����������]" style=color:#000000>���齫 
<option value="/����$ ��֤|��[��ʤ]" style=color:#000000>�鹫֤
<option style="color:blue" value="/��������">��&nbsp;&nbsp;��
<option style="color:blue" value="/˫�˶Ĳ�$ 1">��ʯͷ
<option style="color:blue" value="/˫�˶Ĳ�$ 2">�ļ���
<option style="color:blue" value="/˫�˶Ĳ�$ 3">�Ĳ���
<option style="color:red" value="/�Զ�����$100000">������
<option style="color:red" value="/�Զ�$10">�Ľ��
<option style="color:red" value="/�Ķ�$100">�Ķ���
<option style="color:red" value="/�ķ���$100">�ķ���
<option style="color:red" value="/���Ṧ$100">���Ṧ
<option style="color:red" value="/����$ ��������1-1000">��&nbsp;&nbsp;��
<option style="color:red" value="/�Ĳ�$ ��Ϣ">����Ϣ
<option style="color:blue" value="/�Ĳ�$ ����ׯ">����ׯ
<option style="color:blue" value="/��ע$ ��&��&10000">��Ѻ��
<option style="color:blue" value="/��ע$ ��&С&10000">��ѺС
<option style="color:blue" value="/��ע$ ��&˫&10000">��Ѻ˫
<option style="color:blue" value="/��ע$ ��&��&10000">��Ѻ��
<option style="color:green" value="/�Ĳ�$ ����ׯ">����ׯ
<option style="color:green" value="/��ע$ ��&��&100">��Ѻ��
<option style="color:green" value="/��ע$ ��&С&100">��ѺС
<option style="color:green" value="/��ע$ ��&˫&100">��Ѻ˫
<option style="color:green" value="/��ע$ ��&��&100">��Ѻ��
  </select>

 <select name='tu' onChange="rc(this.value);document.af.tu.options[0].selected=true;" style='font-size:12px; background-color:#b7d4f1'>                         
<option value="" selected>ͼƬ����</option> 
<option value="[IMG]xl.gif[/IMG]">Ц��</option> 
<option value="[IMG]bm.gif[/IMG]">����</option> 
<option value="[IMG]yy.gif[/IMG]">����</option> 
<option value="[IMG]gl.gif[/IMG]">����</option> 
<option value="[IMG]st.gif[/IMG]">��ͷ</option> 
<option value="[IMG]qc.gif[/IMG]">����</option> 
<option value="[IMG]kl.gif[/IMG]">����</option> 
<option value="[IMG]ml.gif[/IMG]">��Ů</option> 
<option value="[IMG]dg.gif[/IMG]">����</option> 
<option value="[IMG]qt.gif[/IMG]">ȭͷ</option> 
<option value="[IMG]zd.gif[/IMG]">ը��</option> 
<option value="[IMG]yy.gif[/IMG]">ҧ��</option> 
<option value="[IMG]xiong.gif[/IMG]">С��</option> 
<option value="[IMG]mg.gif[/IMG]">õ��</option> 
<option value="[IMG]ws.gif[/IMG]">����</option> 
<option value="[IMG]dh.gif[/IMG]">�绰</option> 
</select> 
 
<select name=scrtx onChange='parent.scrtx(this.value);document.af.scrtx.options[0].selected=true;' style='font-size:12px; background-color:#b7d4f1'>
<option value="" selected>��Чѡ��
<option style="color:green" value="">�ر���Ч
<option style="color:red" value='10'>������
<option style="color:red" value='11'>�ر�����
<option style="color:red" value='12'>��ʾ����
<option style="color:red" value='13'>�ر�����
<option style="color:green" value="STARMOON.js">���´�˵
<option style="color:green" value="js.js">���ķ���
<option style="color:green" value="XMAS.js">�������
<option style="color:green" value="test.js">������ʹ
<option style="color:#B000D0" value="0">��ͷ��
<option style="color:#B000D0" value="1">�� ͷ ��
<option style="color:#B000D0" value="2">�� ͷ ��
<option style="color:#B000D0" value="3">С ͷ ��
<option style="color:#B000D0" value="4">�ر�ͷ��
<option style="color:blue" value='5'>������ʾ
<option style="color:blue" value='6'>��ʾ����
<option style="color:blue" value='7'>��ʾ����
<option style="color:blue" value='8'>��ʾ����
<option style="color:blue" value='9'>��ʾŮ��
<option style="color:red" value='10'>������
<option style="color:red" value='11'>�ر�����
</select> 
 
<input type=button name="tbclutch" value="ȫ��" onClick="javascript:parent.tbclutch();" title="����/����/��ֱ�л�" style="height: 18;background-color:006699;color:ccffff;border: 1 double" onmouseover="this.style.color='ffffff'" onmouseout="this.style.color='ccffff'">                                                                          
 <input type=button value='��ʬ' onClick="window.open('co/cins.asp','f3')" title="�ͷŽ�ʬ��" style="height: 18;background-color:006699;color:ccffff;border: 1 double" onmouseover="this.style.color='ffffff'" onmouseout="this.style.color='ccffff'"><br>                                                                    
 
                <input type=text name='clock' style="text-align: right; font-size: 9pt; height: 18; background-color:006699;color:b7d4f1;border: 1 double" value="" size=3 readonly>                                                                           
 
   			<input type="checkbox" name="addvalues" checked=true>                                                                          
			<a href="#" title='�򿪺�ϵͳ��ʱ�Զ����ԣ��Զ���㣬������һ���ݣ��˹���ֻ���ݵ㷿����ã�' onClick='<%if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))<=0 and sjjh_grade<6 then%>{alert("Ϊ�˱����������ݰ�ȫ������ʹ�����빦�ܣ�")}return false;<%else%>document.af.addvalues.checked=!(document.af.addvalues.checked);<%end if%>document.af.sytemp.focus();'>�ݵ�</a>                                                                           
 
          <input type='checkbox' name='py' accesskey='v' value="ON">                                                                          
            <a href=# onClick='document.af.py.checked=!(document.af.py.checked);' title="��ĺ���(��ż)���ߡ��뿪ʱ�Ƿ��Զ�֪ͨ��">����</a>                                                                           
            <input type='checkbox' name='mdsx' accesskey='v' checked value="ON">                                                                          
            <a href=# onClick='document.af.mdsx.checked=!(document.af.mdsx.checked);parent.m.location.reload();' title="�����˽����˳�ʱ�Ƿ��Զ�ˢ��������">����</a>                                                                           
            <input onClick='document.af.dwtx.focus();' type=checkbox name=dwtx checked title='�Ƿ�ʹ��ϵͳ��Ч����(�磺�𶯡����ֵ�)��' value="ON">                                                                          
            <a href='#' onClick='document.af.dwtx.checked=!(document.af.dwtx.checked);document.af.dwtx.focus();' title='�Ƿ�ʹ��ϵͳ��Ч����(�磺�𶯡����ֵ�)��'>��/��</a>                                                                           
            <input type=checkbox name='towhoway' value='1' accesskey='s' <%if sjjh_jhdj<=13 or (sjjh_grade>5 and sjjh_grade<9) then%>disabled<%end if%> onClick="document.af.sytemp.focus();">
            <a href='#' onClick='<%if sjjh_jhdj<=13 or (sjjh_grade>5 and sjjh_grade<9) then%>{alert("�ܱ�Ǹ���˹���13�����º͹ٸ������ã�")}return false;<%else%>document.af.towhoway.checked=!(document.af.towhoway.checked);<%end if%>document.af.sytemp.focus();' title='���Ļ�������˵�����������º͹ٸ�������'>˽��</a>                       
            <input type='checkbox' name=as accesskey='a' checked onClick='parent.f1.scrollit();document.af.sytemp.focus();' value="ON">                                                                          
            <a href='#' onClick='document.af.as.checked=!(document.af.as.checked);                                                                          
           document.af.as.focus();parent.f1.scrollit();document.af.sytemp.focus();' title='�Ƿ������ʾ'>����</a>
        <input type='checkbox' name='gbbox' onClick="gBon()">                                                  
        <a href='#' onMouseOver="window.status='24Сʱ�㲥';return true" onMouseOut="window.status='';return true" onClick="document.af.gbbox.click()" title="��24Сʱ�Ĺ㲥��ˬŶ">�㲥</a>       
            <a href="f2/xp.asp?name=<%=nikename%>" target="f3" title="���˿ˡ����齫">[��]</a> <a href='dushu.asp' onClick="javascript:s()" onMouseOver="window.status='С�ĵ�Ŷ,��Ҫ��ҵ�����^_^';return true" onMouseOut="window.status='';return true" target="f3")" title="С�ĵ�Ŷ,��Ҫ��ҵ�����^_^">[��]</a>
            <%if sjjh_grade>=6 then%>                                                                          
             
            <a href="../dt/ask.asp" target="_blank" title="���⣬�ñ���������">����</a>                                                                           
            <a class=blue href="#" onClick="window.open('../dalie/ZJ.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="���ϻ����ñ�������">��</a>                                                                           
            <a class=blue href="#" onClick="window.open('../jiaofei/ZJ.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="�����ˣ��ñ�������">��</a>
           
            <%else%>                                                                                                                                                  
            <a class=blue href="#" onClick="window.open('../dt/reply.asp','daliwu','scrollbars=no,resizable=no,width=490,height=600')" title="������ɣ����������׬ร�">����</a>                                                                           
            <a class=blue href="#" onClick="window.open('../dalie/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="ֻҪ�ٸ����ϻ��ˣ�����Դ������Ǯ��">��</a>                                                                          
            <a class=blue href="#" onClick="window.open('../jiaofei/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="ֻҪ�ٸ��������ˣ�����Դ�˵�Ǯ��">��</a>                                                                           
  
            <%end if%>                                                                          
  
            <%if chatinfo(7)=0 then%>                                                                          
            <a href=lg.asp title="��������ʱ�䣬��������ǲ��ܴ�ܵģ�" target="d">����</a>                                                                           
            <%end if%> 

            <%if sjjh_grade>=9 then%>     
            <input type='checkbox' name='slbox' accesskey='s' onClick="document.af.sytemp.focus();if (parent.slbox==1){parent.slbox=0}else{parent.slbox=1}" value="ON">                                                                          
            <a href='#' onMouseOver="window.status='ѡ�б���󣬿��Բ鿴���������е�˽�ģ���';return true" onMouseOut="window.status='';return true" onClick="document.af.slbox.checked=!(document.af.slbox.checked);if (parent.slbox==1){parent.slbox=0}else{parent.slbox=1};document.af.sytemp.focus();" title="��һ�������˵һЩʲô���Ļ�">��</a>                                                                          
            <%end if%>                                     
             
                                      
            <script language="JavaScript">function startnosay(){var nosay=<%=Application("sjjh_maxtimeout")*60%>;document.af.clock.value=nosay;}</script>                                                                          
            <script src="readonly/f2.js"></script>                                                                          
            <script>parent.fc();parent.m.location.href="f3.asp";</script>                                     
                                              
            <div id="dh" style="position:absolute; left:-80px; top:-80px; width:0px; height:0px; z-index:1">  
             
 
              <input type="button" value="<" name='gopp' onClick="Javascript:gop();" accesskey=","> 
              <input type="button" value=">" name='gonn' onClick="Javascript:gon();" accesskey="."> 
 
            </div> 
          </div> 
        </td> 
      </tr> 
    </table> 
    </form>
</body></html>