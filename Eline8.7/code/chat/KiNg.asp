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
<html><head><META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<style type=text/css>
body {font-family:verdana,arial,helvetica,Tahoma; CURSOR: url('yinsu.ani');font-size: 9pt; color:'ffff00';
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff

	}

td	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
div 	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
form	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
select	{font-size: 9pt; background-color:D7EBff}
option	{font-size: 9pt; background-color:D7EBff}
	
p	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
br	{font-family:verdana,arial,helvetica,Tahoma; font-size: 9pt}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'ffff77'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff4d4d;  font-size: 9pt}
.U1	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 10px; }
.U2	{font-family: Geneva, Verdana, Arial, Helvetica; font-size: 11px; }


.informat01{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color:'F7DEAC'}
	
.i02	{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; background-color: 'F7DEAC'; 
	color: '291C03'; border: 1 solid '291C03'}

.i03{ background-color:'F7DEAC'; 
	BORDER-TOP-WIDTH: 1px; 
	PADDING-CENTER: 1px; PADDING-CENTER: 1px; 
	BORDER-CENTER-WIDTH: 1px; FONT-SIZE: 9pt; 
	BORDER-LEFT-COLOR:'ffffff'; BORDER-BOTTOM-WIDTH: 1px; 
	BORDER-BOTTOM-COLOR:'EBA82C'; PADDING-BOTTOM: 1px; 
	BORDER-TOP-COLOR:'ffffff'; PADDING-TOP: 1px; 
	HEIGHT: 20px; BORDER-CENTER-WIDTH: 1px; 
	BORDER-CENTER-COLOR:'EBA82C'}

.i04{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'ffffff'; color: '333333'; 
	border: 333333px solid; background=ffffff; } 
	
.i05{font-family:verdana,arial,helvetica,Tahoma; font-size:9pt; 
	background-color: 'F3CB81'; color: '83590C'; 
	border: 1px solid AB7510;} 
.i06{ background-color:'F7DEAC'; 
	border-top-width: 1px; 
	padding-right: 1px; padding-left: 1px; 
	border-left-width: 1px; font-size: 9pt; font-family: verdana,arial,helvetica; 
	border-left-color:'ffffff'; border-bottom-width: 1px; 
	border-bottom-color:'EBA82C'; padding-bottom: 1px; 
	border-top-color:'ffffff'; padding-top: 1px; 
	border-right-width: 1px; 
	border-right-color:'EBA82C'}
</style>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}
function rc(list){if(list!="0"){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}}
</script>

</head>
<body bgcolor="#006699" topmargin="3" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<div align=left class=yy4>

</div>
<div align="center">
  <center>
    <table border="1" width="95%" bgcolor="#006699" bordercolorlight="#000000" bordercolordark="#FFFFFF" cellspacing="0" cellpadding="2">
      <tr>
        <td width="100%" align="center"><font color="#CCCCFF"><strong>���ù���</strong></font></td>
  </tr>
  <tr>
    <td width="100%" align="center"><select name='addsign' onChange="rc(this.value);" style='font-size:12px'>
<option value="0" selected>��������</option>
<option value="//������ŵ��ﵽ%%�����ͳ�һ�Ѵ��ӣ��ٺ�......">͵͵</option>
<option value="//����%%�ݺݵ�һ����ͷ����[img]dz22.gif[/img]΢Ц������%%������ҷ���ȥ�ɣ���">����</option>
<option value="//����������һ��[img]dz33.GIF[/img]��%%��Ϊһ�ѻҽ���">�ҽ�</option>
<option value="//���Ͱ͵ظ�%%��Ǹ��ʵ���ǶԲ�ס�����´���Ҳ�����ˡ���">��Ǹ</option>
<option value="//��%%��а���Ц��[img]dz53.gif[/img]�˳�ͬʱ�뵽ʲô����ͷ�ϡ�">����</option>
<option value="//��%%С���ֹ�[img]dz51.gif[/img]�������ӱ���ʮ�겻��������������㡣��">����</option>
<option value="//����%%͵͵ֱ�֣������ֻ�����[img]dz40.gif[/img]��á�">���</option>
<option value="//Ц�Ǻǵض�%%��ȭ��Ҿ���������´��������׹��������������������ң���">����</option>
<option value="//����ص���%%������[img]dz48.GIF[/img]���ˣ���á�">�к�
<option value="//��������ض�%%˵��������λ��������Ҫ�����ˣ����Ǻ�����ڣ���">�뿪
<option value="//����%%[img]dz19.gif[/img]���ֽкá�">�к�
<option value="//ӭ��������ף���������ˮ���ţ����쳤Х������������Ӣ�ܣ����λ��ּ��ֻأ���">��Х
<option value="//���ɵ�ʮ�������̾����������֮�£�ãã���[img]dz12.gif[/img]����˭����%%���氡����">���
<option value="//��������Ӣ��������˭��[img]dz23.gif[/img]">Ӣ��
<option value="//�ǳ�ͬ��%%�Ľ�����">ͬ��
<option value="//��[img]dz34.gif[/img]���ڳ�ÿ��������ɨ����Ȼ��ͣ��%%�����ʵ����ǲ��Ǹ�������ҥ��������벻Ҫ��Ϲ�����ˡ�">��ҥ
<option value="//��%%���������ú�������ǰ�������ǸϽ���ɡ�[img]dz39.gif[/img]">����
<option value="//����%%���һ����С�����������ӣ�[img]dz27.gif[/img]">����
<option value="//�������ȹ��ϵ����ϣ���ͣ�ش�У���㣡��㣡��ô��[img]pic44.gif[/img]��������">̫��
<option value="//�ŵö��ڽ���ֱ�������𡢱�[img]dz21.gif[/img]����ң�">����
<option value="//����һ��[img]dz49.gif[/img]͵͵��Ц%%��">͵Ц
<option value="//���Ӻ�ͨ[img]dz620.gif[/img]�ȵ��Ҵ�һ�±�%%��">���
<option value="//���̲�ס��[img]dz621.gif[/img]������������%%��">��ʺ
<option value="//��Һð�[img]dz622.gif[/img]��ð�%%��">HI
<option value="//�����˿��������[img]dz623.gif[/img]�ȵ����ĺ���%%��">С��
<option value="//����ƨ��Ť��Ťȥ[img]dz624.gif[/img]��Ц��%%��">����
<option value="//Ť���˹�[img]dz625.gif[/img]%%:˵����ģ������ˡ�">����
<option value="//����������[img]dz626.gif[/img]�᲻�ᰢ%%��">����
<option value="//��������[img]dz627.gif[/img]��ҪЦ��%%��">����
<option value="//�ҵ������������[img]dz628.gif[/img]%%��">����







<option value="//����%%��[img]dz46.gif[/img]�����ε����ʼ硣">����
<option value="//����%%����һ����[img]dz37.gif[/img]˵����С�����۲�ʶ̩ɽ����������������ܳŴ����ɱ��С��һ���ʶ����">����
<option value="//������Ⱥ[img]dz41.GIF[/img]�����������ң���֧��%%��������">֧��
<option value="//����%%���һ��:�����С����������ƻ���">����
<option value="//���ɻ��[img]dz34.gif[/img]%%��">�ɻ�
<option value="//һ���󺰣�����ɽ���ҿ������������ԣ���Ҫ�Ӵ˹���������·�ƣ��� ">ɽ��
<option value="//��ס%%������[img]dz43.GIF[/img]һ�ѱ���һ�����˵������λ�����������кã���ﰳ�ɣ���">����
<option value="//����[img]dz44.gif[/img]����%%һ���ֵ������¹����ˣ���ұ˴˱˴ˡ�">�˴�
<option value="//�߾���ȭ[img]qt.gif[/img]ҧ���гݵظߺ�������%%����">��
<option value="//�ܿ���%%[img]dz16.gif[/img]��������">����
<option value="//�ܲ�����˼����%%��Ǹ[img]dz37.gif[/img]">��Ǹ
<option value="//��%%���:�����ģ��Ⱦ��Ұ���">����
<option value="//��%%С���ֹ���[img]dz51.gif[/img]�����ӱ���ʮ�겻��������������㡣��">����
<option value="//��������%%һ�ۣ�����һ���߸ߵİ�ͷ�������������������㡭����">����
<option value="//������Ҿ��������ò�ض�%%˵:����%%�����ָ�̣�">����
<option value="//�ճյ�����[img]dz35.gif[/img]%%����Ӱ[img]dz52.gif[/img]�������ˡ���">�ճ�
<option value="//��%%һ���ֵ���ԭ��С�ֵܾ���־��������������Ͷ�ء�">��̾
<option value="//��ָ�ڿ���һ��������%%��������������һ��˫���ë���ӡ�">ë��
<option value="//�������������%%��ת[img]dz36.gif[/img]��֪���ֹ�Щʲô��">��ת
<option value="//�����ض�%%����[img]dz45.gif[/img]����������������ͯ���˵Ķ�������">����
<option value="//��%%�޵��������˳���ʲôʲô֮�У�������ʲôʲô֮�⣬�Ǹ��Ǹ������������Ѽ��������Ѽ�ѽ��">���
<option value="//�����¹���������Ŀ���%%[img]dz15.gif[/img]��˵��������������˵������֡�����">����
<option value="//�ݺ�����%%���������[img]dz02.gif[/img]�������ð���ǡ�">����
<option value="//�뵽����Ȣ%%��������[img]dz101.gif[/img]�˷ܵļ���˯���ž���">��Ȣ
<option value="//��ʼ����ʲôʱ���ܹ��޸�%%������[img]dz14.gif[/img]�˷ܵ�����ͨ�졣">���
<option value="//����%%���������������ţ������㶼����Щʲô������">����
<option value="//��%%�ߵ�������¿���ʱ��������ơ�">��в
<option value="//��%%����ԭ��Ϊֻ��[img]dz32.gif[/img]����ô��·��û�뵽С�ֵ�Ҳϲ�����У�ʧ��ʧ����">����
<option value="//��ৡ���һ�����һ��������������[img]dz24.gif[/img]����ʱȫ�����쬵�ֻ�к������ˡ�">����
<option value="//һ���������ˣ��͵ر�ס%%[img]dz13.gif[/img]һ���ࣶ�~~��">����
<option value="//����һ��[img]dz04.gif[/img]��%%�߳���̫��ϵ��">�߷�
<option value="//��ݺݵ���в����%%�����ٲ���ʵ����[img]dz03.gif[/img]��������³���">�³�
<option value="//��ǰһ��������[img]dz17.gif[/img]��">����
<option value="//�����书[img]dz06.gif[/img]����Ĭ�%%������ţ���һ������ѩ�ޣ�">����
<option value="//��ݺݵض�%%[img]dz07.gif[/img]�������㣡�����㣡��">����
<option value="//������Ƥ�޳���%%[img]dz08.gif[/img]���㲻������">�޴�
<option value="//����%%�����е�[img]dz31.gif[/img]��Ҫѽ��">���
<option value="//�˷ܵĴ���[img]dz20.gif[/img]̩̹��ˣ�">����
<option value="//����%%[img]dz09.gif[/img]�����㡢�����㣡">����
<option value="//�������ؿ���%%��[img]dz26.gif[/img]���ߣ�����ס[img]dz30.gif[/img]">סԺ
<option value="//����%%�ö��ģ����ܲ����ˣ���Ҫ~Ҫ~[img]dz10.gif[/img]">����
<option value="//���ĵ�Ų���Ų�������%%���ԣ�������[img]dz50.gif[/img]">����
<option value="//��%%ʹ���˵��������������Ͱ��Ϲ������������Ҫ��飬������û�����ˡ���">���
 <option value="//�����ؿ��ž�������Լ���ɵ������Ҫ�;�����������֣�˭֪����������˾�Ȼ�����%%��ģ�������������������鰡��">����
 <option value="//�ڽ�����Ȧ֮�ϣ����ֶ��������������������ܵ���Ⱥ����Ⱥ�ͱ�ѹ��Ⱥ��ֻ�������¹���Ӣ��������%%��˭��~~">����
 <option value="//��%%ʹ����һ���л�¥������ȫϯ���Լ��������½�965987458,�����½�5000���õ�965987�ľ���ֵ��ɱ�˶Է�����96598745685��%%�ҽ�һ����������Ѫ�������ĵ�����ȥ......���ڳ�����λ�������޲������л�¥�������ݶ����书������ͬʱҲ����##�����ĺ��ֺ���">ɱ��
<option value="//�������֣�����һЦ�����ŷ��Ƴ��˽ܵ��飬С�ܳ���է����������λ�����㡱��">����
<option value="//������˵����������û�в�ɢ����ϯ��������һ���ˣ���ұ��ء���">���� 
<option value="//���������쳤���Ӷ��ſ��������˼��£� ����%%�����޵������㣬����!">����
<option value="//���һֱ�����Լ�ͷ��Ŀѣ,��������,��֪�ǲ���.....">ͷ��
<option value="//��%%��ע��ʱ��%%������͵͵����һ�£�û�뵽%%�е�������ںó������û�Կ��ˣ�">͵��
<option value="//��������%%�Ļ��������ૣ�����쨶��������һ�">����
<option value="//���һ�����������������,�����Ⱦ��ܣ�Ƭ�̾���ʧ��һƬãã������"">����
<option value="//һ������һ�ѱ���غ���ԩ������%%�������ү!!"">��ԩ
<option value="//Ʈ��Ʈ��һ���飬����Ҳ�з硭����Ȼ���������ڣ������쳾�С��������ˡ���"">����
<option value="//�󱩷���һ�����˹���������˵�ţ�%%������û���㰡��������ϧ������һ����һ�����ڶԷ���Ь�����ϡ�"">����
<option value="//��京�߲�һ��������������ͻȻ�ܻ���������ء���������%%һ�£�ת������ˡ�"">����
<option value="//���������������������������ͻȻ������һ����ͷֱֱ�����·�ߵ�ˮ���"">����
<option value="//��ֻСè�Ƶ�������%%�Ļ���!!"">Сè
<option value="//ͻȻ�γ�һ����ǹ��ָ���Լ���ͷ��������涵Ŀ��Ŵ�ҡ�"">��ɱ
<option value="//����ҧס�Լ����촽���۾���ˮ��һ��������ע���ŶԷ����ÿ��Դݻ�һ�����ң�ɱ��һ������������һ���ĵ�����˵�����������ѣ������ݣ�������ô�᲻���ء�"">�嶯
<option value="//ǿ��ס�����һɲ�ǵ��һ���ģ��ӷ�����嵽��԰�����ϵ�����������������Ӿ���У���Ҿ��ȵķ��֣���һ�����ŵ��һ����Ӿ�ص�ˮ��Ȼȫ�����ˣ�"">�嶯
<option value="//ͻȻ��һ���������ҵĵ���ǧ�˰ٴγ����������㣬%%�������ҵ���һСƬ����ɣ�"">��
<option value="//����ߵķ���һ��ͽ�͵����������£�Ȼ���޿ɾ�ҩ�������%%���������塭��"">����
<option value="//���Լ��ĳ��񼫶Ȳ�����, ������˫���������.">�Ա�
<option value="//�����ޱ�����ơ�Ƹ�%%��Ȼ��˵:������ !��">����
<option value="//�����������ĥ����%%�ķ���, ����˵��: �Һ�ϲ����Ŷ....">����
<option value="//��������%%���°�����ĳ��� [IMG]dz57.GIF[/IMG] ~@~"">����
<option value="//���������:����,��������֮�Ƕ���, ��������֮�ֶ���...����������̫�ס�">����
<option value="//����ĬĬ��������ɰ������ǣ��ɰ��������������ڰ��ߣ��������޲�����">Ĭ��
<option value="//����ͷ�㽵�������������ϰ����й��ʣ���û���ϰ�ܣ��ǲ�֪�����ж�á���">��<option value="//������������Ŀ���飬ֻ���ô���Ļ����ҡ�">����
<option value="//�������е��־���ߵذ�ͷŤ���������ˣ���ʲô������">����
<option value="//����˫Ŀ����������ˮ��ĳ�����лл����ҵİ������������Ҳ�������лл����ҵ����ᣬ���Ҷȹ��Ǹ��������">����
<option value="//��������������������֪�õ�˭�Ҳ䷹ȥ��......��">�䷹
<option value="//˵�������ԹԲ����ˣ����µ��ϰ����ˣ�û�������ˡ��ؼ�����">�ϰ�
<option value="//ü����Ц���룺�˶�˵�����ۣ����ǰ����˼һ᲻�ᣬ������......">����
<option value="//���һƲ����Ц������������å������˭������">��å
<option value="//���ž����������ݵ�������Թ�԰���������##��������࣡��">����
<option value="//һ������ã����֪����Χ������ʲô�¡�">��ã
<option value="//ʹ�����İ��ƣ�žžžžž.....�ĵ��ֶ��������ˣ�[IMG]dz19.gif[/IMG]">����
<option value="//������һ�ɣ����������������Ҫ�����ˣ���Ҫ����ң���">����
<option value="//��ȭ����һ�ݣ��ֺõ�˵���������˶Ը�λ�ľ���֮�飬�������ν�ˮ���಻������">�ֺ�
<option value="//���������ˣ���û�еض��������ȥ��������">����
<option value="//���ظ��ĵñ�ž�죺��������ȭͷ���˵�������ֵ������Ȼ��Ȼ�����">��ţ
<option value="//����˵��������������˲�ǳ�����ս�����꣬�ٵ������Ի������Ǿʹ˱������˵������һ����ƮȻ��ȥ��">�ټ�
<option value="//����ǽ��ɪɪ����������������Ҳ��ã��Ҽ��֣��Ҳ��ԣ��������">����
<option value="//��Ҫ�������������Է������������������Ǹ֣������ټ��ˣ�">����
<option value="//�ν�������:��ʮ��ĥһ����˪��δ���ԡ����հѾ��ʣ�˭�в�ƽ�£���">��ʫ
<option value="//��Ц�������������һ����ͷ�����ţ�����û�й����ˣ��������������塣">����<option value="//ˢ�ĳ��һ����������Ľ䵶��һ�����Լ��Ĵ��ͷ­�������������ﻹ��ͣ���ֹ������ұȿᣬ˭�ܱ��ҿ᣿">�ȿ�
<option value="//ͻȻ����θһ����, �����������������ء�">Ż��
<option value="//���һ������������һ�����й����٣���̾һ����֪�����٣��հհգ�����ʲô��÷�����������Ҵ�ϡ�">����
<option value="//���ճ�������̾������֮�󣬾���һ�����֣���į��������">�޵�
<option value="//�����Լ�һ���⣬������˼��˵���������ǶԲ�������һʱʧ�ԣ���Ҷ���������">��Ǹ
<option>----</option> 
<option value="//����%%��У�����ӭ����ӭ�����һ�ӭ��[IMG]dz63.gif[/IMG]��">��ӭ
<option value="//�������%%���к���������һ����ŨCoffee��һ������Ĳ裬�Լ�ѡŶ��">����
<option value="//�ó�һ��[IMG]dz01.gif[/IMG]ҡͷ���Ե���ض�%%˵����������������졭����">����<option value="//��%%˵����Щ���Ҵ�绰[IMG]dz58.gif[/IMG]�����򡭡���">�绰
<option value="//��%%˵���������ˣ��������ˣ���һ����Ҳ���������ˣ���������[IMG]dz16.gif[/IMG]">����
<option value="//��һ����֦ͱ��ͱ%%�ı��ӣ�������ѽ�������ѡ���">����
<option value="//���ž��е�%%��ֻ��%%��ɫ糺죬�۲���ת������һЦ�����ľ������˲��ɵù�����֣����ſ��ţ�##���ɵó��ˡ�">�ɰ�
<option value="//����%%�����룺���ڵ�Ů������������������[IMG]dz12.gif[/IMG">����
<option value="//����Ķ�%%˵������������[IMG]dz11.gif[/IMG]������ɢŶ~~~~~~��">����
<option value="//��������%%��ͷ�����ǽ���~!~~~~~[IMG]dz15.gif[/IMG]">����
<option value="//����~~~~~~~~~~[IMG]dz17.gif[/IMG]">�ε�
<option value="//ף%%���տ���[IMG]dz59.gif[/IMG]">����
<option value="//����%%����[IMG]dz13.gif[/IMG]˵���޸��Һ����װ��ģ��ҵ���10����~~~~">�ʻ�
<option value="//��ͷһ˦��û�п����һ�ۣ�[IMG]dz16.gif[/IMG]�߷ߵ�˵����������������������">����
<option value="//�ֺǺǵض�%%�У�������å����˭������[IMG]33.gif[/IMG]">��å</option>
<option value="//һ��Բ����������������Ű���%%��������˰��ɣ������[IMG]dz21.gif[/IMG]">����<option value="//��%%�������⣬��æץ��绰��114��119��120��122��96315 ... ���룺������������ô������ѽ����">����
<option value="//��������%%һ�ۣ�����һ�����߸ߵİ�ͷ�����������������㡭��">����
<option value="//����%%�����ž��½�����ȭ[IMG]dz06.gif[/IMG]">ȭ��
<option value="//��%%���ڵ��¡���[IMG]dz07.gif[/IMG]">�ȱ�
<option value="//�����ˡ�ɱ���ơ���һ�ư�%%�����������ˡ�[IMG]dz05.gif[/IMG]">����
<option value="//���ĵذ�һ����˪����%%���������һ�ԣ�ϲ���̵ؿ�����������ȥ��>�¶�
<option value="//��%%�¸ҵع�������������Ը��޸����𡱩������������ɼΰ������ֹ���[IMG]8.gif[/IMG]">���
<option value="//��žžž��������%%�������⣬����ͨ��˵��ǿ����[IMG]dz02.gif[/IMG]��">����
<option value="//����%%�����������Ǵ��Ƶ��ֵܽ��þȾ��Ұ�[IMG]dz03.gif[/IMG]��">�Ⱦ�
<option value="//���������%%���������������˵�����Һ�ϲ����Ŷ......��">ϲ��
<option value="//ץ��%%��������һ�ţ���Ź�������ӣ�����һ�����һ������ô���˿�±����!!!!">±��
<option value="//��%%�첻���ˣ���æ����[IMG]dz60.gif[/IMG]����%%�ͽ�����Ժ����">�Ȼ�
<option value="//���������%%һ�£��������ϡ���[IMG]dz57.GIF[/IMG]">����
<option value="//%%ͻȻ�䱻����ס��˫�ۣ�����һ����##վ����ǰ�����Ϸ��źô�һ��õ��[img]dz61.gif[/img]��(">õ�� 
<option value="//���˵�������%%��С������������һ�����е�����˵������%%С��ô���Ŷ���������Ҿͷ�����������Ҫע��СС����ͱ�̫����ඡ���%%�����佱���ٺ�ɵЦ������">����
<option value="//�����§��%%�����������ĵ����������µ��գ��㿴���϶�����Ҳ��Ľ����Ү����">��§
<option value="//��%%˵�������������������Ļ�����Ȼ�����죬�Ĳ��������ǹ�����񣡲�֪��������������ɷ�ѧ�����������һ����">Ƥ��
<option value="//�Ͻ���ס��������%%���ĵ�����С�ֵ�����ǧ������ˣ���������춣��ѽ���������������Ĺĵģ���ʵ����ɶ�����Ҷ��ѽ����">����
<option value="//����ض�%%˵������һ��������ʮ���죬һ���ʮ��Сʱ��ÿ��ÿ���Ҷ������������......��">����
<option value="//����%%��ȥ�ı�Ӱ����ˮֹ��ס�������£����͵�������֮�ģ���֮�ǣ�֪�������䡣һư�Ǿƾ��໶���������κ�����">���
<option value="//�������أ���縳���Ʈ�����У����������Ƚ����������˼��㣬��֪����%%��һ�ŷ���������ϵ����##�����ϡ�">����
<option value="//����Բ�������������Ķ���%%˵���������һʵ۶���������̫�༱��ƨ�á�������">�ż�<option value="//��������%%�ı���һ������С�ӣ��ɵĺá��ı��顣">�ı�
<option value="//�����ӵ�Ŀ�ⶢ��%%˵��������ƾ��.....Ҳ�䣿��������">����
<option value="//�Ӷ���������һ����Ͱ͵��̾���ݸ�%%����һ������">����
<option value="//����˵������%%����ͬ־�ˣ���Ȼ�����е��Ϳ�ˣ����Ǵ�һ���Ӧ���������">��
<option value="//���п���������ˣ�����%%˵��:����ɽ������ˮ���������ӱ���ʮ�겻��������">����
<option value="//§ס%%��������ָ�ų���������˵�����������£������ǵ�֤�ˡ���">֤��
<option value="//����%%���еĵ�������̾������ĵ��������������Ҳ��������">����
<option value="//��%%�����е�����ʲô�£������Դ��������Ị̇���С���е��͸���ƴһ����">Ӣ��<option value="//����%%ĬĬԶ�ߵı�Ӱ���������ڷ���Ʈɢ�ĳ������ǹ¼ŵı�Ӱ����ˮֹ��ס���䵽�ؿڣ����к����š�%%����ʵ����İ��㡱��">�뿪
<option value="//��Ȼ̾���������ڽ�����������һ��ã������³�����Ҫ��%%һ�������˶�����֪����˭�µ��֣���">С��

</select>
</td>
  </tr>
  <tr>
    <td width="100%" align="center">
      <table border="0" width="12%">
        <tr>
              <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a  href="#" onClick="window.open('../hcjs/jhzb/index.asp','wupin','scrollbars=yes,resizable=yes,width=550,height=300')" title="��Ҫʹ��ҩƷ��װ����Ʒʱ�����"><font size="-1"><img src="img/gn/wupin.gif" width="30" height="19" border="0"></font></a></td>
              <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../bbs/index.asp" target="_blank"><img src="img/gn/luntan.gif" width="30" height="19" border="0" alt="������Ϣ|���ѽ���"></a></td>        </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#" onClick="window.open('../hcjs/jhjs/dan.asp')" ><img src="img/gn/dangpu.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../hcjs/jhjs/yaopu.asp')" title="������ҽ����������"><img src="img/gn/yaodian.gif" width="30" height="19" border="0"></a></td>        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../hcjs/jhjs/wuqi.asp')" title="����������"><img src="img/gn/bingqi.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../hcjs/card/card.asp"  target="_blank" title="��Ա��Ƭ��"><img src="img/gn/kapin.gif" width="30" height="19" border="0"></a></td>        </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../yamen/laofan.asp')" title="�����㻵��Ŷ...��Ȼ...�ٺ�..."><img src="img/gn/laofang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="#"onClick="window.open('../yamen/minan.asp')" title="��������"><img src="img/gn/mingan.gif" width="30" height="19" border="0"></a></td></tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" nowrap width="3%"><a href="#"onClick="window.open('../hcjs/yilao.asp')" title="���۶��ص��ˣ�ҽ�ɺ���ţ����������������"><img src="img/gn/liaoshang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a href="work.asp" target="f3"  title="�ҵ�ְҵ-�ҵĹ���"><font size="-1"><img src="img/gn/gongzuo.gif" width="30" height="19" border="0"></font></a></td></tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="3%"><a href="#" onClick="window.open('../yamen/regmodi.asp','zhuangti','scrollbars=yes,resizable=no,width=450,height=290')" title="���������޸�"><font size="-1"><img src="img/gn/ziliao.gif" width="30" height="19" border="0"></font></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" height="2" width="1%"><a href="../Diary/Diary.asp" target="_blank" title="ǩд�����ռ�"><img src="img/gn/rji.gif" width="30" height="19" border="0"></a></td>      </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jl/jlmain.asp" target="_blank" title="����ÿ�췢���Ĵ���"><img src="img/gn/dashi.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jhjd/jd.asp" target="_blank" title="E�ߴ�Ƶ�"><img src="img/gn/jiudian.gif" width="30" height="19" border="0"></a></td>
        </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../hcjs/xuewu/xuetang.htm" target="_blank" title="�������"><img src="img/gn/wuguan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../jhmp/wuguan.asp" target="_blank" title="����������...�ð�ʦ�����ܽ���"><img src="img/gn/mishi.gif" width="30" height="19" border="0"></a></td>
          </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../top/top.htm" target="_blank" title="������������"><img src="img/gn/paihang.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../gupiao/stock.asp" target="_blank" title="�������ɣ��뷢������������"><img src="img/gn/gupiao.gif" width="30" height="19" border="0"></a></td>
                  </tr>
        <tr>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../dk/Daikuan.asp" target="_blank" title="���ڲ���Ǯ�ᱻɱ��׷ɱŶ"><img src="img/gn/daikuan.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../bet/betindex.asp" target="_blank" title="�����Ĺ�"><img src="img/gn/duguan.gif" width="30" height="19" border="0"></a></td>
        </tr>
        <tr>
            <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="12%"><a href="../jhmp/index.asp" target="_blank" title="���������Ҹ����ϴ��չ�׼û��"><img src="img/gn/menpai.gif" width="30" height="19" border="0"></a></td>
    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="1%"><a href="../dg/dg.asp" target="_blank" title="��׬Ǯ"><img src="img/gn/dagong.gif" width="30" height="19" border="0"></a></td>
        </tr>
        <tr>

    <td align="center" onMouseOut="this.bgColor='';"onMouseOver="this.bgColor='#8B851C';" valign="middle" width="58%"><a href="../jhmp/gry.asp" target="_blank" title="�������Ƶĺ��£������ӵ���Ŷ"><img src="img/gn/guer.gif" width="30" height="19" border="0"></a></td>
          <td></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
      </center>
    </div>
