<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
nosaytime=Application("Ba_jxqy_nosaytime")
username=session("Ba_jxqy_username")
chatroomsn=session("Ba_jxqy_userchatroomsn")
chatroomname=Application("Ba_jxqy_systemname"&chatroomsn)
%>
<head>
<link rel="stylesheet" href="css.css">
<script language=javascript>
if(window==window.top){top.location.href="chatroom.asp";}
var maxlingual=19;
var pos=0;
var lingualnum=0;
var i;
var lingualarr=new Array(maxlingual);
var nosaytime=<%=nosaytime%>;
for(i=0;i<maxlingual;i++){lingualarr[i]='';}
function checklingual(){if(document.talkform.talkmsg.value==''){alert('��ѡ�����Ի�����');document.talkform.talkmsg.focus();return false;}if (lingualnum<maxlingual+1){lingualarr[lingualnum]=document.talkform.talkmsg.value;lingualnum++;}else{for (i=0;i<maxlingual;i++)lingualarr[i]=lingualarr[i+1];lingualarr[i]=document.talkform.talkmsg.value;}pos=lingualnum;{document.talkform.talkmsgtmp.value=document.talkform.talkmsg.value;document.talkform.talkmsg.value='';document.talkform.allownosaytime.value=nosaytime;document.talkform.subm.disabled=true;setTimeout('document.talkform.subm.disabled=false;',3000);document.talkform.talkmsg.focus();document.talkform.filter.value='';document.talkform.action.value='';return true;}return false;}
function goback(){if(pos>0)pos--;document.talkform.talkmsg.value=lingualarr[pos];document.talkform.talkmsg.focus();}
function goforw(){if(pos<lingualnum-1)pos++;document.talkform.talkmsg.value=lingualarr[pos];document.talkform.talkmsg.focus();}
function nosay(){
	document.talkform.allownosaytime.value=document.talkform.allownosaytime.value-1;
	if (document.talkform.allownosaytime.value < nosaytime - 120)
	{
		autoSay();
	}
	setTimeout('nosay()',1000);if(document.talkform.allownosaytime.value==300){window.open('goback.asp','goback','Status=yes,scrollbars=no,resizable=no');}if(document.talkform.allownosaytime.value<0){top.location.href='chaterror.asp?id=001';}}
function init(){var rndcolor=Math.floor(Math.random()*18);document.talkform.namecolor.value=document.talkform.namecolor.options[rndcolor].value;rndcolor=Math.floor(Math.random()*18);document.talkform.wordcolor.value=document.talkform.wordcolor.options[rndcolor].value;parent.msgfrm0.document.open();parent.msgfrm0.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css.css'></head><\Script Language=JavaScript>var autoScroll=1;function chgautoscroll(){if(!parent.talkfrm.document.talkform.autoscroll.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}function autoscrollnow(){if(autoScroll==1){this.scroll(0,65000);parent.msgfrm1.window.scroll(0,65000);setTimeout('autoscrollnow()',200);}}autoscrollnow();<\/script><body text=000000 oncontextmenu=self.event.returnValue=false><font color=FF0000>����������¡�</font>��ӭ<font color=FF0000>��<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>��</font>����<%=chatroomname%><font class=timsty><%=time()%></font><br>");parent.msgfrm1.document.open();parent.msgfrm1.document.writeln("<html><head><meta http-equiv=Content-Type content='text/html; charset=gb2312'><link rel='stylesheet' href='css.css'></head><body text=000000 ><font color=FF0000>����������¡�</font>��ӭ<font color=FF0000>��<a href='javascript:parent.chgsendto(\"<%=username%>\");' target='talkfrm' onmouseover=\"window.status='ѡ��˵����������';return true;\" onmouseout=\"window.status='';return true;\"><font color=FF0000><%=username%></font></a>��</font>����<%=chatroomname%><font class=timsty><%=time()%></font><br>");parent.getfrm.location.replace('getmsg.asp');parent.chgfrm.location.replace('selectchat.asp');nosay();}
function settalk(str1,str2){document.talkform.talkmsg.value=str1+' '+str2;document.talkform.talkmsg.focus();}
function addtalk(str1){document.talkform.talkmsg.value=document.talkform.talkmsg.value+str1+' ';}
function autoSayclick() {
	document.talkform.autosay.checked=!(document.talkform.autosay.checked);
	document.talkform.talkmsg.focus();
	if (document.talkform.autosay.checked == true)
	{
		document.talkform.talkmsg.value='�㿪���Զ��ݵ㹦��';
		document.talkform.subm.click();
	}
	else {
		document.talkform.talkmsg.value='';
	}
	
}
function autoSay() {
	if (document.talkform.autosay.checked == true)
	{
		var chars = [
			"/#Ц�ǺǵĶ�%%˵���ҿ����Զ��ݵ㹦��������Ҳ�����ĵ�����",
			"/#��%%��Цһ����һ�������֣�����˭��",
			"/#�����Ц�Ķ�%%˵���ҿ����Զ��ݵ㹦��������ѡ�ײ��ġ��Զ��ݵ㡯�Ϳ������޹һ���"
		];
		var id = Math.ceil(Math.random()*4)-1;		
		if (id >=3 || id <=0)
		{
			id = 0;
		}
		document.talkform.talkmsg.value=chars[id] + ",," + id;
		document.talkform.subm.click();
	}
}
</script>
</head>
<body oncontextmenu='self.event.returnValue=false' background="../images/bg.gif" bgcolor="#FFE4CA" onload='init();' leftmargin="5" topmargin="7" marginwidth="5" marginheight="5">
<form name=talkform action="sendmsg.asp" method=post target="sendfrm" onsubmit="javascript:return(checklingual());">
  <input type=button name="lrclutch" value="�����ܲ˵�"  onclick="javascript:parent.lrclutch();" title="��/�ر���߿�" onmouseover="window.status='��/�ر���߿�';return true" onMouseOut="window.status='';return true">
		<select name=namecolor onchange="javascript:document.talkform.talkmsg.focus();">
			<option style="background-color:000000;color:000000" value="000000">����</option>
			<option style="background-color:000088;color:000088" value="000088">����</option>
			<option style="background-color:0000ff;color:0000ff" value="0000ff">����</option>
			<option style="background-color:008800;color:008800" value="008800">����</option>
			<option style="background-color:008888;color:008888" value="008888">����</option>
			<option style="background-color:0088ff;color:0088ff" value="0088ff">����</option>
			<option style="background-color:004400;color:004400" value="004400">����</option>
			<option style="background-color:004488;color:004488" value="004488">����</option>
			<option style="background-color:0044ff;color:0044ff" value="0044ff">����</option>
			<option style="background-color:880000;color:880000" value="880000">����</option>
			<option style="background-color:880088;color:880088" value="880088">����</option>
			<option style="background-color:8800ff;color:8800ff" value="8800ff">����</option>
			<option style="background-color:888800;color:888800" value="888800">����</option>
			<option style="background-color:888888;color:888888" value="888888">����</option>
			<option style="background-color:8888ff;color:8888ff" value="8888ff">����</option>
			<option style="background-color:884400;color:884400" value="884400">����</option>
			<option style="background-color:884488;color:884488" value="884488">����</option>
			<option style="background-color:8844ff;color:8844ff" value="8844ff">����</option>		</select><input type="hidden" name="talkmsgtmp" value="">
			<input type=text name="talkmsg" value="" size=56 maxlength=100 class=norstyle>
		<a href="#" onclick="javascript:goback();" title="ȡ��ǰһ�仰������ʮ��" onmouseover="window.status='ȡ��ǰһ�仰������ʮ�䡣';return true" onMouseOut="window.status='';return true"><font face=webdings>7</font></a>
		<a href="#" onclick="javascript:goforw();" title="ȡ�ú�һ�仰������ʮ��" onmouseover="window.status='ȡ�ú�һ�仰������ʮ�䡣';return true" onMouseOut="window.status='';return true"><font face=webdings>8</font></a>
		<input name="subm"type=submit value="����" class=norstyle>
		<input type="button" name=allownosaytime value="<%=nosaytime%>" readonly class=norstyle disabled><br>
		
  <input type=button name="tbclutch" value="�ط�������" onclick="javascript:parent.tbclutch();" title="����/�����л�" onmouseover="window.status='����/�����л���';return true" onMouseOut="window.status='';return true">
		<select name=wordcolor onchange="javascript:document.talkform.talkmsg.focus();">
    <option style="background-color:000000;color:000000" value="000000" selected>˵��</option>
    <option style="background-color:000088;color:000088" value="000088">˵��</option>
    <option style="background-color:0000ff;color:0000ff" value="0000ff">˵��</option>
    <option style="background-color:008800;color:008800" value="008800">˵��</option>
    <option style="background-color:008888;color:008888" value="008888">˵��</option>
    <option style="background-color:0088ff;color:0088ff" value="0088ff">˵��</option>
    <option style="background-color:004400;color:004400" value="004400">˵��</option>
    <option style="background-color:004488;color:004488" value="004488">˵��</option>
    <option style="background-color:0044ff;color:0044ff" value="0044ff">˵��</option>
    <option style="background-color:880000;color:880000" value="880000">˵��</option>
    <option style="background-color:880088;color:880088" value="880088">˵��</option>
    <option style="background-color:8800ff;color:8800ff" value="8800ff">˵��</option>
    <option style="background-color:888800;color:888800" value="888800">˵��</option>
    <option style="background-color:888888;color:888888" value="888888">˵��</option>
    <option style="background-color:8888ff;color:8888ff" value="8888ff">˵��</option>
    <option style="background-color:884400;color:884400" value="884400">˵��</option>
    <option style="background-color:884488;color:884488" value="884488">˵��</option>
    <option style="background-color:8844ff;color:8844ff" value="8844ff">˵��</option>
  </select><select name=sendto class=norstyle onclick="javascript:parent.chgsendto('���');">
			<option value='���'>���</option>
		</select><select name=action class=norstyle onchange="javascript:document.talkform.talkmsg.value=this.value;document.talkform.talkmsg.focus();">
		<option value="/#������ŵ��ﵽ%%�����ͳ�һ�Ѵ��ӣ��ٺ�......">͵͵</option>
        <option value="/#����%%�ݺݵ�һ����ͷ���£�΢Ц������%%������ҷ���ȥ�ɣ���">����</option>
        <option value="/#����������һ�������%%��Ϊһ�ѻҽ���">�ҽ�</option>
        <option value="/#���Ͱ͵ظ�%%��Ǹ����ʵ���ǶԲ�ס�����´���Ҳ�����ˡ���">��Ǹ</option>
        <option value="/#��%%��а���Ц�ţ��˳�ͬʱ�뵽ʲô����ͷ�ϡ�">����</option>
        <option value="/#��%%С���ֹ��ţ������ӱ���ʮ�겻��������������㡣��">����</option>
        <option value="/#����%%͵͵ֱ�֣������ֻ����룺��á�">���</option>
        <option value="/#Ц�Ǻǵض�%%��ȭ��Ҿ�����������´��������׹��������������������ң���">����</option>
        <option value="/#����ص���%%���������ˣ���á�">�к�
		<option value="/#��������ض�%%˵��������λ��������Ҫ�����ˣ����Ǻ�����ڣ���">�뿪
		<option value="/#����%%���ֽкá�">�к�
		<option value="/#ӭ��������ף���������ˮ���ţ����쳤Х��������������Ӣ�ܣ����λ��ּ��ֻأ���">��Х
		<option value="/#���ɵ�ʮ�������̾����������֮�£�ãã��񷣬����˭����%%���氡����">���
		<option value="/#��������Ӣ��������˭��">Ӣ��
		<option value="/#�ǳ�ͬ��%%�Ľ�����">ͬ��
		<option value="/#���۹���ڳ�ÿ��������ɨ����Ȼ��ͣ��%%���ϣ��ʵ����ǲ��Ǹ�������ҥ��������벻Ҫ��Ϲ�����ˡ�">��ҥ
		<option value="/#��%%���������ú�������ǰ�������ǸϽ���ɡ�">����
		<option value="/#����%%���һ����С���˵��������ӣ�">����
		<option value="/#�������ȹ��ϵ����ϣ���ͣ�ش�У���㣡��㣡�������������">̫��
		<option value="/#�ŵö��ڽ���ֱ������">����
		<option value="/#����һ��͵͵��Ц%%��">͵Ц
		<option value="/#����%%�������ε����ʼ硣">����
		<option value="" selected style="color:0000FF">����</option>
		<option value="/#����%%����һ����˵����С�����۲�ʶ̩ɽ����������������ܳŴ����ɱ��С��һ���ʶ����">����
		<option value="/#������Ⱥ�������������ң���֧��%%��������">֧��
		<option value="/#����%%���һ���������С���˵�������ƻ���">����
		<option value="/#���ɻ�ؿ���%%��">�ɻ�
		<option value="/#һ���󺰣�����ɽ���ҿ������������ԣ���Ҫ�Ӵ˹���������·�ƣ��� ">ɽ��
		<option value="/#��ס%%�����ȣ�һ�ѱ���һ�����˵������λ�����������кã���ﰳ�ɣ���">����
		<option value="/#������Ц����%%һ���ֵ������¹����ˣ���ұ˴˱˴ˡ�">�˴�
		<option value="/#�߾���ȭ��ҧ���гݵظߺ�������%%����">��
		<option value="/#����ô�ô��������%%ͷ������һ�ã�***���ϣ���*** %%������ͣ��������������ݷ𿴵�������****5000000000Pt****">����
		<option value="/#�������ؿ���%%��">����
		<option value="/#�ܲ�����˼����%%������Ǹ��">��Ǹ
		<option value="/#��%%��У������ģ��Ⱦ��Ұ���">����
		<option value="/#��%%��а���Ц�ţ��˳�ͬʱ�뵽ʲô����ͷ�ϡ�">����
		<option value="/#��%%С���ֹ��ţ������ӱ���ʮ�겻��������������㡣��">����
		<option value="/#��������%%һ�ۣ�����һ�����߸ߵİ�ͷ�������������������㡭����">����
		<option value="/#������Ҿ��������ò�ض�%%˵������##�����ָ�̣�">����
		<option value="/#�ճյ�����%%����Ӱ���������ˡ���">�ճ�
		<option value="/#��%%һ���ֵ���ԭ��С�ֵܾ���־�����¡���������Ͷ�ء�">��̾
		<option value="/#��ָ���ǵ�һ��������%%��������������һ��˫���ë���ӡ�">ë��
		<option value="/#�������������%%��ת������ﲻ֪���ֹ�Щʲô��">��ת
		<option value="/#�����ض�%%����������������������ͯ���˵Ķ�������">����
		<option value="/#��%%�޵��������˳���ʲôʲô֮�У�������ʲôʲô֮�⣬�Ǹ��Ǹ������������Ѽ��������Ѽ�ѽ��">���
		<option value="/#�����¹���������Ŀ���%%����˵��������������˵������֡�����">����
		<option value="/#�ݺ�����%%��������⣬�������ð���ǡ�">����
		<option value="/#�뵽����Ȣ%%�������ˣ��˷ܵļ���˯���ž���">��Ȣ
		<option value="/#��ʼ����ʲôʱ���ܹ��޸�%%�������˷ܵ�����ͨ�졣">���
		<option value="/#����%%���������������ţ������㶼����Щʲô������">����
		<option value="/#��%%�ߵ���������¿���ʱ��������ơ�">��в
		<option value="/#��ݺݵ���в����%%�����С�����������ұ��죡">����
		<option value="/#��%%˵����ԭ��Ϊֻ���з����ô��·��û�뵽С�ֵ�Ҳϲ�����У�ʧ��ʧ����">����
		<option value="/#��ৡ���һ�����һ�������������У���ʱȫ�����쬵�ֻ�к������ˡ�">����
		</select><select name=expression class=norstyle onchange="javascript:document.talkform.talkmsg.focus();">
			<option value="" style="color:0000FF" selected>����</option>
			<option value="΢΢Ц��">΢Ц</option>
			<option value="������Ц��">��Ц</option>
			<option value="�ٺٵؼ�Ц��">��Ц</option>
			<option value="��Ȼ���˵�">����</option>
			<option value="������">����</option>
			<option value="�����ϴǵ�">����</option>
			<option value="����������">����</option>
			<option value="ȭ����ߵ�">ȭ��</option>
			<option value="�Ҹ���">�Ҹ�</option>
			<option value="ʮ����������">����</option>
			<option value="��ƤЦ����">��Ƥ</option>
			<option value="�����">����</option>
		</select><select name="filter" class=norstyle onchange="javascript:document.talkform.talkmsg.value=this.value;document.talkform.talkmsg.focus();">
		<option value="" style="color:0000FF" selected>����</option>
		<option value="//��� ">���</option>
		<option value="//�鿴 ">�鿴</option>
		<option value="//���� ">����</option>
		<option value="//���� ">����</option>
		<option value="//���� ">����</option>
		<option value="//��Ǯ ">��Ǯ</option>
		<option value="//���� ">����</option>
		<option value="//���� ">����</option>
		<option value="//���� ">����</option>
		<option value="//��Ǯ ">��Ǯ</option>
		<option value="//͵�� ">͵��</option>
		<option value="//Ͷ�� ">Ͷ��</option>
		<option value="//���� ">����</option>
		<option value="//���� ">����</option>
		<option value="//��� ">���</option>
		</select><input type=checkbox name="isprivacy" class=norstyle>
  <a href="#" onmouseover="window.status='���Ļ����Ķ���˵��������������';return true;" onmouseout="window.status='';return true;" onclick="javascript:document.talkform.isprivacy.checked=!(document.talkform.isprivacy.checked);document.talkform.talkmsg.focus();">˽��</a> 
  <input type=checkbox name="autoscroll" class=norstyle checked onclick="javascript:document.talkform.talkmsg.focus();parent.msgfrm0.chgautoscroll();">
  <a href="#" onmouseover="window.status='�Զ�/�ֶ��������ã�';return true;" onmouseout="window.status='';return true;" onclick="javascript:document.talkform.autoscroll.checked=!(document.talkform.autoscroll.checked);document.talkform.talkmsg.focus();parent.msgfrm0.chgautoscroll();">����</a> 

  <input type=checkbox name="autosay" class=norstyle checked onclick="javascript:autoSayclick()">
  <a href="#" onmouseover="window.status='�Զ�/�ֶ��������ã�';return true;" onmouseout="window.status='';return true;" onclick="javascript:autoSayclick()">�Զ��ݵ�</a> 
  
  <font color="#FF0000">[<b><a href=# onClick="javascript:window.open('../readme.htm','readme',' width=550,height=500,left=200,top=100,status=no,toolbars=yes,menubars=yes,scrollbars=yes,resize=no')" title=��������>����</a></b></font>] 
  <b>
 <%if session("Ba_jxqy_usercorp")="�ٸ�" then%> 
<a href=ask.asp target="_blank">[����]</a>
<%else%>
<a href=reply.asp target="_blank">[����]</a>
<%end if%>
</b>
</form>
</body> 