<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("aqjh_chatbgcolor")
chatimage=Application("aqjh_chatimage")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
'�Ƿ�����ݵ�
if cstr(chatinfo(11))=1 then
	jhpd=0
else
	jhpd=1
end if
if chatbgcolor="" then chatbgcolor="008888"%>
<HTML><HEAD><TITLE>������</TITLE>
<STYLE type=text/css>
BODY {FONT-SIZE: 9pt}
INPUT {FONT-SIZE: 9pt}
A {FONT-SIZE: 9pt; COLOR: #ffffff; TEXT-DECORATION: none}
A:hover {COLOR: red; TEXT-DECORATION: underline}
SELECT {BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-BOTTOM-WIDTH: 1px; COLOR: black; FONT-FAMILY: ����; BACKGROUND-COLOR: #efefef; BORDER-RIGHT-WIDTH: 1px}
TD {FONT-SIZE: 9pt; COLOR: #ffffff; FONT-FAMILY: ����}
</STYLE>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}</script>
<script language="JavaScript"> 
//if(window.top==window.self){var i=1;while (i<=50){window.alert("������ʲôѽ�����ң������ǲ��еģ�ȥ����ȥ�ɣ�����������50�Σ���");i=i+1;}top.location.href="../exit.asp"}
</script>
</HEAD>
<BODY bgcolor=<%=Application("aqjh_chatbgcolor")%> leftMargin=0 topMargin=0 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false onselectstart=event.returnValue=false>
<TABLE height=83 cellSpacing=0 cellPadding=0 width=100% align=right border=0 bordercolor=yellow style="border-collapse: collapse"><TR><TD>
<DIV align=center>
<form name=af method=POST action='say.asp'  target='d' onsubmit='javascript:if(document.af.ctz.checked==true){if(document.af.sytemp.value.indexOf("/")!=0){document.af.sytemp.value="/������$"+document.af.sytemp.value;}}  return(parent.checksays());'>
<input type=hidden name='fs' value='10'>
<input type=hidden name='lh' value='125'>
<input type=hidden name='sy' value=''>
<input type=hidden name='oldsays' value>
<input type=hidden name='oldact' value>
<input type=hidden name='oldtowho' value>
<input type=hidden name='npc' value='<%=Application("aqjh_npc")%>'>
<input type=hidden name='username' readonly size=5 maxlength=10>

<!����һ��>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;background-color:#F7F7F7'>
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
<option style=background-color:"#EEFFFF" value="/��ϯ$ ���">���
<option style=background-color:"#EEFFFF" value="/��ϯ$ ���">���
</select>  
<font color=red>��</font> <select name='addsays' onchange="document.af.sytemp.focus();" style='font-size:12px;background-color:#F7F7F7'> 
<option value="��" selected style="color:red">����</option> 
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
</select><font color=red>��</font>                                                                       
<input type="text" name="towho" readonly size="8" onClick="dj()" value="���" style="height: 18; text-align: center; border-style: groove; border-width:1;background-color:#F7F7F7">  
<font color=blue>˵</font> <input type=text name='sytemp' size=38 maxlength=250 style="height: 18; border-style: groove; border-width: 1;background-color:#F7F7F7"> <a href="#" onClick="gop()"><img src="ico/qian.gif" alt="��һ��" border="0" align="absmiddle"></a> <a href="#" onClick="gon()"><img src="ico/hou.gif" alt="��һ��" border="0" align="absmiddle"></a>        
<input type=submit name='subsay' value='��' style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #0088ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"> <input type=button value='��' onClick="javascript:parent.qp()" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: dd11ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"> <input type=button value='ħ��'  onClick="window.open('qq/SelQQ.asp','f3')"style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #8800ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'">
<input type=button value='���'  onClick="window.open('zhu/index.asp','f3')"style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #008888; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='ffffff'"><br>
<!���ܶ���>
<select name='addwordcolor' style='font-size:12px;background-color:#F7F7F7'> 
<option style="color:0088FF" value="0088FF" selected>��</option> 
<option style="color:000000" value="000000">��</option> 
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
<option style="color:red" value="red">��</option> 
</select><select name='saycolor'  style='font-size:12px;background-color:#F7F7F7'>  
<option style="color:008800" value="008800" selected>��</option> 
<option style="color:000000" value="000000">��</option> 
<option style="color:0088FF" value="0088FF">��</option> 
<option style="color:0000FF" value="0000FF">��</option> 
<option style="color:000088" value="000088">��</option> 
<option style="color:888800" value="888800">��</option> 
<option style="color:008888" value="008888">��</option> 
<option style="color:888888" value="888888">��</option> 
<option style="color:8888FF" value="8888FF">��</option> 
<option style="color:AA00CC" value="AA00CC">��</option> 
<option style="color:8800FF" value="8800FF">��</option> 
<option style="color:CCAA00" value="CCAA00">��</option> 
<option style="color:FF8800" value="FF8800">��</option> 
<option style="color:CC3366" value="CC3366">��</option> 
<option style="color:FF00FF" value="FF00FF">��</option> 
<option style="color:3366CC" value="3366CC">��</option>
<option style="color:red" value="red">��</option>
</select><%aqjh_name=aqjh_name%><select name='command' onchange="rc(this.value);document.af.command.options[0].selected=true;" style='font-size:12px;'><%                                                                    
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("aqjh_usermdb") 
rs.open "select ��Ա�ȼ�,����,���,ְҵ from �û� where ����='" & aqjh_name &"'",conn,2,2 
hydj=rs("��Ա�ȼ�") 
jhmp=rs("����") 
jhzy=rs("ְҵ") 
jhsf=rs("���") 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
%> 
<option value="" selected style="background-color:#F7F7F7">����һ��
<option style="color:#FF6633" value="">==����==
<option style="color:blue" value="/��ӭ����$">��ӭ����
<option style="color:#FF6633" value="/����$ �ҳ�ȥһ�£�����һ������� ">��ʱ�뿪
<option style="color:#FF6633" value="/����$ ">���߻���
<option style="color:blue" value="/���˷�$ 100000">�����˷�
<option style="color:red" value="/���˸���$">���˸���
<option style="color:#0808CF" value="/������$">������
<option style="color:#B000D0" value="/�����ͻ�$ ">�����ͻ�
<option style="color:#B000D0" value="/��������$ ">���뿴��
<%if chatinfo(5)=0 then%>
<option style="color:#FF6633" value="/���봮��$">���봮��
<option style="color:#FF6633" value="/���ͼ�԰$ ">���ͼ�԰
<option style="color:blue" value="/���������$ ">�������
<option style="color:#FF6633" value="/����$ ">�������� <%end if%>
<option style="color:#FF6633" value="/����$ ����">�����ֵ�
<option style="color:#FF6633" value="/����$ �鿴">�鿴�ֵ�
<option style="color:#FF6633" value="/����$ ɾ��">���۶���
<option style="color:#FF6633" value="/���$ Ҫ����ĺ�������">����ֵ�
<option style="color:#FF6633" value="/����$ ���˹ٸ�������!">���ı��� <% if Application("aqjh_baowu")>0 then%>
<option style="color:#FF6633" value="/����$ ">�������� <%end if%>
<option style="color:blue" value="/����$ ����">��������
<option style="color:blue" value="/����$ �鿴">�鿴����
<option style="color:blue" value="/����$ ɾ��">ȥ������
<option style="color:blue" value="/��Ϣ$">�鿴��Ϣ
<option style="color:red" value="">==����==
<option style="color:#B000D0" value="/��Ǯ$ 1000">���д�Ǯ
<option style="color:#B000D0" value="/ȡǮ$ 1000">����ȡǮ
<option style="color:#B000D0" value="/ת��$ 10000">����ת��
<option style="color:#B000D0" value="">==���==
<option style="color:#B000D0" value="/��Ҵ��$ 10">��Ҵ��
<option style="color:#B000D0" value="/���ȡ��$ 10">���ȡ��
<option style="color:red" value="">==����==
<option style="color:#0808CF" value="/��������$">��������
<option style="color:#0808CF" value="/���ŷ�Ǯ$">���ŷ�Ǯ
<option style="color:#0808CF" value="/ָ������$">ָ������
<option style="color:#0808CF" value="/����$">���Ż���
<option style="color:#0808CF" value="/��ͽ$ ">����ͽ��
<option style="color:#0808CF" value="/���$ ͽ����">���ʦ��
<option style="color:#0808CF" value="/��ʦ$ ">��ʦϰ��
<option style="color:blue" value="/����ͽ��$ 100">����ͽ��
<option style="color:blue" value="/�ͷ�ͽ��$ 100">�ͷ�ͽ��
<option style="color:#0808CF" value="/ѧ����ɽ$ ʦ����">ѧ���г�
<option style="color:#0808CF" value="/��ս����$ �����������˵�����">��ս����
<option style="color:red" value="">==����==
<option style="color:#FF6699" value="/���$ �氮���">ʾ�����
<option style="color:#FF6699" value="/��ϲ��$ 500">�ַ�ϲ��
<option style="color:#FF6699" value="/����$">���߷���
<option style="color:#FF6699" value="/���ܽӴ�$">���ܽӴ�
<option style="color:#FF6699" value="/����$">��ǿ��Ů
<option style="color:red" value="">==����==
<option style="color:#0808CF" value="/����$ ">ǧ�ﴫ��
<option style="color:#0808CF" value="/�Ķ�$ ">�Ķ��о�
<option style="color:#0808CF" value="/ŭ��$ ">��ʨŭ��
<option style="color:#0808CF" value="/����$ ">��������
<option style="color:red" value="">==͵Ϯ==
<option style="color:#333333" value="/����$ ����ק������ק">��������
<%if chatinfo(5)=0 then%>
<option style="color:#333333" value="/���Ǵ�$ ">��ȡ����
<option style="color:#333333" value="/͵Ǯ$ ">͵ȡǮ��
<option style="color:#333333" value="/͵���$ ">͵ȡ���
<option style="color:#333333" value="/Ǭ��һ��$ 10000">Ǭ��һ�� 
<option style="color:#333333" value="/�¶�$ ��ҩ��">͵͵�¶�
<option style="color:#333333" value="/Ͷ��$ ������">Ͷ������
<option style="color:#333333" value="/����$ ��ʽ��">���й���
<option style="color:#333333" value="/���ﲫ��$ ��������">���ﲫ��
<option style="color:#333333" value="/��Ǯ����$ 10000">��Ǯ���� <%end if%>
<%if jhsf="����" or jhsf="Ԫ��" then%>
<option style="color:#993333" value="">==����==
<option value="/�������$">�������
<option value="/���$ �����">������
<option style="color:blue" value="/���$ Ԫ��">���Ԫ��
<option value="/���$ ����">��ⳤ��
<option value="/���$ ����">��⻤��
<option value="/���$ ����">�������
<option value="/���$ ����">ȡ�����
<%end if
if (jhsf="����" or jhsf="����" or jhsf="Ԫ��") and aqjh_grade>=4 then%>
<option value="/���յ���$">���յ���
<option value="/��Ѩ$ ��д��ԭ�� ">��ѵ����
<option value="/����$ 100 ">��������
<option value="/������$ ���� ">��������
<option value="/����ip$">�� ��ip
<option style="color:red" value="/׷ɱ��$ ���� ">׷ ɱ ��
<option value="/���ɷ���$ 1000 ">���ɷ���
<option value="/�����̷�$ 1000 ">�����̷�
<option value="/���ŷ���$">���ŷ��� <%end if%>
</select><select name='addsign' onChange="rc(this.value);" style='font-size:12px;background-color:#F7F7F7'>  
        <option value="" selected>���ܶ���</option>
<option style="color:red" value="">=�¹���=
        <option style="color:blue" value="/�Է�$">�������
        <option style="color:blue" value="/����$">ף������
        <option style="color:blue" value="/��������$ ">��������
        <option style="color:blue" value="/ɱ�����$ 1">ɱ�����
        <option style="color:blue" value="/���齵��$ 100000">���齵��
        <option style="color:blue" value="/�չ�����$ 5000">�չ�����
        <option style="color:red" value="/������$">��������
        <option style="color:red" value="/����$">ת������  
        <option style="color:red" value="/��������$">�������� 
        <option style="color:red" value="/��������$">��������
        <option style="color:red" value="/��Ϊ����$">��Ϊ����
        <option style="color:red" value="/������$ ����">�� �� ��
<option style="color:red" value="">==����==
        <option style="color:blue" value="/���콱��$">���콱��</option>
        <option style="color:blue" value="/���˽���$ 1">���˽���</option>
        <option style="color:blue" value="/�߼�����$ 1">�߼�����</option>
<option style="color:red" value="">==����==
<option style="color:blue" value="/�����ʽ�$">�鿴����</option>
 <option style="color:blue" value="/���Ҳ��$ ة��/����/����/�Զ���">�������</option>
 <option style="color:blue" value="/������$  д����Ҫ˵�Ļ�">���ҹ���</option>
 <option style="color:blue" value="/��������$ 10000 ">��������</option>
<option style="color:blue" value="/��������$ ">��������</option>
<option style="color:blue" value="/��ս����$ ">��ս����</option>
<option style="color:red" value="">==��ս==
<option style="color:blue" value="/���Ҵ�ս$ ">���Ҵ�ս</option>
<option style="color:blue" value="/������ս$  д�϶Է�����������,��սʧ��,��������,ע��!">��������</option>
<option style="color:blue" value="/�������$  д�϶Է�����������">�������</option>
<option style="color:red" value="">==����==
<option style="color:blue" value="/���һ��$ ">���һ��
<option style="color:blue" value="/������$ 10000">������
<option style="color:blue" value="/��Ȼ����$ 10000">��Ȼ����
<option style="color:blue" value="/��ɱ$">�����Խ� 
<option style="color:red" value="">==����==
<option style="color:red" value="/��������$">�������� 
<option style="color:red" value="/��������$">��������
<option style="color:red" value="/ħ������$">ħ������
<option style="color:red" value="/��������$ ���������">�������� 
<option style="color:red" value="/����ħ��$ ħ�������">����ħ��
<option style="color:red" value="/��������$ ���ɵ�����">��������
<option style="color:red" value="">==����==
        <option style="color:#B000D0" value="/����$ ���|����|����">������Ʒ
        <option style="color:#B000D0" value="/����$ 100000">��������
        <option style="color:#B000D0" value="/�ÿ�$ ��Ƭ��">ʹ�ÿ�Ƭ
        <option style="color:#B000D0" value="/��Ǯ$ 1000">�����ֽ�
        <option style="color:#B000D0" value="/�Ͷ���$ 10">���Ͷ���
        <option style="color:#B000D0" value="/�ͽ��$ 1">���ͽ��
        <option style="color:blue" value="/��������$ 100">��������
        <option style="color:blue" value="/��������$ 100">��������
        <option style="color:blue" value="/�����书$ 100">�����书
        <option style="color:blue" value="/���ھ���$ 100">���ھ���
        <option style="color:blue" value="/��������$ 100">��������
        <option style="color:blue" value="/���ڵ���$ 100">���ڵ���
        <option style="color:blue" value="/���ͷ���$ 100">���ͷ���</option>
        <option style="color:blue" value="/�����Ṧ$ 100">�����Ṧ</option>
<option style="color:red" value="">==����==
<option style="color:#339933" value="/����$ ">��������
<option style="color:#339933" value="/��Ŀ$ ">��Ŀ����
<option style="color:#339933" value="/����$ ">���ϰ��
<option style="color:#339933" value="/����$ ">��������
<option style="color:#339933" value="/�������$ ">�������
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/�����Ṧ$ ">�����Ṧ
<option style="color:#996633" value="/��ϰ��$">��ϰ��
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/��ľ����$ ">��ľ����
<option style="color:#339933" value="/��ˮ����$ ">��ˮ����
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/��������$ ">��������
<option style="color:#339933" value="/��������$ ">��������
 <option style="color:red" value="">==ħ��==</option>
        <option style="color:#B000D0" value="/������ɽ$">������ɽ</option>
        <option style="color:#B000D0" value="/�����Ϸ�$">�����Ϸ�</option>
        <option style="color:#B000D0" value="/��ʯ�ɽ�$">��ʯ�ɽ�</option>
        <option style="color:#B000D0" value="/���Ĵ�$">���Ĵ�</option>
        <option style="color:#B000D0" value="/�вƽ���$">�вƽ���</option>
        <option style="color:#B000D0" value="/���ֻش�$">���ֻش�</option>
 <option style="color:red" value="">==���==</option> 
        <option style="color:green" value="/�����$ 100">�� �� �</option>       
        <option style="color:green" value="/�����$ 100">�� �� �</option>
        <option style="color:green" value="/�����$ 100">�� �� �</option>
        <option style="color:green" value="/ħ��ʯ$ 1">ħ �� ʯ</option>
        <option style="color:green" value="/��ȸ��$ 50">�� ȸ ��</option>
        <option style="color:green" value="/˧����$ ">˧ �� ��</option>
        <option style="color:green" value="/������$ ">�� �� ��</option>
        <option style="color:green" value="/���黷$ ">�� �� ��</option>
        <option style="color:green" value="/������$ ">�� �� ��</option>
 <option style="color:red" value="">==����==</option>
        <option style="color:blue" value="/ħ��ˮ��$ ">ħ��ˮ��</option>
        <option style="color:blue" value="/�����ӵ�$ ">�����ӵ�</option>
        <option style="color:blue" value="/ʹ��$ ������ ">������</option>
        <option style="color:blue" value="/ʹ��$ ������ ">�� �� �� </option>
        <option style="color:blue" value="/ʹ��$ ��ڤ�� ">�� ڤ �� </option>
        <option style="color:blue" value="/ʹ��$ ��ȡ�� ">�� ȡ �� </option>
        <option style="color:blue" value="/ʹ��$ ����ȭ ">�� �� ȭ </option>
        <option style="color:blue" value="/ʹ��$ ʥ���� ">ʥ �� �� </option>
        <option style="color:blue" value="/���鵶$ ">�� �� �� </option>
        <option style="color:blue" value="/������$ ">�� �� �� </option>
        <option style="color:blue" value="/ʹ��$ �������� ">�������� </option>
 <option style="color:red" value="">==����==</option>
        <option style="color:green" value="/����$ ">�� �� �� 
        <option style="color:green" value="/С���$ ">С �� ��
        <option style="color:#FF6699" value="/������$ ">�� �� ��</option>
        <option style="color:#FF6699" value="/������$ ">�� �� ��</option>
        <option style="color:#FF6699" value="/��ť��$ ">�� ť ��</option>
        <option style="color:#FF6699" value="/������ť$ ">������ť</option>     
        <option style="color:#FF6699" value="/���°�ť$ ">���°�ť</option>
         <option style="color:#FF6699" value="/����$ ">������</option>
         <option style="color:#FF6699" value="/�ߵ�$ ">�ߵ�����</option>
        <option style="color:#FF6699" value="/����ħ��$ ����|��ɫ|��С|����">����ħ��</option>
        <option style="color:#FF6699" value="/�ƶ�ħ��$ ����|��ɫ|����|��ʽ|�ٶ�|��ʱ">����ħ��</option>
        <option style="color:#FF6699" value="/��ťħ��$ ����|��ɫ|��������">��ťħ��</option>              
 <option style="color:red" value="">==NPC==</option>
        <option style="color:#FF6699" value="/�ٻ�$ ">�� �� ��</option>
        <option style="color:#FF6699" value="/��Ѫ$ ">�� �� ��</option>
</select><select name='com1' onchange="rc(this.value);document.af.com1.options[0].selected=true;" style='font-size:12px;'>
<option value="" selected>ְҵ����</option>
 <option style="color:red" value="">==����==</option> 
<option style="color:blue" value="/����$ ���Һ�ֻ�ܳ�ի�����ܳ���ͽ�ŮɫŶ">��������
<option style="color:blue" value="/����$">�޳ɻ���
<option style="color:blue" value="/�$">�ն�����
<option style="color:blue" value="/���崺��$">���崺��
 <option style="color:red" value="">==����==</option>
<option style="color:blue" value="/��������$ �����������ܱ���">��������
<option style="color:blue" value="/�뿪����$">��������
 <option style="color:red" value="">==����==</option>
<option style="color:#996633" value="/��ƭ��Ů$ ">��ƭ��Ů
<option style="color:#996633" value="/��ƭ����$ ">��ƭ����
<option style="color:#996633" value="/ץС͵$ ">ץ С ͵
<option style="color:#996633" value="/����$ С������">����С�� <%if jhzy="���" then%>
<option style="color:#996633" value="/�β�$ ">���ֻش� <%end if%> <%if jhzy="��ʦ" then%>
<option style="color:#996633" value="/����$ ">�̵��书 <%end if%> <%if jhzy="����ʦ" then%>
<option style="color:#996633" value="/��ǩ$ ">�����ʲ� <%end if%> <%if jhzy="��ҹ��" then%>
<option style="color:#996633" value="/��ҹ��$ ">��ҹ��ʦ <%end if%>
<option style="color:#996633" value="/ɨ���ж�$">ɨ���ж�
<option style="color:#996633" value="/ͬ���ھ�$">ͬ���ھ�
<option style="color:red" value="">=�ٻ�ʦ=</option>
<option style="color:#333333" value="/���޹���$">���޹���
<option style="color:#333333" value="/��������$">��������
 <option style="color:red" value="">=ð�ռ�=</option>        
        <option style="color:green" value="/Ѱ���ʻ�$ ">Ѱ���ʻ�</option>
        <option style="color:green" value="/Ѱ�ҽ�$ ">Ѱ�ҽ�</option>
        <option style="color:green" value="/Ѱ�ҽ��$ ">Ѱ�ҽ��</option>
        <option style="color:green" value="/Ѱ�ҿ�Ƭ$ ">Ѱ�ҿ�Ƭ</option>
        <option style="color:green" value="/Ѱ������$ ">Ѱ������</option>
        <option style="color:green" value="/Ѱˮ����$ ">Ѱ��ˮ��</option>
        <option style="color:green" value="/Ѱ�ҷ���$ ">Ѱ�ҷ���</option>
        <option style="color:green" value="/Ѱ��ħ��$ ">Ѱ��ħ��</option>
        <option style="color:green" value="/������$ ">Ѱ�ұ���</option>
 <option style="color:red" value="">=ħ��ʦ=</option>
        <option style="color:blue" value="/��������$ ">��������</option>
        <option style="color:blue" value="/��ʩ��$ ">��ʩ����</option>
        <option style="color:blue" value="/�Ȼ��˼�$ ">�Ȼ��˼�</option>
        <option style="color:blue" value="/ħ������$ ">ħ������</option>
        <option style="color:blue" value="/�Ի��$ ">�Ի��</option>
        <option style="color:blue" value="/��������$ ">��������</option>
        <option style="color:blue" value="/���Ʊ�ʯ$ ">���Ʊ�ʯ</option>
 <option style="color:red" value="">=����ʦ=</option>
        <option style="color:blue" value="/�����ϳ�$ ">�����ϳ�</option>
        <option style="color:blue" value="/��Ƭ�ϳ�$ ">��Ƭ�ϳ�</option>
        <option style="color:blue" value="/�ٱ���ͨ$ ">�ٱ���ͨ</option>
        <option style="color:blue" value="/������$ 100">�� �� ��</option>
        <option style="color:blue" value="/��ħ��ʯ$">��ħ��ʯ</option>
        <option style="color:blue" value="/�ⶾ��$ 100">�� �� ��</option>
        <option style="color:blue" value="/ƽ����$ 100">ƽ �� ��</option>
        <option style="color:blue" value="/͵����$ 100">͵ �� ��</option>
        <option style="color:blue" value="/ҡǮ��$ 100">ҡ Ǯ ��</option>
        <option style="color:blue" value="/ħ����ʯ$ ">ħ����ʯ</option>
        <option style="color:blue" value="/���յ���$ ">���յ���</option>
 <option style="color:red" value="">==�ٸ�==</option>
        <option style="color:blue" value="/û�շ���$ ">û�շ���</option>
        <option style="color:blue" value="/û��ħ��$ ">û��ħ��</option>
 <option style="color:red" value="">==����==</option>
        <option style="color:#B000D0" value="/�������$ 10">�������
        <option style="color:#B000D0" value="/ȡ������$ 10">ȡ������
        <option style="color:#B000D0" value="/��ľ����$ 10">��ľ����
        <option style="color:#B000D0" value="/ȡľ����$ 10">ȡľ����
        <option style="color:#B000D0" value="/��ˮ����$ 10">��ˮ����
        <option style="color:#B000D0" value="/ȡˮ����$ 10">ȡˮ����
        <option style="color:#B000D0" value="/�������$ 10">�������
        <option style="color:#B000D0" value="/ȡ������$ 10">ȡ������
        <option style="color:#B000D0" value="/��������$ 10">��������
        <option style="color:#B000D0" value="/ȡ������$ 10">ȡ������
        <option style="color:#FF6699" value="/�淨��$ 100">��ŷ���</option>
        <option style="color:#FF6699" value="/ȡ����$ 100">ȡ�ط���</option>
        <option style="color:#FF6699" value="/�Ṧ�ݴ�$ 100">�Ṧ�ݴ�</option>
        <option style="color:#FF6699" value="/����Ṧ$ 100">����Ṧ</option>
 <option style="color:red" value="">==����==</option>
        <option style="color:FF8800" value="/��������$">��������</option>
        <option style="color:FF8800" value="/���ٽ��$">���ٽ��</option>
 <option style="color:red" value="">==�ҷ�==</option>
        <option style="color:#B000D0" value="/š��$">š �� ��</option>
        <option style="color:#B000D0" value="/��ë����$">��ë����</option>
        <option style="color:#B000D0" value="/����$">�� �� ��</option>
</select><select name='dubo' onchange="rc(this.value);document.af.dubo.options[0].selected=true;" style='font-size:12px;background-color:#F7F7F7'>                                                       
<option value="" selected>=�Ĳ�=
        <option value="/���˿�$ 1000|����[�򣺽�ң�����������]" style=color:red>���˿� 
        <option value="/����$ ��֤|��[��ʤ]" style=color:red>-��֤-
        <option value="/���齫$ 1000|����[�򣺽�ң�����������]" style=color:red>���齫 
        <option value="/����$ ��֤|��[��ʤ]" style=color:red>-��֤-
<option style="color:blue" value="/��������">��&nbsp;&nbsp;��
<option style="color:blue" value="/˫�˶Ĳ�$ 1">��ʯͷ
<option style="color:blue" value="/˫�˶Ĳ�$ 2">�ļ���
<option style="color:blue" value="/˫�˶Ĳ�$ 3">�Ĳ���
<option style="color:red" value="/�ķ���$100">�ķ���
<option style="color:red" value="/���Ṧ$100">���Ṧ
<option style="color:red" value="/�Ľ�����$100">�Ľ���
<option style="color:red" value="/��ľ����$100">��ľ��
<option style="color:red" value="/��ˮ����$100">��ˮ��
<option style="color:red" value="/�Ļ�����$100">�Ļ���
<option style="color:red" value="/��������$100">������
<option style="color:red" value="/�Ķ�$100">�Ķ���
<option style="color:red" value="/�Զ�$2">�Ľ��
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
<option style="color:blue" value="/�Ĳ�$ ����ׯ">����ׯ
<option style="color:blue" value="/��ע$ ��&��&100">��Ѻ��
<option style="color:blue" value="/��ע$ ��&С&100">��ѺС
<option style="color:blue" value="/��ע$ ��&˫&100">��Ѻ˫
<option style="color:blue" value="/��ע$ ��&��&100">��Ѻ��
<option style="color:green" value="/�Ĳ�$ ����ׯ">����ׯ
<option style="color:green" value="/��ע$ ��&��&100">��Ѻ��
<option style="color:green" value="/��ע$ ��&С&100">��ѺС
<option style="color:green" value="/��ע$ ��&˫&100">��Ѻ˫
<option style="color:green" value="/��ע$ ��&��&100">��Ѻ��
</select><select name='tu' onChange="rc(this.value);document.af.tu.options[0].selected=true;" style='font-size:12px;background-color:#F7F7F7'>
<option value="" selected>ͼƬ��</option>
<option style="color:red" value="">A��ͼ</option> 
<option value="[IMG]a1.gif[/IMG]">�� ��</option>
<option value="[IMG]a2.gif[/IMG]">�� ��</option>
<option value="[IMG]a3.gif[/IMG]">�� ��</option>
<option value="[IMG]a4.gif[/IMG]">�� ��</option>
<option value="[IMG]a5.gif[/IMG]">�� ��</option>
<option value="[IMG]a6.gif[/IMG]">�� Ƥ</option>
<option value="[IMG]a7.gif[/IMG]">�� ��</option>
<option value="[IMG]a27.gif[/IMG]">�� ɫ</option>
<option value="[IMG]a9.gif[/IMG]">�� ��</option>
<option value="[IMG]a10.gif[/IMG]">�� ��</option>
<option value="[IMG]a11.gif[/IMG]">�� ��</option>
<option value="[IMG]a12.gif[/IMG]">�� ��</option>
<option value="[IMG]a13.gif[/IMG]">�� ��</option>
<option value="[IMG]a14.gif[/IMG]">ͷ ��</option>
<option value="[IMG]a15.gif[/IMG]">ʹ ��</option>
<option value="[IMG]a16.gif[/IMG]">�� ��</option>
<option value="[IMG]a17.gif[/IMG]">�� ��</option>
<option value="[IMG]a18.gif[/IMG]">װ ��</option>
<option value="[IMG]a19.gif[/IMG]">�� ��</option>
<option value="[IMG]a20.gif[/IMG]">װ ɵ</option>
<option value="[IMG]a21.gif[/IMG]">�� ��</option>
<option value="[IMG]a22.gif[/IMG]">�� ��</option>
<option value="[IMG]a23.gif[/IMG]">͵ Ц</option>
<option value="[IMG]a24.gif[/IMG]">�� ��</option>
<option value="[IMG]a25.gif[/IMG]">Ц ��</option>
<option style="color:red" value="">B��ͼ</option>
<option value="[IMG]b1.gif[/IMG]">ɵ Ц</option>
<option value="[IMG]b2.gif[/IMG]">װ ��</option>
<option value="[IMG]b3.gif[/IMG]">�� ��</option>
<option value="[IMG]b4.gif[/IMG]">�� ��</option>
<option value="[IMG]b5.gif[/IMG]">�� ˮ</option>
<option value="[IMG]b6.gif[/IMG]">�� ��</option>
<option value="[IMG]b7.gif[/IMG]">�� ��</option>
<option style="color:red" value="">C��ͼ</option>
<option value="[IMG]c1.gif[/IMG]">�� ò</option>
<option value="[IMG]c2.gif[/IMG]">�� ��</option>
<option value="[IMG]c3.gif[/IMG]">�� ��</option>
<option value="[IMG]c4.gif[/IMG]">�� ��</option>
<option value="[IMG]c5.gif[/IMG]">�� ��</option>
<option value="[IMG]c6.gif[/IMG]">�� ��</option>
<option value="[IMG]c7.gif[/IMG]">�� ��</option>
<option value="[IMG]c8.gif[/IMG]">�� ��</option>
<option value="[IMG]c9.gif[/IMG]">�� ��</option>
<option value="[IMG]c10.gif[/IMG]">˯ ��</option>
<option value="[IMG]c11.gif[/IMG]">�� ��</option>
<option value="[IMG]c12.gif[/IMG]">�� ��</option>
<option value="[IMG]c13.gif[/IMG]">������</option>
<option value="[IMG]c14.gif[/IMG]">�� ��</option>
<option value="[IMG]c15.gif[/IMG]">ɵ Ц</option>
<option value="[IMG]c16.gif[/IMG]">���뿪</option>
<option value="[IMG]c17.gif[/IMG]">������</option>
<option value="[IMG]c18.gif[/IMG]">��è��</option>
<option value="[IMG]c19.gif[/IMG]">ȥ����</option>
<option value="[IMG]c20.gif[/IMG]">ͬ ��</option>
<option value="[IMG]c21.gif[/IMG]">ǿ ��</option>
<option value="[IMG]c22.gif[/IMG]">������</option>
<option style="color:red" value="">D��ͼ</option>
<option value="[IMG]d1.gif[/IMG]">�� ��</option>
<option value="[IMG]d2.gif[/IMG]">��һ��</option>
<option value="[IMG]d3.gif[/IMG]">�� ��</option>
<option value="[IMG]d4.gif[/IMG]">�� ��</option>
<option value="[IMG]d5.gif[/IMG]">�� ��</option>
<option value="[IMG]d6.gif[/IMG]">�� ��</option>
<option value="[IMG]d7.gif[/IMG]">�� ��</option>
<option value="[IMG]d8.gif[/IMG]">�� ��</option>
<option value="[IMG]d9.gif[/IMG]">�� ��</option>
<option value="[IMG]d10.gif[/IMG]">������</option>
<option value="[IMG]d11.gif[/IMG]">�� ��</option>
<option value="[IMG]d12.gif[/IMG]">�� ��</option>
<option style="color:red" value="">E��ͼ</option>
<option value="[IMG]e1.gif[/IMG]">�� ��</option>
<option value="[IMG]e2.gif[/IMG]">�� ��</option>
<option value="[IMG]e8.gif[/IMG]">ײ ͷ</option>
<option value="[IMG]e22.gif[/IMG]">�� ��</option>
<option value="[IMG]e23.gif[/IMG]">�� ��</option>
<option value="[IMG]e24.gif[/IMG]">�� ��</option>
<option value="[IMG]e25.gif[/IMG]">�� ��</option>
<option value="[IMG]e26.gif[/IMG]">ȭ ��</option>
<option value="[IMG]e27.gif[/IMG]">�� ��</option>
<option value="[IMG]e28.gif[/IMG]">������</option>
<option value="[IMG]e29.gif[/IMG]">�� Ĭ</option>
<option value="[IMG]e30.gif[/IMG]">��Ů��</option>
<option value="[IMG]e31.gif[/IMG]">�� ��</option>
<option value="[IMG]e32.gif[/IMG]">�� ��</option>
<option value="[IMG]e33.gif[/IMG]">�� ��</option>
<option value="[IMG]e34.gif[/IMG]">�� ��</option>
<option value="[IMG]e35.gif[/IMG]">�� ŭ</option>
<option value="[IMG]e36.gif[/IMG]">�� ��</option>
<option value="[IMG]e37.gif[/IMG]">ţ ��</option>
<option value="[IMG]e38.gif[/IMG]">�� ��</option>
<option value="[IMG]e39.gif[/IMG]">�� å</option>
<option value="[IMG]e40.gif[/IMG]">������</option>
<option value="[IMG]e41.gif[/IMG]">������</option>
<option value="[IMG]e42.gif[/IMG]">������</option>
<option value="[IMG]e43.gif[/IMG]">�� ��</option>
<option value="[IMG]e44.gif[/IMG]">�� ��</option>
<option value="[IMG]e45.gif[/IMG]">������</option>
<option value="[IMG]e46.gif[/IMG]">�� ��</option>
<option value="[IMG]e3.gif[/IMG]">��Ӱ��</option>
<option value="[IMG]e4.gif[/IMG]">̫��ȭ</option>
<option value="[IMG]e5.gif[/IMG]">�� ��</option>
<option value="[IMG]e6.gif[/IMG]">�� ��</option>
<option value="[IMG]e7.gif[/IMG]">������</option>
<option value="[IMG]e9.gif[/IMG]">��ƨ��</option>
<option value="[IMG]e10.gif[/IMG]">������</option>
<option value="[IMG]e11.gif[/IMG]">��å��</option>
<option value="[IMG]e12.gif[/IMG]">�� ��</option>
<option value="[IMG]e13.gif[/IMG]">�� ��</option>
<option value="[IMG]e14.gif[/IMG]">�� ��</option>
<option value="[IMG]e15.gif[/IMG]">�� ��</option>
<option value="[IMG]e16.jpg[/IMG]">�� ��</option>
<option value="[IMG]e17.gif[/IMG]">�� ԥ</option>
<option value="[IMG]e18.jpg[/IMG]">�� ��</option>
<option value="[IMG]e19.jpg[/IMG]">�� ��</option>
<option value="[IMG]e20.jpg[/IMG]">�� ��</option>
<option value="[IMG]e21.jpg[/IMG]">ɵ Ц</option>
<option style="color:red" value="">F��ͼ</option>
<option value="[IMG]f1.jpg[/IMG]">û����</option> 
<option value="[IMG]f2.gif[/IMG]">��ɱ��</option> 
<option value="[IMG]f3.gif[/IMG]">�ֱ�Ť</option> 
<option value="[IMG]f4.gif[/IMG]">�� ��</option> 
<option value="[IMG]f5.gif[/IMG]">˯����</option> 
<option value="[IMG]f6.gif[/IMG]">�� ��</option> 
<option value="[IMG]f7.gif[/IMG]">�� ��</option> 
<option value="[IMG]f8.gif[/IMG]">Ц����</option> 
<option value="[IMG]f9.gif[/IMG]">�ұ���</option> 
<option value="[IMG]f10.gif[/IMG]">�����</option> 
<option value="[IMG]f11.gif[/IMG]">�� ŭ</option> 
<option value="[IMG]f12.gif[/IMG]">���Ұ�</option> 
<option value="[IMG]f13.gif[/IMG]">ð�亹</option> 
<option value="[IMG]f14.gif[/IMG]">�� ��</option> 
<option value="[IMG]f15.gif[/IMG]">�����</option> 
<option value="[IMG]f16.gif[/IMG]">������</option>
<option value="[IMG]f17.gif[/IMG]">�� ��</option>
<option value="[IMG]f18.gif[/IMG]">����ͷ</option>
<option value="[IMG]f19.JPG[/IMG]">Ц ��</option>
<option value="[IMG]f20.gif[/IMG]">�� ��</option>
<option value="[IMG]f21.gif[/IMG]">ɫ ɫ</option>
<option value="[IMG]f22.gif[/IMG]">�� ��</option>
<option value="[IMG]f23.gif[/IMG]">ʤ ��</option>
<option value="[IMG]f24.gif[/IMG]">�� ͷ</option>
<option value="[IMG]f25.gif[/IMG]">û ��</option>
<option value="[IMG]f26.gif[/IMG]">�� ��</option>
<option value="[IMG]f27.gif[/IMG]">ͯ ��</option>
<option style="color:red" value="">G��ͼ</option>
<option value="[IMG]g1.gif[/IMG]">������</option>
<option value="[IMG]g2.gif[/IMG]">�Ҵ���</option>
<option value="[IMG]g3.gif[/IMG]">Ⱥ Ź</option>
<option value="[IMG]g4.gif[/IMG]">��ʤ��</option>
<option value="[IMG]g5.gif[/IMG]">�� ��</option>
<option value="[IMG]g6.gif[/IMG]">������</option>
<option value="[IMG]g7.gif[/IMG]">������</option>
<option value="[IMG]g8.gif[/IMG]">������</option>
<option value="[IMG]g9.gif[/IMG]">��ί��</option>
<option value="[IMG]g10.gif[/IMG]">������</option>
<option value="[IMG]g11.gif[/IMG]">Ц���</option>
<option value="[IMG]g12.gif[/IMG]">����Զ</option>
<option value="[IMG]g13.gif[/IMG]">������</option>
<option value="[IMG]g14.gif[/IMG]">�� ��</option>
<option value="[IMG]g15.jpg[/IMG]">�󻵵�</option>
<option value="[IMG]g16.gif[/IMG]">��һ��</option>
<option value="[IMG]g17.gif[/IMG]">�򿪴�</option>
<option value="[IMG]g18.gif[/IMG]">һ���</option>
<option value="[IMG]g19.gif[/IMG]">�� ��</option>
<option value="[IMG]g20.gif[/IMG]">������</option>
<option value="[IMG]g21.gif[/IMG]">�Ұ���</option>
<option value="[IMG]g22.gif[/IMG]">�� ��</option>
<option value="[IMG]g23.gif[/IMG]">������</option>
<option value="[IMG]g24.gif[/IMG]">������</option>
<option value="[IMG]g25.gif[/IMG]">�Ҵ���</option>
<option value="[IMG]g26.gif[/IMG]">������</option>
<option value="[IMG]g27.gif[/IMG]">��׷��</option>
<option value="[IMG]g28.gif[/IMG]">�޸���</option>
<option value="[IMG]g29.gif[/IMG]">������</option>
<option value="[IMG]g30.gif[/IMG]">��Ů��</option>
<option value="[IMG]g31.gif[/IMG]">˧����</option>
<option value="[IMG]g32.gif[/IMG]">�ҷ���</option>
<option value="[IMG]g33.gif[/IMG]">������</option>
<option value="[IMG]g34.gif[/IMG]">������</option>
<option value="[IMG]g35.gif[/IMG]">������</option>
<option style="color:red" value="">H��ͼ</option>
<option value="[IMG]h1.gif[/IMG]">ɫ����</option>
<option value="[IMG]h2.gif[/IMG]">� ��</option>
<option value="[IMG]h3.gif[/IMG]">̾ ��</option>
<option value="[IMG]h4.gif[/IMG]">������</option>
<option value="[IMG]h5.gif[/IMG]">�� ��</option>
<option value="[IMG]h6.gif[/IMG]">���ǰ�</option>
<option value="[IMG]h7.gif[/IMG]">�� ù</option>
<option value="[IMG]h8.gif[/IMG]">�� װ</option>
<option value="[IMG]h9.gif[/IMG]">Ǯ ��</option>
<option value="[IMG]h10.gif[/IMG]">�� ˮ</option>
<option value="[IMG]h11.gif[/IMG]">�� Ƥ</option>
<option value="[IMG]h12.gif[/IMG]">�� ��</option>
<option value="[IMG]h13.gif[/IMG]">�� ϲ</option>
<option value="[IMG]h14.gif[/IMG]">�� ŭ</option>
<option value="[IMG]h15.gif[/IMG]">�� ��</option>
<option value="[IMG]h16.gif[/IMG]">͵ ��</option>
<option value="[IMG]h17.gif[/IMG]">͵ ��</option>
<option value="[IMG]h18.gif[/IMG]">͵ Ц</option>
<option value="[IMG]h19.gif[/IMG]">�� ��</option>
<option value="[IMG]h20.gif[/IMG]">�� ��</option>
<option value="[IMG]h21.gif[/IMG]">�� ��</option>
<option value="[IMG]h22.gif[/IMG]">�� ��</option>
<option value="[IMG]h23.gif[/IMG]">ɵ Ц</option>
<option value="[IMG]h24.gif[/IMG]">�� ��</option>
<option value="[IMG]h25.gif[/IMG]">�� ��</option>
<option value="[IMG]h26.gif[/IMG]">�� ��</option>
<option value="[IMG]h27.gif[/IMG]">�����</option>
<option value="[IMG]h28.gif[/IMG]">�� ù</option>
<option value="[IMG]h29.gif[/IMG]">�� ��</option>
<option value="[IMG]h30.gif[/IMG]">ͷ ��</option>
<option value="[IMG]h31.gif[/IMG]">�� ��</option>
<option value="[IMG]h32.gif[/IMG]">�� ��</option>
<option value="[IMG]h33.gif[/IMG]">�� ��</option>
<option value="[IMG]h34.gif[/IMG]">�� ��</option>
<option value="[IMG]h35.gif[/IMG]">�� ��</option>
<option value="[IMG]h36.gif[/IMG]">�� ��</option>
<option value="[IMG]h37.gif[/IMG]">�� ��</option>
<option value="[IMG]h38.gif[/IMG]">Ͷ ��</option>
<option value="[IMG]h39.gif[/IMG]">������</option>
<option value="[IMG]h40.gif[/IMG]">�� ��</option>
<option value="[IMG]h41.gif[/IMG]">С����</option>
<option value="[IMG]h42.gif[/IMG]">ҧ ��</option>
<option value="[IMG]h43.gif[/IMG]">�� ��</option>
<option value="[IMG]h44.gif[/IMG]">�� ק</option>
<option value="[IMG]h45.gif[/IMG]">�� ��</option>
</select><select name="bgc" onchange="parent.f1.document.bgColor=parent.f0.document.bgColor=this.options[selectedIndex].value;this.value='#eeeeee';"  style="font-size:12px">
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
<select name="bgc1" onchange="parent.gg.document.bgColor=parent.f3.document.bgColor=parent.f2.document.bgColor=parent.CW_MENU.document.bgColor=this.options[selectedIndex].value;this.value='#eeeeee';"  style="font-size:12px">
<option value="#eeeeee" selected>�Ҳ���
<option value="<%=Application("aqjh_chatbgcolor")%>" STYLE="background-color:#<%=Application("aqjh_chatbgcolor")%>">Ĭ��
<option value="FBE7DB" STYLE="background-color:#FBE7DB">����
<option value="ffeaea" STYLE="background-color:#ffeaea">�ۺ�
<option value="FCF8E2" STYLE="background-color:#FCF8E2">���
<option value="eaffea" STYLE="background-color:#eaffea">ǳ��
<option value="effaff" STYLE="background-color:#effaff">����
<option value="f2f2ff" STYLE="background-color:#f2f2ff">ǳ��
<option value="eaeaff" STYLE="background-color:#eaeaff">����
<option value="000000" STYLE="background-color:#eaeaff">��ɫ
<option value="ffffff" STYLE="background-color:#ffffff">��ɫ
</select>
<input type=button name="tbclutch" value="ȫ��" onClick="javascript:parent.tbclutch();" title="����/����/��ֱ�л�" onMouseOver="window.status='����/����/��ֱ�л���';return true" onMouseOut="window.status='';return true" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #008888; BORDER-RIGHT-WIDTH: 1px">  <input type=button value='��' onClick="window.open('fafang/cins.asp','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #8800ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"> <input type=button value='NPC' onClick="window.open('npc_list.asp','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #cc0000; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'">  <input type=button value='����' onClick="window.open('../hcjs/pig/zhu.asp','aqjh_win','scrollbars=0,resizable=0,width=670,height=400')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #cc0000; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'">  <br>
<!��������>
<select name='command1' onchange="rc(this.value);document.af.command1.options[0].selected=true;" style='font-size:12px;'>
<option value="" selected>��������</option>
<%if aqjh_grade>=6 then%>
<option style="color:red" value="/����$ ">�ٸ�����
<option style="color:blue" value="/��Ѩ$ ��д��ԭ�� ">��Ѩ����
<option style="color:blue" value="/��Ѩ$ ��д��ԭ�� ">��Ѩ����
<option style="color:blue" value="/��Ѩ$ �û���">��Ѩ����
<option style="color:blue" value="/����$ ��д��ԭ�� ">�������
<option style="color:blue" value="/����$ ��д��ԭ�� ">��������
<option style="color:blue" value="/����$ ��д��ԭ�� ">���β���
<option style="color:blue" value="/��ip$ ">��IP��ַ
<option style="color:blue" value="/����ͬip$ ">����ͬIP
<option style="color:red" value="/���$ ͨ��������">ȡ��ͨ��
<option style="color:blue" value="/����$ ��д��ԭ�� ">��������
<option style="color:blue" value="/ը��$ �������㣿">����ը�� <%
end if
if aqjh_grade>=8 then%>
<option style="color:red" value="/״̬$ ">�鿴״̬
<option style="color:blue" value="/���$ ��д��ԭ�� ">�������
<option style="color:blue" value="/�ͷ�$ �û���">������
<option style="color:red" value="/����$ ����">�������� <%
end if
if aqjh_grade>=9 then%>
<option style="color:red" value="/���ջ���$ ף�����ÿ���...">���ջ���
<option style="color:red" value="/�Ŵ�$ ">�Ŵ�����
<option style="color:red" value="/���ip$ ĳһ�˵��ң��������ip��ͬ�û���">��IP��ַ
<option style="color:red" value="/����ip$ ԭ�����������ж�ε��ң�">��IP��ַ
<option style="color:red" value="/ͨ��$ ͨ��������">ͨ���˷�
<option style="color:red" value="/�ٸ��Ź�$">�ٸ�����
<option style="color:red" value="/ն��$ ԭ�� ">ն���˷� <%
end if
if aqjh_grade>=10 then%>
<option style="color:red" value="/�ͽ�$ 1">���ͽ�</option>
<option style="color:red" value="/���ܽ���$ 1">���ܽ���</option>
<option style="color:red" value="/���⽱��$">���⽱��</option>
<option style="color:blue" value="/�ٸ�����$ 500000">�ٸ�����
<option style="color:blue" value="/��⻨��$ 20">��⻨��
<option style="color:blue" value="/���״Ԫ$ 20">���״Ԫ
<option style="color:blue" value="/�������$">�������
<option style="color:red" value="/�鿴��Ʒ$">�鿴��Ʒ
<option style="color:red" value="<script>parent.f3.location.href='savevalue.asp';</script>�������ȫ�����!">ȫ�����
<option style="color:red" value="/վ������$">վ������
<option style="color:red" value="/ԭ�ӵ�$ ̫������">����ը�� <%end if%>
</select><input type=text name='clock' style="FONT-SIZE: 9pt; COLOR: #000000; BACKGROUND-COLOR: #c0c0c0; TEXT-ALIGN: right" value="" size=3 readonly> <input type=button value='��' onClick="window.open('qiaozui.asp','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #8800ff; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"> <input type=button value='��' onClick="window.open('f2/xp.asp?name=<%=aqjh_name%>','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: ff5522; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"><%if zhuan>=10 or aqjh_grade>=9 then%> <input type=button value='צ' onClick="window.open('sjfunc/ckys.asp','f3')" style="BORDER-TOP-WIDTH: 1px; BORDER-LEFT-WIDTH: 1px; BORDER-BOTTOM-WIDTH: 1px; COLOR: #ffffff; BACKGROUND-COLOR: #cc0000; BORDER-RIGHT-WIDTH: 1px" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'"><%end if%>
 <input type="checkbox" name="addvalues" checked=true onClick='if (<%=jhpd%>>=1){alert("�˷����ֹ�ݵ㣡");document.af.addvalues.checked=false;document.af.sytemp.focus();return false;}'><a href="#" onClick='if (<%=jhpd%>>=1){alert("�˷����ֹ�ݵ㣡");document.af.addvalues.checked=false;document.af.sytemp.focus();return false;};document.af.addvalues.checked=!(document.af.addvalues.checked);document.af.sytemp.focus()' style="color:red">��</a> <input type=checkbox name='Zshow' value='1' accesskey='z' <%if aqjh_jhdj<90 then%>disabled<%end if%> onClick="document.af.Zshow.focus();"><a href='#' onClick='<%if aqjh_jhdj<90 then%>{alert("�����㹦��90�����²���ʹ�ã�")}return false;<%else%>document.af.Zshow.checked=!(document.af.Zshow.checked);<%end if%>document.af.Zshow.focus();' title='�����㹦�ܣ�90�����²�����' style="color:red">��</a> <input type='checkbox' name='py' accesskey='v'><a href=# onClick='document.af.py.checked=!(document.af.py.checked);' title="��ĺ���(��ż)���ߡ��뿪ʱ�Ƿ��Զ�֪ͨ����" style="color:red">��</a> <input type='checkbox' name='mdsx' accesskey='b' checked><a href=# onClick='document.af.mdsx.checked=!(document.af.mdsx.checked);parent.m.location.reload();' title="�����˽����˳�ʱ�Ƿ��Զ�ˢ��������" style="color:red">��</a>  <input onClick='document.af.dwtx.focus();' type=checkbox name=dwtx checked title='�Ƿ�ʹ��ϵͳ��Ч����(�磺�𶯡����ֵ�)��' value="ON"><a href='#' onClick='document.af.dwtx.checked=!(document.af.dwtx.checked);document.af.dwtx.focus();' title='�Ƿ�ʹ��ϵͳ��Ч����(�磺�𶯡����ֵ�)��' style="color:red">��</a>  <input type=checkbox name='towhoway' value='1' accesskey='s' <%if aqjh_jhdj<60 or (aqjh_grade>5 and aqjh_grade<9) then%>disabled<%end if%> onClick="document.af.sytemp.focus();"><a href='#' onClick='<%if aqjh_jhdj<60 or (aqjh_grade>5 and aqjh_grade<9) then%>{alert("�ܱ�Ǹ���˹���60�����º͹ٸ������ã�")}return false;<%else%>document.af.towhoway.checked=!(document.af.towhoway.checked);<%end if%>document.af.sytemp.focus();' title='���Ļ�������˵��60�����º͹ٸ�������' style="color:red">˽</a> 
            <input type='checkbox' name=as accesskey='a' checked onClick='parent.f1.scrollit();document.af.sytemp.focus();' value="ON"><a href='#' onClick='document.af.as.checked=!(document.af.as.checked);document.af.as.focus();parent.f1.scrollit();document.af.sytemp.focus();' title='�Ƿ������ʾ; ' style="color:red">��</a>  
            <input type='checkbox' name='ctz' accesskey='v' value="OFF"><a href=# onMouseOver="window.status='�ڷ���ʱ���Ƿ���ʾ�����������ͷ��';return true" onMouseOut="window.status='���ֽ����״����������žͿ���ʹ�ã�';return true"  onClick='document.af.ctz.checked=!(document.af.ctz.checked);' title="�ڷ���ʱ���Ƿ���ʾ�����������ͷ��" style="color:red">��</a>  
           <%if aqjh_grade>=6 then%>  
            <a href="#" onClick="window.open('../twt/dt/ask.asp','ask','scrollbars=no,resizable=no,width=400,height=300')" style="color:red">��</a>    <a class=blue href="#" onClick="window.open('../twt/dalie/ZJ.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="ֻҪ�ٸ��������ˣ�����Դ������Ǯ��" style="color:red">��</a>                                                                     
            <a class=blue href="#" onClick="window.open('../twt/jiaofei/ZJ.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="ֻҪ�ٸ��ŷ��ˣ�����Դ�˵�Ǯ��" style="color:red">��</a>                                                                     
     <%else%>                                                                    
            <a href="hy.asp" target="_blank" style="color:red">��</a>                                                                    
            <a class=blue href="#" onClick="window.open('../twt/dt/reply.asp','daliwu','scrollbars=no,resizable=no,width=490,height=600')" title="�����ֻ��ɴ����������׬ร�" style="color:red">��</a>                                                                     
            <a class=blue href="#" onClick="window.open('../twt/dalie/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="ֻҪ�ٸ��������ˣ�����Դ������Ǯ��" style="color:red">��</a>                                                                    
            <a class=blue href="#" onClick="window.open('../twt/jiaofei/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="ֻҪ�ٸ��ŷ��ˣ�����Դ�˵�Ǯ��" style="color:red">��</a><%end if%> <%if chatinfo(7)=0 then%>                                                                    
 <a href=lg.asp title="��������ʱ�䣬��������ǲ��ܴ�ܵģ�" target="d" style="color:blue">��</a>
<%end if%> 
            <%if instr("," & Application("aqjh_slbox") & ",","," & aqjh_name & ",")<>0 then%>
            <input type='checkbox' name='slbox' accesskey='s' onClick="document.af.sytemp.focus();if (parent.slbox==1){parent.slbox=0}else{parent.slbox=1}" value="ON"><a href='#' onMouseOver="window.status='ѡ�б���󣬿��Բ鿴���������е�˽�ģ���';return true" onMouseOut="window.status='';return true" onClick="document.af.slbox.checked=!(document.af.slbox.checked);if (parent.slbox==1){parent.slbox=0}else{parent.slbox=1};document.af.sytemp.focus();" title="��һ�������˵һЩʲô���Ļ�" style="color:blue">��</a>                                                                     
            <%end if%>
<script language="JavaScript">function startnosay(){var nosay=<%=Application("aqjh_maxtimeout")*60%>;document.af.clock.value=nosay;}</script>                                                                    
            <script src="readonly/f2.js"></script>                                                                    
            <script>parent.fc();parent.m.location.href="f3.asp";</script>                               
</DIV></BODY></HTML>