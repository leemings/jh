var Quote = 0;
var Bold  = 0;
var Italic = 0;
var Underline = 0;
var Code = 0;
var Center = 0;
var Strike = 0;
var Sound = 0;
var Swf = 0;
var Ra = 0;
var Rm = 0;
var Marquee = 0;
var Fly = 0;
var fanzi=0;
var text_enter_url      = "������������ַ";
var text_enter_txt      = "����������˵��";
var text_enter_image    = "������ͼƬ��ַ";
var text_enter_sound    = "�����������ļ���ַ";
var text_enter_swf      = "������FLASH������ַ";
var text_enter_ra      = "������Real������ַ";
var text_enter_rm      = "������RealӰƬ��ַ";
var text_enter_wmv      = "������MediaӰƬ��ַ";
var text_enter_wma      = "������Media������ַ";
var text_enter_mov      = "������QuickTime������ַ";
var text_enter_sw      = "������shockwave������ַ";
var text_enter_email    = "�������ʼ���ַ";
var error_no_url        = "������������ַ";
var error_no_txt        = "����������˵��";
var error_no_title      = "������������ҳ����";
var error_no_email      = "�����������ʼ���ַ";
var error_no_gset       = "������ȷ���ո�ʽ���룡";
var error_no_gtxt       = "�����������֣�";
var text_enter_guang1   = "���ֵĳ��ȡ���ɫ�ͱ߽��С";
var text_enter_guang2   = "Ҫ����Ч�������֣�";
var text_enter_points    = "��������ֵ���磺1000 (�����ƻ�����1000�����µ��û��������������!)";
var error_no_points       = "������������ֵ���磺1000";
var text_enter_money    = "��������ֵ���磺1000 (�����ƽ�Ǯ��1000�����µ��û��������������!)";
var error_no_money       = "������������ֵ���磺1000";
var text_enter_power    = "��������ֵ���磺1000 (������������1000�����µ��û��������������!)";
var error_no_power       = "������������ֵ���磺1000";
var text_enter_post    = "��������ֵ���磺1000 (�����Ʒ���������1000�����µ��û��������������!)";
var error_no_post       = "������������ֵ���磺1000";
var text_enter_usercp    = "��������ֵ���磺1000 (����������ֵ��1000�����µ��û��������������!)";
var error_no_usercp       = "������������ֵ���磺1000";
var text_enter_usemoney    = "��������ֵ���磺50 (��ֻ���û�֧����50Ԫ�ֽ�������������!)";
var text_enter_usemoneyset    = "��������ֵ0��1��(������˹����Ƿ�ʹ�ö���֪ͨ�ң�0-�� 1-��)";
var error_no_usemoney       = "����������50���µ���ֵ���磺50";
var error_no_usemoneyset       = "����������0��1";
var text_enter_group    = "���������ɣ��磺�������� (��ֻ������Ϊ�������ɵ��û��������������!)";
var error_no_group       = "�����������������ƣ��磺��������";
var text_enter_userage    = "��������ֵ���磺30 (������ע��ʱ����30�����µ��û��������������!)";
var error_no_userage       = "������������ֵ���磺30";
var text_enter_sex    = "��������ֵ���磺1 (��ֻ�������û��������������!)";
var error_no_sex       = "������������ֵ0��1";
var text_enter_spec    = "�������û����磺�����ƽ�|С���� (��ֻ�������ƽ��С����������������!)";
var error_no_spec       = "�����������û����磺�����ƽ�|С����";
var text_enter_sec    = "�������û���ID���磺1|2|3 (��ֻ�й���Ա�����������Ͱ����������������!)";
var error_no_sec       = "�����������û����磺1|2|3";
var text_enter_timel    = "��������ֵ���磺2 (��2����û��������������!)";
var error_no_timel       = "������������ֵ���磺2";
var text_enter_timex    = "��������ֵ���磺2 (�������û�2��������������!)";
var error_no_timex       = "������������ֵ���磺2";

function commentWrite(NewCode) {
	document.frmAnnounce.Content.value+=NewCode;
	document.frmAnnounce.Content.focus();
	return;
}

function storeCaret(text) { 
	if (text.createTextRange) {
		text.caretPos = document.selection.createRange().duplicate();
	}
        if(event.ctrlKey && window.event.keyCode==13){i++;if (i>1) {alert('�������ڷ���,�����ĵȴ���');return false;}this.document.form.submit();}
}

function AddText(text) {
	if (document.frmAnnounce.Content.createTextRange && document.frmAnnounce.Content.caretPos) {      
		var caretPos = document.frmAnnounce.Content.caretPos;      
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
		text + ' ' : text;
	}
	else document.frmAnnounce.Content.value += text;
	document.frmAnnounce.Content.focus(caretPos);
}

function inputs(str) {
	AddText(str);
}

function Curl() {
	var FoundErrors = '';
	var enterURL   = prompt(text_enter_url, "http://");
	var enterTxT   = prompt(text_enter_txt, enterURL);
	if (!enterURL)    {
		FoundErrors += "\n" + error_no_url;
	}
	if (!enterTxT)    {
		FoundErrors += "\n" + error_no_txt;
	}
	if (FoundErrors)  {
		alert("����"+FoundErrors);
		return;
	}
	var ToAdd = "[URL="+enterURL+"]"+enterTxT+"[/URL]";
	document.frmAnnounce.Content.value+=ToAdd;
	document.frmAnnounce.Content.focus();
}

function Cimage() {
	var FoundErrors = '';
	var enterURL   = prompt(text_enter_image, "http://");
	if (!enterURL) {
		FoundErrors += "\n" + error_no_url;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	var ToAdd = "[IMG]"+enterURL+"[/IMG]";
	document.frmAnnounce.Content.value+=ToAdd;
	document.frmAnnounce.Content.focus();
}

function Cemail() {
	var emailAddress = prompt(text_enter_email,"");
	if (!emailAddress) { alert(error_no_email); return; }
	var ToAdd = "[EMAIL]"+emailAddress+"[/EMAIL]";
	commentWrite(ToAdd);
}

function Ccode() {
	if (Code == 0) {
		ToAdd = "[CODE]";
		document.form.code.value = " ����*";
		Code = 1;
	} else {
		ToAdd = "[/CODE]";
		document.form.code.value = " ���� ";
		Code = 0;
	}
	commentWrite(ToAdd);
}

function Cquote() {
	fontbegin="[QUOTE]";
	fontend="[/QUOTE]";
	fontchuli();
}

function Cbold() {
	fontbegin="[B]";
	fontend="[/B]";
	fontchuli();
}

function Citalic() {
	fontbegin="[I]";
	fontend="[/I]";
	fontchuli();
}

function Cunder() {
	fontbegin="[U]";
	fontend="[/U]";
	fontchuli();
}

function Csup() {
	fontbegin="[SUP]";
	fontend="[/SUP]";
	fontchuli();
}

function Csub() {
	fontbegin="[SUB]";
	fontend="[/SUB]";
	fontchuli();
}

function Ccenter() {
	fontbegin="[center]";
	fontend="[/center]";
	fontchuli();
}

function Cstrike() {
	fontbegin="[strike]";
	fontend="[/strike]";
	fontchuli();
}

function point() {
	var FoundErrors = '';
	var enterpoints  =prompt(text_enter_points,"1000");
	if (!enterpoints) {
		FoundErrors += "\n" + error_no_points;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[Point="+enterpoints+"]";
	fontend="[/Point]";
	fontchuli();
}

function money() {
	var FoundErrors = '';
	var entermoney  =prompt(text_enter_money,"1000");
	if (!entermoney) {
		FoundErrors += "\n" + error_no_money;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[Money="+entermoney+"]";
	fontend="[/Money]";
	fontchuli();
}

function usemoney() {
	var FoundErrors = '';
	var entermoney  =prompt(text_enter_usemoney,"50");
	var enterSet   = prompt(text_enter_usemoneyset, "1");
	if (!entermoney||entermoney>50) {
		FoundErrors += "\n" + error_no_usemoney;
	}
	if (!enterSet||(enterSet!=0&&enterSet!=1)) {
		FoundErrors += "\n" + error_no_usemoneyset;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[UseMoney="+entermoney+","+enterSet+"]";
	fontend="[/UseMoney]";
	fontchuli();
}

function usercp() {
	var FoundErrors = '';
	var enterusercp  =prompt(text_enter_usercp,"1000");
	if (!enterusercp) {
		FoundErrors += "\n" + error_no_usercp;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[UserCP="+enterusercp+"]";
	fontend="[/UserCP]";
	fontchuli();
}

function power() {
	var FoundErrors = '';
	var enterpower  =prompt(text_enter_power,"1000");
	if (!enterpower) {
		FoundErrors += "\n" + error_no_power;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[Power="+enterpower+"]";
	fontend="[/Power]";
	fontchuli();
}

function article() {
	var FoundErrors = '';
	var enterpost  =prompt(text_enter_post,"1000");
	if (!enterpost) {
		FoundErrors += "\n" + error_no_post;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[Post="+enterpost+"]";
	fontend="[/Post]";
	fontchuli();
}

function replyview() {
	fontbegin="[replyview]";
	fontend="[/replyview]";
	fontchuli();
}

function usermem() {
	fontbegin="[usermem]";
	fontend="[/usermem]";
	fontchuli();
}

function usergroup() {
	var FoundErrors = '';
	var enterpost  =prompt(text_enter_group,"��������");
	if (!enterpost) {
		FoundErrors += "\n" + error_no_group;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[usergroup="+enterpost+"]";
	fontend="[/usergroup]";
	fontchuli();
}

function userage() {
	var FoundErrors = '';
	var enterpost  =prompt(text_enter_userage,"30");
	if (!enterpost) {
		FoundErrors += "\n" + error_no_userage;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[userage="+enterpost+"]";
	fontend="[/userage]";
	fontchuli();
}

function sex() {
	var FoundErrors = '';
	var enterpost  =prompt(text_enter_sex,"1");
	if (!enterpost) {
		FoundErrors += "\n" + error_no_sex;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[sex="+enterpost+"]";
	fontend="[/sex]";
	fontchuli();
}

function uservip() {
	fontbegin="[uservip]";
	fontend="[/uservip]";
	fontchuli();
}

function sec() {
	var FoundErrors = '';
	var enterpost=prompt(text_enter_sec,"1|2|3");
	if (!enterpost) {
		FoundErrors += "\n" + error_no_spec;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[sec="+enterpost+"]";
	fontend="[/sec]";
	fontchuli();
}

function timel() {
	var FoundErrors = '';
	var enterpost  =prompt(text_enter_timel,"2");
	if (!enterpost) {
		FoundErrors += "\n" + error_no_timel;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[timel="+enterpost+"]";
	fontend="[/timel]";
	fontchuli();
}

function timex() {
	var FoundErrors = '';
	var enterpost  =prompt(text_enter_timex,"2");
	if (!enterpost) {
		FoundErrors += "\n" + error_no_timex;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[timex="+enterpost+"]";
	fontend="[/timex]";
	fontchuli();
}

function spec() {
	var FoundErrors = '';
	var enterpost  =prompt(text_enter_spec,"�����ƽ�|С����");
	if (!enterpost) {
		FoundErrors += "\n" + error_no_spec;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	fontbegin="[spec="+enterpost+"]";
	fontend="[/spec]";
	fontchuli();
}

function Csound() {
	var FoundErrors = '';
	var enterURL   = prompt(text_enter_sound, "http://");
	if (!enterURL) {
		FoundErrors += "\n" + error_no_url;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	var ToAdd = "[SOUND]"+enterURL+"[/SOUND]";
	document.frmAnnounce.Content.value+=ToAdd;
	document.frmAnnounce.Content.focus();
}

helpstat = false;
stprompt = true;
basic = false;

function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 0) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 2) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}

function Cswf() {
 	if (helpstat){
		alert("Flash\nFlash ����.\n�÷�: [flash=����, �߶�]Flash �ļ��ĵ�ַ[/flash]");
	} else if (basic) {
		AddTxt="[flash=500,350][/flash]";
		AddText(AddTxt);
	} else {                  
		txt2=prompt("flash����,�߶�","500,350"); 
		if (txt2!=null) {
                txt=prompt("Flash �ļ��ĵ�ַ","http://");
		if (txt!=null) {
                          if (txt2=="") {             
			AddTxt="[flash=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
               } else {
		        AddTxt="[flash="+txt2+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
		 }        
	    }  
       }
    }
}

function Crm() {
	if (helpstat) {
               alert("realplay\n����realplay�ļ�.\n�÷�: [rm=����, �߶�]�ļ���ַ[/rm]");
	} else if (basic) {
		AddTxt="[rm=500,350][/rm]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("��Ƶ�Ŀ���,�߶�","500,350"); 
		if (txt2!=null) {
			txt=prompt("��Ƶ�ļ��ĵ�ַ","������");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[rm=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/rm]";
					AddText(AddTxt);
				} else {
					AddTxt="[rm="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/rm]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function Cwmv() {
	if (helpstat) {
               alert("Media Player\n����Media Player�ļ�.\n�÷�: [mp=����, �߶�]�ļ���ַ[/mp]");
	} else if (basic) {
		AddTxt="[mp=500,350][/mp]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("��Ƶ�Ŀ���,�߶�","500,350"); 
		if (txt2!=null) {
			txt=prompt("��Ƶ�ļ��ĵ�ַ","������");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[mp=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/mp]";
					AddText(AddTxt);
				} else {
					AddTxt="[mp="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/mp]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function Cmov() {
	if (helpstat) {
               alert("QuickTime\n����QuickTime�ļ�.\n�÷�: [qt=����, �߶�]�ļ���ַ[/qt]");
	} else if (basic) {
		AddTxt="[qt=500,350][/qt]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("��Ƶ�Ŀ���,�߶�","500,350"); 
		if (txt2!=null) {
			txt=prompt("��Ƶ�ļ��ĵ�ַ","������");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[qt=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/qt]";
					AddText(AddTxt);
				} else {
					AddTxt="[qt="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/qt]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function Cdir() {
	if (helpstat) {
               alert("Shockwave\n����Shockwave�ļ�.\n�÷�: [dir=����, �߶�]�ļ���ַ[/dir]");
	} else if (basic) {
		AddTxt="[dir=500,350][/dir]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("Shockwave�ļ��Ŀ���,�߶�","500,350"); 
		if (txt2!=null) {
			txt=prompt("Shockwave�ļ��ĵ�ַ","�������ַ");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[dir=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/dir]";
					AddText(AddTxt);
				} else {
					AddTxt="[dir="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/dir]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function Cra() {
	var FoundErrors = '';
	var enterURL   = prompt(text_enter_ra, "http://");
	if (!enterURL) {
		FoundErrors += "\n" + error_no_url;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	var ToAdd = "[RA]"+enterURL+"[/RA]";
	document.frmAnnounce.Content.value+=ToAdd;
	document.frmAnnounce.Content.focus();
}

function Cfanzi() {
	fontbegin="[xray]";
	fontend="[/xray]";
	fontchuli();
}

function Cwma() {
	var FoundErrors = '';
	var enterURL   = prompt(text_enter_wma, "http://");
	if (!enterURL) {
		FoundErrors += "\n" + error_no_url;
	}
	if (FoundErrors) {
		alert("����"+FoundErrors);
		return;
	}
	var ToAdd = "[wma]"+enterURL+"[/wma]";
	document.frmAnnounce.Content.value+=ToAdd;
	document.frmAnnounce.Content.focus();
}

function Cmarquee() {
	fontbegin="[move]";
	fontend="[/move]";
	fontchuli();
}

function Cmark() {
	fontbegin="[mark]";
	fontend="[/mark]";
	fontchuli();
}
function Cfly() {
	fontbegin="[fly]";
	fontend="[/fly]";
	fontchuli();
}

function paste(text) {
	if (opener.document.frmAnnounce.Content.createTextRange && opener.document.frmAnnounce.Content.caretPos) {      
		var caretPos = opener.document.frmAnnounce.Content.caretPos;      
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
		text + ' ' : text;
	}
	else opener.document.frmAnnounce.Content.value += text;
	opener.document.frmAnnounce.Content.focus(caretPos);
}

function showsize(size){
	fontbegin="[size="+size+"]";
	fontend="[/size]";
	fontchuli();
}

function showfont(font){
	fontbegin="[face="+font+"]";
	fontend="[/face]";
	fontchuli();
}

function showcolor(color){
	fontbegin="[color="+color+"]";
	fontend="[/color]";
	fontchuli();
}

function fontchuli(){
	if ((document.selection)&&(document.selection.type == "Text")) {
		var range = document.selection.createRange();
		var ch_text=range.text;
		if(range.text.length>0)	range.text = fontbegin + ch_text + fontend;
		else document.frmAnnounce.Content.value=fontbegin+document.frmAnnounce.Content.value+fontend;
	} 
	else {
		document.frmAnnounce.Content.value=fontbegin+document.frmAnnounce.Content.value+fontend;
	}
	document.frmAnnounce.Content.focus();
}

function Cguang() {
	var FoundErrors = '';
	var enterSET   = prompt(text_enter_guang1, "255,red,2");
	var enterTxT   = prompt(text_enter_guang2, "����");
	if (!enterSET)    {
		FoundErrors += "\n" + error_no_gset;
	}
	if (!enterTxT)    {
		FoundErrors += "\n" + error_no_gtxt;
	}
	if (FoundErrors)  {
		alert("����"+FoundErrors);
		return;
	}
	var ToAdd = "[glow="+enterSET+"]"+enterTxT+"[/glow]";
	document.frmAnnounce.Content.value+=ToAdd;
	document.frmAnnounce.Content.focus();
}

function Cying() {
	var FoundErrors = '';
	var enterSET   = prompt(text_enter_guang1, "255,blue,1");
	var enterTxT   = prompt(text_enter_guang2, "����");
	if (!enterSET)    {
		FoundErrors += "\n" + error_no_gset;
	}
	if (!enterTxT)    {
		FoundErrors += "\n" + error_no_gtxt;
	}
	if (FoundErrors)  {
		alert("����"+FoundErrors);
		return;
	}
	var ToAdd = "[SHADOW="+enterSET+"]"+enterTxT+"[/SHADOW]";
	document.frmAnnounce.Content.value+=ToAdd;
	document.frmAnnounce.Content.focus();
}

ie = (document.all)? true:false
if (ie){
	function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.frmAnnounce.submit();}}
}

function DoTitle(addTitle) { 
	var revisedTitle; 
	var currentTitle = document.frmAnnounce.subject.value; 
	revisedTitle = currentTitle+addTitle; 
	document.frmAnnounce.subject.value=revisedTitle; 
	document.frmAnnounce.subject.focus(); 
	return; 
}

function DoTitle2(addTitle) { 
	var revisedTitle; 
	var currentTitle = document.frmAnnounce.subject.value; 
	revisedTitle = addTitle+currentTitle; 
	document.frmAnnounce.subject.value=revisedTitle; 
	document.frmAnnounce.subject.focus(); 
	return; 
}

function insertsmilie(smilieface) {
	document.frmAnnounce.Content.value+=smilieface;
}

function gopreview() {
	document.forms[1].title.value=document.forms[0].subject.value;
	document.forms[1].body.value=document.forms[0].Content.value;
	var popupWin = window.open('preview.asp', 'preview_page', 'scrollbars=yes,width=750,height=450');
	document.forms[1].submit()
}