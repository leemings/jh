<script Language="Javascript">
if(window==window.top){
var i=1;while(i<=800){
window.alert("��ȫ���棡");
i=i+1;}window.close();}
</script>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��������_���齭��</title>
</head>
<Style>
body {font-family:"����"; font-size: 9pt;}
A{COLOR: green; TEXT-DECORATION: none;font-size: 8pt}
A:hover {COLOR: #006600; TEXT-DECORATION: none}
td{font-family:verdana,arial,helvetica,Tahoma,"����"; font-size: 9pt}
</Style>
<script>
var Audibles_List = "|||�����°��������˶�֪��|||�ٺ٣������棡|||�����ˣ�����Ҫ�����|||�����ⲻ�����Ҵ����|||һ��Ҫ�Ҹ�ร�|||��ÿɰ�Ү��|||����ร�|||�Һ������ ~|||�㲻Ҫ������|||��ĺøм��㰡|||���ֱ̫���Ҹж��ˣ�|||����..����...�ٽ���...Kiss��|||ǿ��ʵ����ǿ��|||�����ô����Ҫ���ˣ�|||�����ĺ�����Ŷ��|||��������а�..|||��������~|||����ΰɣ�|||���ɲ��ɾ��ǲ��ɣ�|||������������Է�|||�������������Ҷ����ڽ�|||�ߣ�С�����ҿ��㻹�����ұ⣡|||̹�״ӿ���ͽ����˰ɣ�|||�ߣ�����ƭ��|||�Ҿ���ϲ���㶯���ǳ�����ʱ��...|||Hi!����ɣ���֪�����ڣ��ҿɿ������ˣ�|||��~~��������ĺܼ�į��|||���������ĵ�û��|||��ô˵�㻹�����ף�����ͷ����|||��������|||�ݰ�������|||���ϣ��ϰ����ˣ��������˰���|||���в��У��ҵ��ϲ���...���ϲ�����|||���������Σ�|||�����ò����˵�����һ����|||�����ڼ���|||���ᆪ��|||�簲��|||���ڼӰࣿ��æ��æ����|||��������ˣ��ҵ���ܾ��ˣ�|||����С�|||���Ϻã���������|||��Ҫ�����������ڲų��֣�������Ǳˮȥ��...|||ι..���ϰ����ˣ�����һ���~�´�����..|||������������һ��Ҫ������|||�źߡ��˼�.. ȥƲ����..�źߡ�|||�źߡ�������һ��~|||������һ���ɣ�follow me! one more, two more��|||ף�����տ���~ף�����տ���..�ร�ף�����տ����|||��֪����˵ʲô ..|||����˯�...zzz|||������һ��~~|||���ģ�������|||HELP ME !|||YEAH~~~~|||������|||������������|||��������������|||��������������롫����|||�ҵ�������"
Audibles_List=Audibles_List.split("|||");
function ShowForum_Emot(page){
	var CountLength = Audibles_List.length-1;
	var page_size = 9
	var showlist = ''
	var pagelist = ''
	var Page_Max = CountLength/page_size
	if ((CountLength % page_size)>0)Page_Max = Math.floor(Page_Max+1);
	for (i=page*page_size-page_size+1;i<=page*page_size;i++)
	{
		Audibles_ID = (i<10)?"0"+i:i ;
		Audibles_Url = "aq_swf/"+Audibles_ID + ".swf";
		Exww_Url = ""+Audibles_ID + ".swf";
		if (i<=CountLength)
		{
			showlist = showlist + '<tr><td width="5%">'+i+'</td><td width="40%"><OBJECT codeBase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="48" height="48" Id="Audible' + Audibles_ID +'"><PARAM NAME=movie VALUE="'+ Audibles_Url +'"><param name=menu value=false><PARAM NAME=quality VALUE=high><PARAM NAME=play VALUE=false><param name=""wmode"" value=""transparent""><embed src="' + Audibles_Url +'" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="48" height="48"></embed></OBJECT></td>'
			showlist = showlist + '<td width="25%" align="Left" style="cursor:hand;" title="' + Audibles_List[i] +'" nowrap  onclick="document.Audible' + Audibles_ID + '.Play()">[��]</td><td width="25%" style="cursor:hand;" title="�������">[<a href=javascript:s("'+ Exww_Url +'")>��</a>]<\/td><\/tr>'
		}
	}
	for (i=1;i<=Page_Max;i++)pagelist += (i==page)? '<font color=gray>['+i+']</font> ':'<A href="javascript:ShowForum_Emot('+i+')">['+i+']</A> '
	showlist = showlist + '<Tr><TD colspan="4">'+ pagelist +'</TD></TR>'
	showlist = '<table width="100%"  border="0" cellspacing="3" cellpadding="2">' + showlist + '<\/table>'
	document.getElementById("AudiblesShow").innerHTML = showlist ;
}
function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+"/��������$ "+list+" �㷢�Ͱ�ť����";parent.f2.document.af.sytemp.focus();}</script>
<body leftmargin="0" topmargin="0" background="bg.gif" oncontextmenu=self.event.returnValue=false>
<table width="" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" height="1">
<TR><TD id="AudiblesShow"></TD></TR>
</table>
<script>ShowForum_Emot(1)</script>
</body></html>