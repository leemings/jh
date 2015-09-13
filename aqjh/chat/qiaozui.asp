<script Language="Javascript">
if(window==window.top){
var i=1;while(i<=800){
window.alert("安全警告！");
i=i+1;}window.close();}
</script>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>巧嘴娃娃_快乐江湖</title>
</head>
<Style>
body {font-family:"宋体"; font-size: 9pt;}
A{COLOR: green; TEXT-DECORATION: none;font-size: 8pt}
A:hover {COLOR: #006600; TEXT-DECORATION: none}
td{font-family:verdana,arial,helvetica,Tahoma,"宋体"; font-size: 9pt}
</Style>
<script>
var Audibles_List = "|||就这事啊！地球人都知道|||嘿嘿！逗你玩！|||年轻人，做人要厚道！|||这干嘛？这不是那我打岔吗？|||一定要幸福喔！|||你好可爱耶！|||加油喔！|||我好想你喔 ~|||你不要这样嘛|||真的好感激你啊|||你简直太让我感动了！|||过来..近点...再近点...Kiss！|||强！实在是强！|||真真的么？我要晕了！|||最近真的好想你哦！|||啊你是中邪喔..|||你好讨厌喔~|||抽根菸吧！|||不干不干就是不干！|||饿不饿我请你吃饭|||听见了吗那是我肚子在叫|||哼！小样！我看你还真是找扁！|||坦白从宽，你就交代了吧！|||哼！你这骗子|||我就是喜欢你动作非常慢的时候...|||Hi!是你吧，我知道你在，我可看见你了！|||哎~~看来你真的很寂寞！|||年轻人虚心点没错|||这么说你还不明白，你猪头啊！|||掰掰啰！|||拜拜了您那|||哎呦！老板来了，我先走了啊！|||不行不行！我得上厕所...得上厕所！|||晚安做个好梦！|||我困得不行了得先走一步了|||有人在家吗？|||开会啰！|||早安！|||还在加班？真忙假忙啊？|||你可算来了，我等你很久了！|||来啦小妞|||早上好，您来啦！|||你要死啦～到现在才出现，到哪里潜水去啦...|||喂..我老板来了，先闪一下喽~下次再聊..|||啊～～～这是一定要的啦！|||嗯哼～人家.. 去撇个尿啰..嗯哼～|||嗯哼～来，亲一个~|||起来动一动吧，follow me! one more, two more…|||祝你生日快乐~祝你生日快乐..喔喔！祝你生日快乐喽|||不知道该说什么 ..|||好想睡喔...zzz|||让我想一想~~|||天哪！～～～|||HELP ME !|||YEAH~~~~|||哪里？哪里？|||哈哈！哈哈！|||晴天霹雳～～～|||这个问题让我想想～～！|||我倒～～！"
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
			showlist = showlist + '<td width="25%" align="Left" style="cursor:hand;" title="' + Audibles_List[i] +'" nowrap  onclick="document.Audible' + Audibles_ID + '.Play()">[听]</td><td width="25%" style="cursor:hand;" title="点击发言">[<a href=javascript:s("'+ Exww_Url +'")>言</a>]<\/td><\/tr>'
		}
	}
	for (i=1;i<=Page_Max;i++)pagelist += (i==page)? '<font color=gray>['+i+']</font> ':'<A href="javascript:ShowForum_Emot('+i+')">['+i+']</A> '
	showlist = showlist + '<Tr><TD colspan="4">'+ pagelist +'</TD></TR>'
	showlist = '<table width="100%"  border="0" cellspacing="3" cellpadding="2">' + showlist + '<\/table>'
	document.getElementById("AudiblesShow").innerHTML = showlist ;
}
function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+"/巧嘴娃娃$ "+list+" 点发送按钮即可";parent.f2.document.af.sytemp.focus();}</script>
<body leftmargin="0" topmargin="0" background="bg.gif" oncontextmenu=self.event.returnValue=false>
<table width="" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF" bgcolor="#FFFFFF" height="1">
<TR><TD id="AudiblesShow"></TD></TR>
</table>
<script>ShowForum_Emot(1)</script>
</body></html>