<!--#include file="z_ubbcode.asp"-->
<%
'��ڲ���:strContent���ݣ�PostUserGroup�û���ID��PostTypeʹ�����ͣ�1��ʾ���ӣ�2��ʾ���棬3��ʾ���ţ�
function DvBCode(strContent,PostUserGroup,PostType)
'HTML Code
if PostType=3 or PostType=2 then
	redim Board_Setting(63)
	Board_Setting(6)=1
	Board_Setting(5)=0:Board_Setting(7)=1
	Board_Setting(8)=1:Board_Setting(9)=1
	Board_Setting(10)=0:Board_Setting(11)=0
	Board_Setting(12)=0:Board_Setting(13)=0
	Board_Setting(14)=0:Board_Setting(15)=0
	Board_Setting(23)=0:Board_Setting(53)=0
	Board_Setting(54)=0:Board_Setting(55)=0
	Board_Setting(56)=0:Board_Setting(57)=0
	Board_Setting(58)=0:Board_Setting(59)=0
	Board_Setting(60)=0:Board_Setting(61)=0
	Board_Setting(62)=0:Board_Setting(63)=0
end if
if Cint(Board_Setting(5)) <> 1 then
	strContent = dvHTMLEncode(strContent)
else
	strContent = HTMLcode(strContent)
end if
'UbbCode
if Cint(Board_Setting(6))<>1 and PostUserGroup>2 then
	DvBCode=strContent
	exit function
end if
dim re
dim po,ii
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

re.Pattern="(javascript)"
strContent=re.Replace(strContent,"<I>&#106avascript</I>")
re.Pattern="(jscript:)"
strContent=re.Replace(strContent,"<I>&#106script:</I>")
re.Pattern="(js:)"
strContent=re.Replace(strContent,"<I>&#106s:</I>")
re.Pattern="(value)"
strContent=re.Replace(strContent,"<I>&#118alue</I>")
re.Pattern="(about:)"
strContent=re.Replace(strContent,"<I>about&#58</I>")
re.Pattern="(file:)"
strContent=re.Replace(strContent,"<I>file&#58</I>")
re.Pattern="(document.cookie)"
strContent=re.Replace(strContent,"<I>documents&#46cookie</I>")
re.Pattern="(vbscript:)"
strContent=re.Replace(strContent,"<I>&#118bscript:</I>")
re.Pattern="(vbs:)"
strContent=re.Replace(strContent,"<I>&#118bs:</I>")
re.Pattern="(on(mouse|exit|error|click|key))"
strContent=re.Replace(strContent,"<I>&#111n$2</I>")

'UBB Code
strContent=UBB_UBB(strContent,PostUserGroup,Cint(Board_Setting(7)))

'IMG Code
strContent=UBB_IMG(strContent,PostUserGroup,Cint(Board_Setting(7)))

'UPLOAD Code
strContent=UBB_UPLOAD(strContent,PostUserGroup,Cint(Board_Setting(7)))

'DIR Code
strContent=UBB_DIR(strContent,PostUserGroup,Cint(Board_Setting(9)))

'QT Code
strContent=UBB_QT(strContent,PostUserGroup,Cint(Board_Setting(9)))

'MP Code
strContent=UBB_MP(strContent,PostUserGroup,Cint(Board_Setting(9)))

'RM Code
strContent=UBB_RM(strContent,PostUserGroup,Cint(Board_Setting(9)))

'FLASH Code
strContent=UBB_FLASH(strContent,PostUserGroup,Cint(Board_Setting(9)))

'SOUND Code
strContent=UBB_SOUND(strContent,PostUserGroup,Cint(Board_Setting(9)))

'URL Code
strContent=UBB_URL(strContent)

'EMAIL Code
strContent=UBB_EMAIL(strContent)

'HTML Code
strContent=UBB_HTML(strContent)

'CODE Code
strContent=UBB_CODE(strContent)

'COLOR Code
strContent=UBB_COLOR(strContent)

'FACE Code
strContent=UBB_FACE(strContent)

'CENTER Code
strContent=UBB_CENTER(strContent)

'ALIGN Code
strContent=UBB_ALIGN(strContent)

'FLY Code
strContent=UBB_FLY(strContent)

'MOVE Code
strContent=UBB_MOVE(strContent)

'SHADOW Code
strContent=UBB_SHADOW(strContent)

'GLOW Code
strContent=UBB_GLOW(strContent)

'I Code
strContent=UBB_I(strContent)

'U Code
strContent=UBB_U(strContent)

'B Code
strContent=UBB_B(strContent)

'SUP Code
strContent=UBB_SUP(strContent)

'SUB Code
strContent=UBB_SUB(strContent)

'SIZE Code
strContent=UBB_SIZE(strContent)

'QUOTE Code
strContent=UBB_QUOTE(strContent)

'�ֽ���
strContent=UBB_MONEY(strContent,PostUserGroup,PostType)

'������
strContent=UBB_POINT(strContent,PostUserGroup,PostType)

'������
strContent=UBB_USERCP(strContent,PostUserGroup,PostType)

'������
strContent=UBB_POWER(strContent,PostUserGroup,PostType)

'������
strContent=UBB_POST(strContent,PostUserGroup,PostType)

'�ظ���
strContent=UBB_REPLYVIEW(strContent,PostUserGroup,PostType)

'������
strContent=UBB_USEMONEY(strContent,PostUserGroup,PostType)

'������
strContent=UBB_SEC(strContent,PostUserGroup,PostType)

'��Ա��
strContent=UBB_USERMEM(strContent,PostUserGroup,PostType)

'������
strContent=UBB_SPEC(strContent,PostUserGroup,PostType)

'�Ա���
strContent=UBB_SEX(strContent,PostUserGroup,PostType)

'�߼���
strContent=UBB_USERVIP(strContent,PostUserGroup,PostType)

'������
strContent=UBB_USERAGE(strContent,PostUserGroup,PostType)

'������
strContent=UBB_USERGROUP(strContent,PostUserGroup,PostType)

'��ʾ��
strContent=UBB_TIMEL(strContent,PostUserGroup,PostType)

'������
strContent=UBB_TIMEX(strContent,PostUserGroup,PostType)

'�Զ�ʶ����ַ
re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@\]\':+!]+)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@\]\':+!]+)$"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "([^>=""(&quot;)])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@\]\':+!]+)"
strContent = re.Replace(strContent,"$1<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$2>$2</a>")

'�Զ�ʶ��www�ȿ�ͷ����ַ
're.Pattern = "([^(http://|http:\\)])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
re.Pattern = "([^(http://|http:\\|@)])((www)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=http://$2>$2</a>")

'�Զ�ʶ��Email��ַ����򿪱�������������ݺܶ�����ӻ����������ͣ��
're.Pattern = "([^(=)])((\w)+[@]{1}((\w)+[.]){1,3}(\w)+)"
'strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=""mailto:$2"">$2</a>")

'em code
if (Cint(Board_Setting(8)) = 1 or PostUserGroup<4) then
	for ii=0 to Forum_emotNum
		if len(ii)=1 then po="0" & ii else po=ii
		strContent=replace(strContent,"[em"&po&"]","<img src="&Forum_info(10)&Forum_emot(ii)&" border=0 align=middle>")
		strContent=replace(strContent,"[eM"&po&"]","<img src="&Forum_info(10)&Forum_emot(ii)&" border=0 align=middle>")
		strContent=replace(strContent,"[Em"&po&"]","<img src="&Forum_info(10)&Forum_emot(ii)&" border=0 align=middle>")
		strContent=replace(strContent,"[EM"&po&"]","<img src="&Forum_info(10)&Forum_emot(ii)&" border=0 align=middle>")
	next
else
	re.Pattern="\[em(.[^\[]*)\]"
	strContent=re.Replace(strContent,"")
end if

'����
'strContent=UBB_ACTION(strContent)

'MARK Code
strContent=UBB_MARK(strContent)

strContent=replace(strContent,"<I></I>","")

set re=Nothing
DvBcode=strContent
end function

function DvSignCode(strContent,PostUserGroup)
if Forum_Setting(66)<>1 then
	strContent = dvHTMLEncode(strContent)
else
	strContent = HTMLcode(strContent)
end if
if Forum_Setting(65)<>1 and PostUserGroup>2 then
	DvSignCode=strContent
	exit function
end if
dim re,po,ii
dim reContent,Test
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

re.Pattern="(javascript)"
strContent=re.Replace(strContent,"<I>&#106avascript</I>")
re.Pattern="(jscript:)"
strContent=re.Replace(strContent,"<I>&#106script:</I>")
re.Pattern="(js:)"
strContent=re.Replace(strContent,"<I>&#106s:</I>")
re.Pattern="(value)"
strContent=re.Replace(strContent,"<I>&#118alue</I>")
re.Pattern="(about:)"
strContent=re.Replace(strContent,"<I>about&#58</I>")
re.Pattern="(file:)"
strContent=re.Replace(strContent,"<I>file&#58</I>")
re.Pattern="(document.cookie)"
strContent=re.Replace(strContent,"<I>documents&#46cookie</I>")
re.Pattern="(vbscript:)"
strContent=re.Replace(strContent,"<I>&#118bscript:</I>")
re.Pattern="(vbs:)"
strContent=re.Replace(strContent,"<I>&#118bs:</I>")
re.Pattern="(on(mouse|exit|error|click|key))"
strContent=re.Replace(strContent,"<I>&#111n$2</I>")

'IMG Code
strContent=UBB_Sign_IMG(strContent,PostUserGroup,Cint(Forum_Setting(67)))

'FLASH Code
strContent=UBB_FLASH(strContent,PostUserGroup,Cint(Forum_Setting(67)))

'URL Code
strContent=UBB_URL(strContent)

'EMAIL Code
strContent=UBB_EMAIL(strContent)

'COLOR Code
strContent=UBB_COLOR(strContent)

'FACE Code
strContent=UBB_FACE(strContent)

'ALIGN Code
strContent=UBB_ALIGN(strContent)

'FLY Code
strContent=UBB_FLY(strContent)

'MOVE Code
strContent=UBB_MOVE(strContent)

'SHADOW Code
strContent=UBB_SHADOW(strContent)

'GLOW Code
strContent=UBB_GLOW(strContent)

'I Code
strContent=UBB_I(strContent)

'U Code
strContent=UBB_U(strContent)

'B Code
strContent=UBB_B(strContent)

'SUP Code
strContent=UBB_SUP(strContent)

'SUB Code
strContent=UBB_SUB(strContent)

'SIZE Code
strContent=UBB_SIZE(strContent)

'�Զ�ʶ����ַ
re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)$"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
strContent = re.Replace(strContent,"$1<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$2>$2</a>")

'�Զ�ʶ��www�ȿ�ͷ����ַ
re.Pattern = "([^(http://|http:\\)])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=http://$2>$2</a>")

'�Զ�ʶ��Email��ַ
're.Pattern = "([^(=)])((\w)+[@]{1}((\w)+[.]){1,3}(\w)+)"
'strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=""mailto:$2"">$2</a>")

for ii=0 to Forum_emotNum
	if len(ii)=1 then po="0" & ii else po=ii
	strContent=replace(strContent,"[em"&po&"]","<img src="&Forum_info(10)&Forum_emot(ii)&" border=0 align=middle>")
	strContent=replace(strContent,"[eM"&po&"]","<img src="&Forum_info(10)&Forum_emot(ii)&" border=0 align=middle>")
	strContent=replace(strContent,"[Em"&po&"]","<img src="&Forum_info(10)&Forum_emot(ii)&" border=0 align=middle>")
	strContent=replace(strContent,"[EM"&po&"]","<img src="&Forum_info(10)&Forum_emot(ii)&" border=0 align=middle>")
next

strContent=replace(strContent,"<I></I>","")

set re=Nothing
DvSignCode=strContent
end function

function reUBBCode(CurContent, dvflag)
dim strContent

strContent=CurContent
if dvflag then strContent = dvHTMLEncode(strContent)
dim re
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
strContent=replace(strContent,"&nbsp;"," ")

'COLOR Code
strContent=REUBB_COLOR(strContent)

'FACE Code
strContent=REUBB_FACE(strContent)

'ALIGN Code
strContent=REUBB_ALIGN(strContent)

'CENTER Code
strContent=REUBB_CENTER(strContent)

'FLY Code
strContent=REUBB_FLY(strContent)

'MOVE Code
strContent=REUBB_MOVE(strContent)

'SHADOW Code
strContent=REUBB_SHADOW(strContent)

'GLOW Code
strContent=REUBB_GLOW(strContent)

'I Code
strContent=REUBB_I(strContent)

'U Code
strContent=REUBB_U(strContent)

'B Code
strContent=REUBB_B(strContent)

'SIZE Code
strContent=REUBB_SIZE(strContent)

'QUOTE Code
strContent=REUBB_QUOTE(strContent)

'��Ǯ��
strContent=REUBB_MONEY(strContent)
'������
strContent=REUBB_POINT(strContent)
'������
strContent=REUBB_USERCP(strContent)
'������
strContent=REUBB_POWER(strContent)
'������
strContent=REUBB_POST(strContent)
'�ظ���
strContent=REUBB_REPLYVIEW(strContent)
'������
strContent=REUBB_USEMONEY(strContent)
'������
strContent=REUBB_SEC(strContent)
'��Ա��
strContent=REUBB_USERMEM(strContent)
'������
strContent=REUBB_SPEC(strContent)
'�Ա���
strContent=REUBB_SEX(strContent)
'�߼���
strContent=REUBB_USERVIP(strContent)
'������
strContent=REUBB_USERAGE(strContent)
'������
strContent=REUBB_USERGROUP(strContent)
'��ʾ��
strContent=REUBB_TIMEL(strContent)
'������
strContent=REUBB_TIMEX(strContent)

'����
'strContent=REUBB_ACTION(strContent)

'MARK Code
strContent=REUBB_MARK(strContent)

set re=Nothing
reUBBCode=strContent
end function

function dvHTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = replace(fString, "&#", "<I>&#</I>")
    fString = Replace(fString, CHR(32), "<I></I>&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    'fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    fString=ChkBadWords(fString)
    dvHTMLEncode = fString
end if
end function
%>
