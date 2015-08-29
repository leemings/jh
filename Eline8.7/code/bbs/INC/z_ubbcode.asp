<%
function UBB_UBB(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test,po
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[UBB\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/UBB\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[UBB\]"
				strContent=re.replace(strContent, chr(1) & "UBB" & chr(2))
				re.Pattern="\[\/UBB\]"
				strContent=re.replace(strContent, chr(1) & "/UBB" & chr(2))
				re.Pattern="\x01UBB\x02\x01\/UBB\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="(^.*)\x01UBB\x02(.[^\x01]*)\x01\/UBB\x02(.*)"
				po=re.Replace(strContent,"$2")
				if instr(po,"[")>0 then
					po=replace(replace(po,"[","&#91;"),"]","&#93;")
					re.Pattern="(^.*)\x01UBB\x02(.[^\x01]*)\x01\/UBB\x02(.*)"
					strContent=re.replace(strContent, "$1<br><b>������UBB���룺</b><br>"&po&"<br>$3")
				else
					re.Pattern="(^.*)\x01UBB\x02(.[^\x01]*)\x01\/UBB\x02(.*)"
					strContent=re.replace(strContent, "$1$2$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_UBB=strContent
end function

function UBB_IMG(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[IMG\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/IMG\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[IMG\]"
			strContent=re.replace(strContent, chr(1) & "IMG" & chr(2))
			re.Pattern="\[\/IMG\]"
			strContent=re.replace(strContent, chr(1) & "/IMG" & chr(2))
			re.Pattern="\x01IMG\x02(.[^\x01]*)\x01\/IMG\x02"
			if Flag = 1 or PostUserGroup<4 then
				strContent=re.Replace(strContent,"<a onfocus=this.blur() href=""$1"" target=_blank id=""ImgSpan""><IMG SRC=""$1"" border=0 alt=�������´������ͼƬ></a>")
			else
				strContent=re.Replace(strContent,"<IMG SRC=""images/files/gif.gif"" border=0><a onfocus=this.blur() href=""$1"" target=_blank>$1</a>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_IMG=strContent
end function

function UBB_Sign_IMG(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[IMG\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/IMG\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[IMG\]"
			strContent=re.replace(strContent, chr(1) & "IMG" & chr(2))
			re.Pattern="\[\/IMG\]"
			strContent=re.replace(strContent, chr(1) & "/IMG" & chr(2))
			re.Pattern="\x01IMG\x02(.[^\x01]*)\x01\/IMG\x02"
			if Flag = 1 or PostUserGroup<4 then
				strContent=re.Replace(strContent,"<IMG SRC=""$1"" border=0 onload=""javascript:if(this.readyState=='complete') if(this.width>document.body.clientWidth-500)this.width=document.body.clientWidth-500"">")
			else
				strContent=re.Replace(strContent,"<IMG SRC=""images/files/gif.gif"" border=0><a onfocus=this.blur() href=""$1"" target=_blank>$1</a>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_Sign_IMG=strContent
end function

function UBB_UPLOAD(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/UPLOAD\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png)\]"
			strContent=re.replace(strContent, chr(1) & "UPLOAD=$1" & chr(2))
			re.Pattern="\[\/UPLOAD\]"
			strContent=re.replace(strContent, chr(1) & "/UPLOAD" & chr(2))
			re.Pattern="\x01UPLOAD=(gif|jpg|jpeg|bmp|png)\x02\x01\/UPLOAD\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01UPLOAD=(gif|jpg|jpeg|bmp|png)\x02(.[^\x01]*)\x01\/UPLOAD\x02"
			if Flag = 1 or PostUserGroup<4 then
				strContent= re.Replace(strContent,"<br><IMG SRC=""images/files/$1.gif"" border=0>���������ͼƬ���£�<br><A HREF=""$2"" TARGET=_blank id=""ImgSpan""><IMG SRC=""$2"" border=0 alt=�������´������ͼƬ></A>")
			else
				strContent= re.Replace(strContent,"<br><IMG SRC=""images/files/$1.gif"" border=0><A HREF=""$2"" TARGET=_blank>$2</A>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	re.Pattern="\[UPLOAD=(.[^\[]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/UPLOAD\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[UPLOAD=(.[^\[]*)\]"
			strContent=re.replace(strContent, chr(1) & "UPLOAD=$1" & chr(2))
			re.Pattern="\[\/UPLOAD\]"
			strContent=re.replace(strContent, chr(1) & "/UPLOAD" & chr(2))
			re.Pattern="\x01UPLOAD=(.[^\x01]*)\x02\x01\/UPLOAD\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01UPLOAD=(.[^\x01]*)\x02(.[^\x01]*)\x01\/UPLOAD\x02"
			strContent= re.Replace(strContent,"<br><IMG SRC=""images/files/$1.gif"" border=0> <a href=""$2"">���������ļ�</a>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_UPLOAD=strContent
end function

function UBB_DIR(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[DIR=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/DIR\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[DIR=*([0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "DIR=$1,$2" & chr(2))
			re.Pattern="\[\/DIR\]"
			strContent=re.replace(strContent, chr(1) & "/DIR" & chr(2))
			re.Pattern="\x01DIR=*([0-9]*),*([0-9]*)\x02\x01\/DIR\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01DIR=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/DIR\x02"
			if Flag=1 or PostUserGroup<4 then
				strContent=re.Replace(strContent,"<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>")
			else
				strContent=re.Replace(strContent,"<a href=$3 target=_blank>$3</a>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_DIR=strContent
end function

function UBB_QT(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[QT=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/QT\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[QT=*([0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "QT=$1,$2" & chr(2))
			re.Pattern="\[\/QT\]"
			strContent=re.replace(strContent, chr(1) & "/QT" & chr(2))
			re.Pattern="\x01QT=*([0-9]*),*([0-9]*)\x02\x01\/QT\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01QT=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/QT\x02"
			if Flag=1 or PostUserGroup<4 then
				strContent=re.Replace(strContent,"<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
			else
				strContent=re.Replace(strContent,"<a href=$3 target=_blank>$3</a>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_QT=strContent
end function

function UBB_MP(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[MP=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/MP\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[MP=*([0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "MP=$1,$2" & chr(2))
			re.Pattern="\[\/MP\]"
			strContent=re.replace(strContent, chr(1) & "/MP" & chr(2))
			re.Pattern="\x01MP=*([0-9]*),*([0-9]*)\x02\x01\/MP\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01MP=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/MP\x02"
			if Flag=1 or PostUserGroup<4 then
				strContent=re.Replace(strContent,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3 width=$1 height=$2></embed></object>")
			else
				strContent=re.Replace(strContent,"<a href=$3 target=_blank>$3</a>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_MP=strContent
end function

function UBB_RM(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[RM=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/RM\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[RM=*([0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "RM=$1,$2" & chr(2))
			re.Pattern="\[\/RM\]"
			strContent=re.replace(strContent, chr(1) & "/RM" & chr(2))
			re.Pattern="\x01RM=*([0-9]*),*([0-9]*)\x02\x01\/RM\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01RM=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/RM\x02"
			if Flag=1 or PostUserGroup<4 then
				strContent=re.Replace(strContent,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")
			else
				strContent=re.Replace(strContent,"<a href=$3 target=_blank>$3</a>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_RM=strContent
end function

function UBB_FLASH(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[FLASH\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/FLASH\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[FLASH\]"
			strContent=re.replace(strContent, chr(1) & "FLASH" & chr(2))
			re.Pattern="\[\/FLASH\]"
			strContent=re.replace(strContent, chr(1) & "/FLASH" & chr(2))
			re.Pattern="\x01FLASH\x02\x01\/FLASH\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01FLASH\x02(.[^\x01]*)\x01\/FLASH\x02"
			if Flag=1 or PostUserGroup<4 then
				strContent=re.Replace(strContent,"<a href=""$1"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=������´������͸�FLASH����! height=16 width=16>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$1""><PARAM NAME=quality VALUE=high><embed src=""$1"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$1</embed></OBJECT>")
			else
				strContent=re.Replace(strContent,"<IMG SRC="&Forum_info(7)&"swf.gif border=0><a href=$1 target=_blank>$1</a>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	re.Pattern="\[FLASH=*([0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/FLASH\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[FLASH=*([0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "FLASH=$1,$2" & chr(2))
			re.Pattern="\[\/FLASH\]"
			strContent=re.replace(strContent, chr(1) & "/FLASH" & chr(2))
			re.Pattern="\x01FLASH=*([0-9]*),*([0-9]*)\x02\x01\/FLASH\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01FLASH=*([0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/FLASH\x02"
			if Flag=1 or PostUserGroup<4 then
				strContent=re.Replace(strContent,"<a href=""$3"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=������´������͸�FLASH����! height=16 width=16>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$1 height=$2><PARAM NAME=movie VALUE=""$3""><PARAM NAME=quality VALUE=high><embed src=""$3"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$1 height=$2>$3</embed></OBJECT>")
			else
				strContent=re.Replace(strContent,"<IMG SRC="&Forum_info(7)&"swf.gif border=0><a href=$3 target=_blank>$3</a>")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_FLASH=strContent
end function

function UBB_SOUND(strText,PostUserGroup,Flag)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SOUND\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SOUND\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SOUND\]"
			strContent=re.replace(strContent, chr(1) & "SOUND" & chr(2))
			re.Pattern="\[\/SOUND\]"
			strContent=re.replace(strContent, chr(1) & "/SOUND" & chr(2))
			re.Pattern="\x01SOUND\x02\x01\/SOUND\x02"
			strContent=re.Replace(strContent,"")
			if Flag=1 or PostUserGroup<4 then
				re.Pattern="\x01SOUND\x02(.[^\x01]*)\x01\/SOUND\x02"
				strContent=re.Replace(strContent,"<a href=""$1"" target=_blank><IMG SRC=pic/mid.gif border=0 alt='��������'></a><bgsound src=""$1"" loop=""-1"">")
			end if
		end if
		re.Pattern="\x02"
		strContent=re.replace(strContent, "]")
		re.Pattern="\x01"
		strContent=re.replace(strContent, "[")
	end if
	set re=Nothing
	UBB_SOUND=strContent
end function

function UBB_URL(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[URL\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/URL\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[URL\]"
			strContent=re.replace(strContent, chr(1) & "URL" & chr(2))
			re.Pattern="\[\/URL\]"
			strContent=re.replace(strContent, chr(1) & "/URL" & chr(2))
			re.Pattern="\x01URL\x02\x01\/URL\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01URL\x02(.[^\x01]*)\x01\/URL\x02"
			strContent=re.Replace(strContent,"<A HREF=""$1"" TARGET=_blank>$1</A>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	re.Pattern="\[URL=(.[^\[]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/URL\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[URL=(.[^\[]*)\]"
			strContent=re.replace(strContent, chr(1) & "URL=$1" & chr(2))
			re.Pattern="\[\/URL\]"
			strContent=re.replace(strContent, chr(1) & "/URL" & chr(2))
			re.Pattern="\x01URL=(.[^\x01]*)\x02\x01\/URL\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01URL=(.[^\x01]*)\x02(.[^\x01]*)\x01\/URL\x02"
			strContent=re.Replace(strContent,"<A HREF=""$1"" TARGET=_blank>$2</A>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_URL=strContent
end function

function UBB_EMAIL(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[EMAIL\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/EMAIL\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[EMAIL\]"
			strContent=re.replace(strContent, chr(1) & "EMAIL" & chr(2))
			re.Pattern="\[\/EMAIL\]"
			strContent=re.replace(strContent, chr(1) & "/EMAIL" & chr(2))
			re.Pattern="\x01EMAIL\x02\x01\/EMAIL\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01EMAIL\x02(\S+\@.[^\x01]*)\x01\/EMAIL\x02"
			strContent=re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$1"">$1</A>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	re.Pattern="\[EMAIL=(\S+\@.[^\[]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/EMAIL\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[EMAIL=(\S+\@.[^\[]*)\]"
			strContent=re.replace(strContent, chr(1) & "EMAIL=$1" & chr(2))
			re.Pattern="\[\/EMAIL\]"
			strContent=re.replace(strContent, chr(1) & "/EMAIL" & chr(2))
			re.Pattern="\x01EMAIL=(\S+\@.[^\x01]*)\x02\x01\/EMAIL\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01EMAIL=(\S+\@.[^\x01]*)\x02(.[^\x01]*)\x01\/EMAIL\x02"
			strContent=re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$1"" TARGET=_blank>$2</A>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_EMAIL=strContent
end function

function UBB_HTML(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[HTML\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/HTML\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[HTML\]"
			strContent=re.replace(strContent, chr(1) & "HTML" & chr(2))
			re.Pattern="\[\/HTML\]"
			strContent=re.replace(strContent, chr(1) & "/HTML" & chr(2))
			re.Pattern="\x01HTML\x02\x01\/HTML\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01HTML\x02(.[^\x01]*)\x01\/HTML\x02"
			strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' class='"&abgcolor&"'><td><b>��������Ϊ�������:</b><br>$1</td></table>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_HTML=strContent
end function

function UBB_CODE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[CODE\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/CODE\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[CODE\]"
			strContent=re.replace(strContent, chr(1) & "CODE" & chr(2))
			re.Pattern="\[\/CODE\]"
			strContent=re.replace(strContent, chr(1) & "/CODE" & chr(2))
			re.Pattern="\x01CODE\x02\x01\/CODE\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01CODE\x02(.[^\x01]*)\x01\/CODE\x02"
			strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' class='"&abgcolor&"'><td><b>��������Ϊ�������:</b><br><script>function runEx(){var winEx=open();winEx.document.open(""text/html"", ""replace"");winEx.document.write(unescape(event.srcElement.parentElement.children[3].innerText));winEx.document.close();}</script><div style=""width:100%;border-style:ridge;"">$1</div><br><input onClick=runEx() type=button value=���д˴��� name=Button></td></table>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_CODE=strContent
end function

function UBB_COLOR(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while True
		re.Pattern="\[COLOR=(.[^\[]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/COLOR\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[COLOR=(.[^\[]*)\]"
				strContent=re.replace(strContent, chr(1) & "COLOR=$1" & chr(2))
				re.Pattern="\[\/COLOR\]"
				strContent=re.replace(strContent, chr(1) & "/COLOR" & chr(2))
				re.Pattern="\x01COLOR=(.[^\x01]*)\x02\x01\/COLOR\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="\x01COLOR=(.[^\x01]*)\x02(.[^\x01]*)\x01\/COLOR\x02"
				strContent=re.Replace(strContent,"<font color=$1>$2</font>")
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_COLOR=strContent
end function

function UBB_FACE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while True
		re.Pattern="\[FACE=(.[^\[]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/FACE\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[FACE=(.[^\[]*)\]"
				strContent=re.replace(strContent, chr(1) & "FACE=$1" & chr(2))
				re.Pattern="\[\/FACE\]"
				strContent=re.replace(strContent, chr(1) & "/FACE" & chr(2))
				re.Pattern="\x01FACE=(.[^\x01]*)\x02\x01\/FACE\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="\x01FACE=(.[^\x01]*)\x02(.[^\x01]*)\x01\/FACE\x02"
				strContent=re.Replace(strContent,"<font face=$1>$2</font>")
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_FACE=strContent
end function

function UBB_ALIGN(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[ALIGN=(center|left|right)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/ALIGN\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[ALIGN=(center|left|right)\]"
			strContent=re.replace(strContent, chr(1) & "ALIGN=$1" & chr(2))
			re.Pattern="\[\/ALIGN\]"
			strContent=re.replace(strContent, chr(1) & "/ALIGN" & chr(2))
			re.Pattern="\x01ALIGN=(center|left|right)\x02\x01\/ALIGN\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01ALIGN=(center|left|right)\x02(.[^\x01]*)\x01\/ALIGN\x02"
			strContent=re.Replace(strContent,"<div align=$1>$2</div>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_ALIGN=strContent
end function

function UBB_CENTER(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/CENTER\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[CENTER\]"
			strContent=re.replace(strContent, chr(1) & "CENTER" & chr(2))
			re.Pattern="\[\/CENTER\]"
			strContent=re.replace(strContent, chr(1) & "/CENTER" & chr(2))
			re.Pattern="\x01CENTER\x02\x01\/CENTER\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01CENTER\x02(.[^\x01]*)\x01\/CENTER\x02"
			strContent=re.Replace(strContent,"<div align=center>$1</div>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_CENTER=strContent
end function

function UBB_MARK(strText)
	dim strContent,strContent2
	dim re,Test,po,ii,jj,kk,ll
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[MARK\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/MARK\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[MARK\]"
				strContent=re.replace(strContent, chr(1) & "MARK" & chr(2))
				re.Pattern="\[\/MARK\]"
				strContent=re.replace(strContent, chr(1) & "/MARK" & chr(2))
				re.Pattern="(^.*)\x01MARK\x02\x01\/MARK\x02(.*)"
				strContent=re.Replace(strContent,"")
				re.Pattern="(^.*)\x01MARK\x02(.[^\x01]*)\x01\/MARK\x02(.*)"
				strContent2=re.Replace(strContent,"$2")
				ii=1
				do
					po=instr(ii,strContent2,"</P>")
					if po>0 then
						ll=right(strContent2,len(strContent2)-po+1)
						randomize
						kk=int(10*rnd+10)
						strContent2=left(strContent2,po-1)&"<font color=#fffff1>"
						for jj=1 to kk
							randomize
							strContent2=strContent2&chr(int(63*rnd+63))
						next
						ii=len(strContent2)+11
						strContent2=strContent2&"</font>"&ll
					end if
				loop while po>0 and ii+4<len(strContent2)
				ii=1
				do
					po=instr(ii,strContent2,"<BR>")
					if po>0 then
						ll=right(strContent2,len(strContent2)-po+1)
						randomize
						kk=int(10*rnd+10)
						strContent2=left(strContent2,po-1)&"<font color=#fffff1>"
						for jj=1 to kk
							randomize
							strContent2=strContent2&chr(int(63*rnd+63))
						next
						ii=len(strContent2)+11
						strContent2=strContent2&"</font>"&ll
					end if
				loop while po>0 and ii+4<len(strContent2)
				randomize
				kk=int(10*rnd+10)
				strContent2=strContent2&"<font color=#fffff1>"
				for jj=1 to kk
					randomize
					strContent2=strContent2&chr(int(63*rnd+63))
				next
				strContent2=strContent2&"</font>"
				strContent="<span onbeforecopy=""event.returnValue=false;"" onbeforecut=""event.returnValue=false;"" oncopy=""event.returnValue=false;window.clipboardData.setData('Text','"&forum_info(1)&"');"" oncut=""event.returnValue=false;window.clipboardData.setData('Text','"&forum_info(1)&"');"">"&re.Replace(strContent,"$1"&strContent2&"$3")&"</span>"
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_MARK=strContent
end function

function UBB_FLY(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[FLY\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/FLY\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[FLY\]"
			strContent=re.replace(strContent, chr(1) & "FLY" & chr(2))
			re.Pattern="\[\/FLY\]"
			strContent=re.replace(strContent, chr(1) & "/FLY" & chr(2))
			re.Pattern="\x01FLY\x02\x01\/FLY\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01FLY\x02(.[^\x01]*)\x01\/FLY\x02"
			strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_FLY=strContent
end function

function UBB_MOVE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[MOVE\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/MOVE\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[MOVE\]"
			strContent=re.replace(strContent, chr(1) & "MOVE" & chr(2))
			re.Pattern="\[\/MOVE\]"
			strContent=re.replace(strContent, chr(1) & "/MOVE" & chr(2))
			re.Pattern="\x01MOVE\x02\x01\/MOVE\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01MOVE\x02(.[^\x01]*)\x01\/MOVE\x02"
			strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$1</marquee>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_MOVE=strContent
end function

function UBB_SHADOW(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SHADOW\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "SHADOW=$1,$2,$3" & chr(2))
			re.Pattern="\[\/SHADOW\]"
			strContent=re.replace(strContent, chr(1) & "/SHADOW" & chr(2))
			re.Pattern="\x01SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\x02\x01\/SHADOW\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/SHADOW\x02"
			strContent=re.Replace(strContent,"<table width=$1><tr><td style=""filter:shadow(color=$2, strength=$3)"">$4</td></tr></table>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_SHADOW=strContent
end function

function UBB_GLOW(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/GLOW\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "GLOW=$1,$2,$3" & chr(2))
			re.Pattern="\[\/GLOW\]"
			strContent=re.replace(strContent, chr(1) & "/GLOW" & chr(2))
			re.Pattern="\x01GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\x02\x01\/GLOW\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/GLOW\x02"
			strContent=re.Replace(strContent,"<table width=$1><tr><td style=""filter:glow(color=$2, strength=$3)"">$4</td></tr></table>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_GLOW=strContent
end function

function UBB_I(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[I\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/I\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[I\]"
			strContent=re.replace(strContent, chr(1) & "I" & chr(2))
			re.Pattern="\[\/I\]"
			strContent=re.replace(strContent, chr(1) & "/I" & chr(2))
			re.Pattern="\x01I\x02\x01\/I\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01I\x02(.[^\x01]*)\x01\/I\x02"
			strContent=re.Replace(strContent,"<i>$1</i>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_I=strContent
end function

function UBB_U(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[U\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/U\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[U\]"
			strContent=re.replace(strContent, chr(1) & "U" & chr(2))
			re.Pattern="\[\/U\]"
			strContent=re.replace(strContent, chr(1) & "/U" & chr(2))
			re.Pattern="\x01U\x02\x01\/U\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01U\x02(.[^\x01]*)\x01\/U\x02"
			strContent=re.Replace(strContent,"<u>$1</u>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_U=strContent
end function

function UBB_B(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[B\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/B\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[B\]"
			strContent=re.replace(strContent, chr(1) & "B" & chr(2))
			re.Pattern="\[\/B\]"
			strContent=re.replace(strContent, chr(1) & "/B" & chr(2))
			re.Pattern="\x01B\x02\x01\/B\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01B\x02(.[^\x01]*)\x01\/B\x02"
			strContent=re.Replace(strContent,"<b>$1</b>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_B=strContent
end function

function UBB_SUP(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SUP\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SUP\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SUP\]"
			strContent=re.replace(strContent, chr(1) & "SUP" & chr(2))
			re.Pattern="\[\/SUP\]"
			strContent=re.replace(strContent, chr(1) & "/SUP" & chr(2))
			re.Pattern="\x01SUP\x02\x01\/SUP\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01SUP\x02(.[^\x01]*)\x01\/SUP\x02"
			strContent=re.Replace(strContent,"<SUP>$1</SUP>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_SUP=strContent
end function

function UBB_SUB(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SUB\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SUB\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SUB\]"
			strContent=re.replace(strContent, chr(1) & "SUB" & chr(2))
			re.Pattern="\[\/SUB\]"
			strContent=re.replace(strContent, chr(1) & "/SUB" & chr(2))
			re.Pattern="\x01SUB\x02\x01\/SUB\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01SUB\x02(.[^\x01]*)\x01\/SUB\x02"
			strContent=re.Replace(strContent,"<SUB>$1</SUB>")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_SUB=strContent
end function

function UBB_SIZE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while True
		re.Pattern="\[SIZE=([1-7])\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/SIZE\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[SIZE=([1-7])\]"
				strContent=re.replace(strContent, chr(1) & "SIZE=$1" & chr(2))
				re.Pattern="\[\/SIZE\]"
				strContent=re.replace(strContent, chr(1) & "/SIZE" & chr(2))
				re.Pattern="\x01SIZE=([1-7])\x02\x01\/SIZE\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="\x01SIZE=([1-7])\x02(.[^\x01]*)\x01\/SIZE\x02"
				strContent=re.Replace(strContent,"<font size=$1>$2</font>")
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_SIZE=strContent
end function

function UBB_QUOTE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[QUOTE\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/QUOTE\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[QUOTE\]"
			strContent=re.replace(strContent, chr(1) & "QUOTE" & chr(2))
			re.Pattern="\[\/QUOTE\]"
			strContent=re.replace(strContent, chr(1) & "/QUOTE" & chr(2))
			do
				re.Pattern="\x01QUOTE\x02\x01\/QUOTE\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="\x01QUOTE\x02(.[^\x01]*)\x01\/QUOTE\x02"
				strContent=re.Replace(strContent,"<table style=""width:100%"" cellpadding=5 cellspacing=1 class=tableborder1><TR><TD class="&abgcolor&" width=""100%"">$1</td></tr></table><br>")
				Test=re.Test(strContent)
			loop while(Test)
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_QUOTE=strContent
end function

function UBB_MONEY(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[MONEY=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/MONEY\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[MONEY=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "MONEY=$1" & chr(2))
				re.Pattern="\[\/MONEY\]"
				strContent=re.replace(strContent, chr(1) & "/MONEY" & chr(2))
				re.Pattern="(\x01MONEY=*([0-9]*)\x02)(\x01\/MONEY\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(10))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01MONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/MONEY\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or mymoney>=ii or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ�ֽ�ﵽ<B>$3</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ�ֽ�ﵽ<B>$3</B>�ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01MONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/MONEY\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_MONEY=strContent
end function

function UBB_POINT(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[POINT=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/POINT\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[POINT=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "POINT=$1" & chr(2))
				re.Pattern="\[\/POINT\]"
				strContent=re.replace(strContent, chr(1) & "/POINT" & chr(2))
				re.Pattern="(\x01POINT=*([0-9]*)\x02)(\x01\/POINT\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(11))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01POINT=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POINT\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or myuserscore>=ii or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ���ִﵽ<B>$3</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ���ִﵽ<B>$3</B>�ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01POINT=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POINT\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_POINT=strContent
end function

function UBB_USERCP(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[USERCP=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/USERCP\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[USERCP=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "USERCP=$1" & chr(2))
				re.Pattern="\[\/USERCP\]"
				strContent=re.replace(strContent, chr(1) & "/USERCP" & chr(2))
				re.Pattern="(\x01USERCP=*([0-9]*)\x02)(\x01\/USERCP\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(12))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01USERCP=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USERCP\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or myusercp>=ii or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ�����ﵽ<B>$3</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ�����ﵽ<B>$3</B>�ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01USERCP=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USERCP\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_USERCP=strContent
end function

function UBB_POWER(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[POWER=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/POWER\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[POWER=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "POWER=$1" & chr(2))
				re.Pattern="\[\/POWER\]"
				strContent=re.replace(strContent, chr(1) & "/POWER" & chr(2))
				re.Pattern="(\x01POWER=*([0-9]*)\x02)(\x01\/POWER\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(13))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01POWER=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POWER\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or mypower>=ii or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ�����ﵽ<B>$3</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ�����ﵽ<B>$3</B>�ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01POWER=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POWER\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_POWER=strContent
end function

function UBB_POST(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[POST=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/POST\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[POST=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "POST=$1" & chr(2))
				re.Pattern="\[\/POST\]"
				strContent=re.replace(strContent, chr(1) & "/POST" & chr(2))
				re.Pattern="(\x01POST=*([0-9]*)\x02)(\x01\/POST\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(14))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01POST=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POST\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or myarticle>=ii or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ�������ﵽ<B>$3</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ�������ﵽ<B>$3</B>�ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01POST=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POST\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_POST=strContent
end function

function UBB_REPLYVIEW(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim vrs
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[REPLYVIEW\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/REPLYVIEW\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[REPLYVIEW\]"
			strContent=re.replace(strContent, chr(1) & "REPLYVIEW" & chr(2))
			re.Pattern="\[\/REPLYVIEW\]"
			strContent=re.replace(strContent, chr(1) & "/REPLYVIEW" & chr(2))
			re.Pattern="(\x01REPLYVIEW\x02)(\x01\/REPLYVIEW\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01REPLYVIEW\x02)(.[^\x01]*)(\x01\/REPLYVIEW\x02)"
			if (Cint(Board_Setting(15))=1 or PostUserGroup<4) and PostType=1 then
				set vrs=conn.execute("select AnnounceID from "&TotalUseTable&" where rootid="&Announceid&" and PostUserID="&UserID)
				if membername<>"" and (not (vrs.eof or vrs.bof) or master) then
					strContent=re.Replace(strContent,"<hr noshade size=1><font color=gray>��������ֻ��<B>�ظ�</B>��ſ������</font><BR>$2<hr noshade size=1>")
				else
					strContent=re.Replace(strContent,"<hr noshade size=1><font color="&Forum_body(8)&">��������ֻ��<B>�ظ�</B>��ſ������</font><hr noshade size=1>")
				end if
				vrs.close
				set vrs=nothing
			else
				strContent=re.Replace(strContent,"$2")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_REPLYVIEW=strContent
end function

function UBB_USEMONEY(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii,iii
	dim SplitBuyUser,iPostBuyUser
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[USEMONEY=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/USEMONEY\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[USEMONEY=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "USEMONEY=$1" & chr(2))
				re.Pattern="\[\/USEMONEY\]"
				strContent=re.replace(strContent, chr(1) & "/USEMONEY" & chr(2))
				re.Pattern="(\x01USEMONEY=*([0-9]*)\x02)(\x01\/USEMONEY\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(23))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01USEMONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or master) then
						if (not isnull(PostBuyUser)) and PostBuyUser<>"" then
							SplitBuyUser=split(PostBuyUser,"|")
							iPostBuyUser="<option value=0>�ѹ����û�</option>"
							for iii=0 to ubound(SplitBuyUser)
								iPostBuyUser=iPostBuyUser & "<option value="&iii&">"&SplitBuyUser(iii)&"</option>"
							next
						else
							iPostBuyUser="<option value=0>��û���û�����</option>"
						end if
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ�����ֽ�<B>$3</B>�ſ������</font>&nbsp;&nbsp;<select size=1 name=buyuser>"&iPostBuyUser&"</select><BR>$4<hr noshade size=1>$6")
					else
						if (not isnull(PostBuyUser)) and PostBuyUser<>"" then
							if instr("|"&PostBuyUser&"|","|"&membername&"|")>0 then
								strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ�����ֽ�<B>$3</B>�ſ�����������Ѿ�������</font><BR>$4<hr noshade size=1>$6")
							else
								if mymoney>=ii then
									strContent=re.Replace(strContent,"$1<Form action=""BuyPost.asp"" mothod=post><font color="&Forum_body(8)&">����������Ҫ�����ֽ�<B>$3</B>�ſ������&nbsp;&nbsp;<input type=hidden name=boardid value="&boardid&"><input type=hidden value="&replyid&" name=replyid><input type=hidden value="&AnnounceID&" name=id><input type=hidden value="&totalusetable&" name=posttable><input type=submit name=submit value=�úڰ����ҡ������ˣ�>&nbsp;&nbsp;</font></form>$6")
								else
									strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ�����ֽ�<B>$3</B>�ſ����������û����ô���ֽ�</font><hr noshade size=1>$6")
								end if
							end if
						else
							if mymoney>=ii then
								strContent=re.Replace(strContent,"$1<Form action=""BuyPost.asp"" mothod=post><font color="&Forum_body(8)&">����������Ҫ�����ֽ�<B>$3</B>�ſ������&nbsp;&nbsp;<input type=hidden name=boardid value="&boardid&"><input type=hidden value="&replyid&" name=replyid><input type=hidden value="&AnnounceID&" name=id><input type=hidden value="&totalusetable&" name=posttable><input type=submit name=submit value=�úڰ����ҡ������ˣ�>&nbsp;&nbsp;</font></form>$6")
							else
								strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ�����ֽ�<B>$3</B>�ſ����������û����ô���ֽ�</font><hr noshade size=1>$6")
							end if
						end if
					end if
				else
					re.Pattern="(\x01USEMONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	do while true
		re.Pattern="\[USEMONEY=*([0-9]*),([0|1])\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/USEMONEY\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[USEMONEY=*([0-9]*),([0|1])\]"
				strContent=re.replace(strContent, chr(1) & "USEMONEY=$1,$2" & chr(2))
				re.Pattern="\[\/USEMONEY\]"
				strContent=re.replace(strContent, chr(1) & "/USEMONEY" & chr(2))
				re.Pattern="(\x01USEMONEY=*([0-9]*),([0|1])\x02)(\x01\/USEMONEY\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(23))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01USEMONEY=*([0-9]*),([0|1])\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or master) then
						if (not isnull(PostBuyUser)) and PostBuyUser<>"" then
							SplitBuyUser=split(PostBuyUser,"|")
							iPostBuyUser="<option value=0>�ѹ����û�</option>"
							for iii=0 to ubound(SplitBuyUser)
								iPostBuyUser=iPostBuyUser & "<option value="&iii&">"&SplitBuyUser(iii)&"</option>"
							next
						else
							iPostBuyUser="<option value=0>��û���û�����</option>"
						end if
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ�����ֽ�<B>$3</B>�ſ������</font>&nbsp;&nbsp;<select size=1 name=buyuser>"&iPostBuyUser&"</select><BR>$5<hr noshade size=1>$7")
					else
						if (not isnull(PostBuyUser)) and PostBuyUser<>"" then
							if instr("|"&PostBuyUser&"|","|"&membername&"|")>0 then
								strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ�����ֽ�<B>$3</B>�ſ�����������Ѿ�������</font><BR>$5<hr noshade size=1>$7")
							else
								if mymoney>=ii then
									strContent=re.Replace(strContent,"$1<Form action=""BuyPost.asp"" mothod=post><font color="&Forum_body(8)&">����������Ҫ�����ֽ�<B>$3</B>�ſ������&nbsp;&nbsp;<input type=hidden name=boardid value="&boardid&"><input type=hidden value="&replyid&" name=replyid><input type=hidden value="&AnnounceID&" name=id><input type=hidden value="&totalusetable&" name=posttable><input type=submit name=submit value=�úڰ����ҡ������ˣ�>&nbsp;&nbsp;</font></form>$7")
								else
									strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ�����ֽ�<B>$3</B>�ſ����������û����ô���ֽ�</font><hr noshade size=1>$7")
								end if
							end if
						else
							if mymoney>=ii then
								strContent=re.Replace(strContent,"$1<Form action=""BuyPost.asp"" mothod=post><font color="&Forum_body(8)&">����������Ҫ�����ֽ�<B>$3</B>�ſ������&nbsp;&nbsp;<input type=hidden name=boardid value="&boardid&"><input type=hidden value="&replyid&" name=replyid><input type=hidden value="&AnnounceID&" name=id><input type=hidden value="&totalusetable&" name=posttable><input type=submit name=submit value=�úڰ����ҡ������ˣ�>&nbsp;&nbsp;</font></form>$7")
							else
								strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ�����ֽ�<B>$3</B>�ſ����������û����ô���ֽ�</font><hr noshade size=1>$7")
							end if
						end if
					end if
				else
					re.Pattern="(\x01USEMONEY=*([0-9]*),([0|1])\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_USEMONEY=strContent
end function

function UBB_SEC(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test,po,postr
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[SEC=(.[^\[]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/SEC\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[SEC=(.[^\[]*)\]"
				strContent=re.replace(strContent, chr(1) & "SEC=$1" & chr(2))
				re.Pattern="\[\/SEC\]"
				strContent=re.replace(strContent, chr(1) & "/SEC" & chr(2))
				re.Pattern="(\x01SEC=(.[^\x01]*)\x02)(\x01\/SEC\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(53))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01SEC=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/SEC\x02)(.*)"
				  po=re.Replace(strContent,"$3")
				  if isnull(po) or po="" then po="1|2|3"
				  dim posplit
				  posplit=split(po,"|")
				  for i=0 to ubound(posplit)
				  	if not isInteger(posplit(i)) then
				  		exit for
				  	end if
				  next
				  if i<=ubound(posplit) then po="1|2|3"
					dim vrs
					set vrs=conn.execute("select title from [UserGroups] where usergroupid in ("&replace(po,"|",",")&")")
					if not vrs.bof and not vrs.eof then
						postr=""
					else
						postr=username
					end if
				  do while not vrs.eof
				  	if postr<>"" then postr=postr&"|"
				  	postr=postr&vrs(0)
				  	vrs.movenext
				  loop
				  vrs.close
				  set vrs=nothing
					if membername<>"" and (instr("|"&trim(po)&"|","|"&usergroupid&"|")>0 or master or membername=username) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>��������ֻ��<B>"&postr&"</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">��������ֻ��<B>"&postr&"</B>�ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01SEC=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/SEC\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_SEC=strContent
end function

function UBB_USERMEM(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[USERMEM\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USERMEM\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USERMEM\]"
			strContent=re.replace(strContent, chr(1) & "USERMEM" & chr(2))
			re.Pattern="\[\/USERMEM\]"
			strContent=re.replace(strContent, chr(1) & "/USERMEM" & chr(2))
			re.Pattern="(\x01USERMEM\x02)(\x01\/USERMEM\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USERMEM\x02)(.[^\x01]*)(\x01\/USERMEM\x02)"
			if (Cint(Board_Setting(54))=1 or PostUserGroup<4) and PostType=1 then
				if membername<>""  then
					strContent=re.Replace(strContent,"<hr noshade size=1><font color=gray>��������ֻ��<B>ע���Ա</B>�ſ������</font><BR>$2<hr noshade size=1>")
				else
					strContent=re.Replace(strContent,"<hr noshade size=1><font color="&Forum_body(8)&">��������ֻ��<B>ע���Ա</B>�ſ������</font><hr noshade size=1>")
				end if
			else
				strContent=re.Replace(strContent,"$2")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_USERMEM=strContent
end function

function UBB_SPEC(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test,po
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[SPEC=(.[^\[]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/SPEC\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[SPEC=(.[^\[]*)\]"
				strContent=re.replace(strContent, chr(1) & "SPEC=$1" & chr(2))
				re.Pattern="\[\/SPEC\]"
				strContent=re.replace(strContent, chr(1) & "/SPEC" & chr(2))
				re.Pattern="(\x01SPEC=(.[^\x01]*)\x02)(\x01\/SPEC\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(55))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01SPEC=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/SPEC\x02)(.*)"
				  po=re.Replace(strContent,"$3") 
					if membername<>"" and (instr("|"&lcase(trim(po))&"|","|"&lcase(trim(membername))&"|")>0 or UserName=membername or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>��������ֻ��<B>$3</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">��������ֻ��<B>$3</B>�ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01SPEC=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/SPEC\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_SPEC=strContent
end function

function UBB_SEX(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test,po
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[SEX=([0|1])\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/SEX\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[SEX=([0|1])\]"
				strContent=re.replace(strContent, chr(1) & "SEX=$1" & chr(2))
				re.Pattern="\[\/SEX\]"
				strContent=re.replace(strContent, chr(1) & "/SEX" & chr(2))
				re.Pattern="(\x01SEX=([0|1])\x02)(\x01\/SEX\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(56))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01SEX=([0|1])\x02)(.[^\x01]*)(\x01\/SEX\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if membername<>"" and (mysex=po or UserName=membername or master) then
						if po="1" then
							strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>��������ֻ��<B>˧��</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
						else
							strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>��������ֻ��<B>��Ů</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
						end if
					else
						if po="1" then
							strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">��������ֻ��<B>˧��</B>�ſ������</font><hr noshade size=1>$6")
						else
							strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">��������ֻ��<B>��Ů</B>�ſ������</font><hr noshade size=1>$6")
						end if
					end if
				else
					re.Pattern="(\x01SEX=([0|1])\x02)(.[^\x01]*)(\x01\/SEX\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_SEX=strContent
end function

function UBB_USERVIP(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[USERVIP\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USERVIP\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USERVIP\]"
			strContent=re.replace(strContent, chr(1) & "USERVIP" & chr(2))
			re.Pattern="\[\/USERVIP\]"
			strContent=re.replace(strContent, chr(1) & "/USERVIP" & chr(2))
			re.Pattern="(\x01USERVIP\x02)(\x01\/USERVIP\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USERVIP\x02)(.[^\x01]*)(\x01\/USERVIP\x02)"
			if (Cint(Board_Setting(57))=1 or PostUserGroup<4) and PostType=1 then
				if membername<>"" and ((not isnull(myvip) and myvip=1) or membername=username or master) then
					strContent=re.Replace(strContent,"<hr noshade size=1><font color=gray>��������ֻ��<B>VIP��Ա</B>�ſ������</font><BR>$2<hr noshade size=1>")
				else
					strContent=re.Replace(strContent,"<hr noshade size=1><font color="&Forum_body(8)&">��������ֻ��<B>VIP��Ա</B>�ſ������</font><hr noshade size=1>")
				end if
			else
				strContent=re.Replace(strContent,"$2")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_USERVIP=strContent
end function

function UBB_USERAGE(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[USERAGE=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/USERAGE\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[USERAGE=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "USERAGE=$1" & chr(2))
				re.Pattern="\[\/USERAGE\]"
				strContent=re.replace(strContent, chr(1) & "/USERAGE" & chr(2))
				re.Pattern="(\x01USERAGE=*([0-9]*)\x02)(\x01\/USERAGE\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(58))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01USERAGE=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USERAGE\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or datediff("d",myaddDAte,Now())>=ii or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫע��ʱ��ﵽ<B>$3</B>��ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫע��ʱ��ﵽ<B>$3</B>��ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01USERAGE=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USERAGE\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_USERAGE=strContent
end function

function UBB_USERGROUP(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[USERGROUP=(.[^\[]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USERGROUP\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USERGROUP=(.[^\[]*)\]"
			strContent=re.replace(strContent, chr(1) & "USERGROUP=$1" & chr(2))
			re.Pattern="\[\/USERGROUP\]"
			strContent=re.replace(strContent, chr(1) & "/USERGROUP" & chr(2))
			re.Pattern="(\x01USERGROUP=(.[^\x01]*)\x02)(\x01\/USERGROUP\x02)"
			strContent=re.Replace(strContent,"")
			if (Cint(Board_Setting(59))=1 or PostUserGroup<4) and PostType=1 then
				re.Pattern="(^.*)(\x01USERGROUP=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/USERGROUP\x02)(.*)"
				if membername<>"" and (myusergroup=re.Replace(strContent,"$3") or UserName=membername or master) then
					strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>����������Ҫ������<B>$3</B>�ſ������</font><BR>$4<hr noshade size=1>$6")
				else
					strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">����������Ҫ������<B>$3</B>�ſ������</font><hr noshade size=1>$6")
				end if
			else
				re.Pattern="(\x01USERGROUP=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/USERGROUP\x02)"
				strContent=re.Replace(strContent,"$3")
			end if
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	UBB_USERGROUP=strContent
end function

function UBB_TIMEL(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[TIMEL=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/TIMEL\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[TIMEL=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "TIMEL=$1" & chr(2))
				re.Pattern="\[\/TIMEL\]"
				strContent=re.replace(strContent, chr(1) & "/TIMEL" & chr(2))
				re.Pattern="(\x01TIMEL=*([0-9]*)\x02)(\x01\/TIMEL\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(60))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01TIMEL=*([0-9]*)\x02)(.[^\x01]*)(\x01\/TIMEL\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or datediff("d",session("dateandtime"),Now())>=ii or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>�������ݷ���<B>$3</B>���ſ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">�������ݷ���<B>$3</B>���ſ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01TIMEL=*([0-9]*)\x02)(.[^\x01]*)(\x01\/TIMEL\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_TIMEL=strContent
end function

function UBB_TIMEX(strText,PostUserGroup,PostType)
	dim strContent
	dim re,Test
	dim po,ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while true
		re.Pattern="\[TIMEX=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/TIMEX\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[TIMEX=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "TIMEX=$1" & chr(2))
				re.Pattern="\[\/TIMEX\]"
				strContent=re.replace(strContent, chr(1) & "/TIMEX" & chr(2))
				re.Pattern="(\x01TIMEX=*([0-9]*)\x02)(\x01\/TIMEX\x02)"
				strContent=re.Replace(strContent,"")
				if (Cint(Board_Setting(61))=1 or PostUserGroup<4) and PostType=1 then
					re.Pattern="(^.*)(\x01TIMEX=*([0-9]*)\x02)(.[^\x01]*)(\x01\/TIMEX\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					if membername<>"" and (membername=username or datediff("d",session("dateandtime"),Now())<ii or master) then
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color=gray>�������ݷ�����<B>$3</B>���ڿ������</font><BR>$4<hr noshade size=1>$6")
					else
						strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">�������ݷ�����<B>$3</B>���ڿ������</font><hr noshade size=1>$6")
					end if
				else
					re.Pattern="(\x01TIMEX=*([0-9]*)\x02)(.[^\x01]*)(\x01\/TIMEX\x02)"
					strContent=re.Replace(strContent,"$3")
				end if
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	UBB_TIMEX=strContent
end function

function REUBB_COLOR(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while True
		re.Pattern="\[COLOR=(.[^\[]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/COLOR\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[COLOR=(.[^\[]*)\]"
				strContent=re.replace(strContent, chr(1) & "COLOR=$1" & chr(2))
				re.Pattern="\[\/COLOR\]"
				strContent=re.replace(strContent, chr(1) & "/COLOR" & chr(2))
				re.Pattern="\x01COLOR=(.[^\x01]*)\x02\x01\/COLOR\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="\x01COLOR=(.[^\x01]*)\x02(.[^\x01]*)\x01\/COLOR\x02"
				strContent=re.Replace(strContent,"$2")
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	REUBB_COLOR=strContent
end function

function REUBB_FACE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while True
		re.Pattern="\[FACE=(.[^\[]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/FACE\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[FACE=(.[^\[]*)\]"
				strContent=re.replace(strContent, chr(1) & "FACE=$1" & chr(2))
				re.Pattern="\[\/FACE\]"
				strContent=re.replace(strContent, chr(1) & "/FACE" & chr(2))
				re.Pattern="\x01FACE=(.[^\x01]*)\x02\x01\/FACE\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="\x01FACE=(.[^\x01]*)\x02(.[^\x01]*)\x01\/FACE\x02"
				strContent=re.Replace(strContent,"$2")
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	REUBB_FACE=strContent
end function

function REUBB_ALIGN(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[ALIGN=(center|left|right)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/ALIGN\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[ALIGN=(center|left|right)\]"
			strContent=re.replace(strContent, chr(1) & "ALIGN=$1" & chr(2))
			re.Pattern="\[\/ALIGN\]"
			strContent=re.replace(strContent, chr(1) & "/ALIGN" & chr(2))
			re.Pattern="\x01ALIGN=(center|left|right)\x02\x01\/ALIGN\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01ALIGN=(center|left|right)\x02(.[^\x01]*)\x01\/ALIGN\x02"
			strContent=re.Replace(strContent,"$2")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_ALIGN=strContent
end function

function REUBB_CENTER(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[CENTER\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/CENTER\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[CENTER\]"
			strContent=re.replace(strContent, chr(1) & "CENTER" & chr(2))
			re.Pattern="\[\/CENTER\]"
			strContent=re.replace(strContent, chr(1) & "/CENTER" & chr(2))
			re.Pattern="\x01CETNER\x02\x01\/CENTER\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01CETNER\x02(.[^\x01]*)\x01\/CENTER\x02"
			strContent=re.Replace(strContent,"$1")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_CENTER=strContent
end function

function REUBB_MARK(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[MARK\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/MARK\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[MARK\]"
			strContent=re.replace(strContent, chr(1) & "MARK" & chr(2))
			re.Pattern="\[\/MARK\]"
			strContent=re.replace(strContent, chr(1) & "/MARK" & chr(2))
			re.Pattern="\x01MARK\x02\x01\/MARK\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01MARK\x02(.[^\x01]*)\x01\/MARK\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_MARK=strContent
end function

function REUBB_FLY(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[FLY\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/FLY\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[FLY\]"
			strContent=re.replace(strContent, chr(1) & "FLY" & chr(2))
			re.Pattern="\[\/FLY\]"
			strContent=re.replace(strContent, chr(1) & "/FLY" & chr(2))
			re.Pattern="\x01FLY\x02\x01\/FLY\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01FLY\x02(.[^\x01]*)\x01\/FLY\x02"
			strContent=re.Replace(strContent,"$1")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_FLY=strContent
end function

function REUBB_MOVE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[MOVE\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/MOVE\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[MOVE\]"
			strContent=re.replace(strContent, chr(1) & "MOVE" & chr(2))
			re.Pattern="\[\/MOVE\]"
			strContent=re.replace(strContent, chr(1) & "/MOVE" & chr(2))
			re.Pattern="\x01MOVE\x02\x01\/MOVE\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01MOVE\x02(.[^\x01]*)\x01\/MOVE\x02"
			strContent=re.Replace(strContent,"$1")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_MOVE=strContent
end function

function REUBB_SHADOW(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SHADOW\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "SHADOW=$1,$2,$3" & chr(2))
			re.Pattern="\[\/SHADOW\]"
			strContent=re.replace(strContent, chr(1) & "/SHADOW" & chr(2))
			re.Pattern="\x01SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\x02\x01\/SHADOW\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/SHADOW\x02"
			strContent=re.Replace(strContent,"$4")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_SHADOW=strContent
end function

function REUBB_GLOW(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/GLOW\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "GLOW=$1,$2,$3" & chr(2))
			re.Pattern="\[\/GLOW\]"
			strContent=re.replace(strContent, chr(1) & "/GLOW" & chr(2))
			re.Pattern="\x01GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\x02\x01\/GLOW\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\x02(.[^\x01]*)\x01\/GLOW\x02"
			strContent=re.Replace(strContent,"$4")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_GLOW=strContent
end function

function REUBB_I(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[I\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/I\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[I\]"
			strContent=re.replace(strContent, chr(1) & "I" & chr(2))
			re.Pattern="\[\/I\]"
			strContent=re.replace(strContent, chr(1) & "/I" & chr(2))
			re.Pattern="\x01I\x02\x01\/I\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01I\x02(.[^\x01]*)\x01\/I\x02"
			strContent=re.Replace(strContent,"$1")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_I=strContent
end function

function REUBB_U(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[U\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/U\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[U\]"
			strContent=re.replace(strContent, chr(1) & "U" & chr(2))
			re.Pattern="\[\/U\]"
			strContent=re.replace(strContent, chr(1) & "/U" & chr(2))
			re.Pattern="\x01U\x02\x01\/U\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01U\x02(.[^\x01]*)\x01\/U\x02"
			strContent=re.Replace(strContent,"$1")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_U=strContent
end function

function REUBB_B(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[B\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/B\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[B\]"
			strContent=re.replace(strContent, chr(1) & "B" & chr(2))
			re.Pattern="\[\/B\]"
			strContent=re.replace(strContent, chr(1) & "/B" & chr(2))
			re.Pattern="\x01B\x02\x01\/B\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01B\x02(.[^\x01]*)\x01\/B\x02"
			strContent=re.Replace(strContent,"$1")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_B=strContent
end function

function REUBB_SUP(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SUP\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SUP\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SUP\]"
			strContent=re.replace(strContent, chr(1) & "SUP" & chr(2))
			re.Pattern="\[\/SUP\]"
			strContent=re.replace(strContent, chr(1) & "/SUP" & chr(2))
			re.Pattern="\x01SUP\x02\x01\/SUP\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01SUP\x02(.[^\x01]*)\x01\/SUP\x02"
			strContent=re.Replace(strContent,"$1")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_SUP=strContent
end function

function REUBB_SUB(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SUB\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SUB\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SUB\]"
			strContent=re.replace(strContent, chr(1) & "SUB" & chr(2))
			re.Pattern="\[\/SUB\]"
			strContent=re.replace(strContent, chr(1) & "/SUB" & chr(2))
			re.Pattern="\x01SUP\x02\x01\/SUB\x02"
			strContent=re.Replace(strContent,"")
			re.Pattern="\x01SUP\x02(.[^\x01]*)\x01\/SUB\x02"
			strContent=re.Replace(strContent,"$1")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_SUP=strContent
end function

function REUBB_SIZE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	do while True
		re.Pattern="\[SIZE=([1-7])\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/SIZE\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[SIZE=([1-7])\]"
				strContent=re.replace(strContent, chr(1) & "SIZE=$1" & chr(2))
				re.Pattern="\[\/SIZE\]"
				strContent=re.replace(strContent, chr(1) & "/SIZE" & chr(2))
				re.Pattern="\x01SIZE=([1-7])\x02\x01\/SIZE\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="\x01SIZE=([1-7])\x02(.[^\x01]*)\x01\/SIZE\x02"
				strContent=re.Replace(strContent,"$2")
				re.Pattern="\x02"
				strContent=re.replace(strContent, "]")
				re.Pattern="\x01"
				strContent=re.replace(strContent, "[")
			else
				exit do
			end if
		else
			exit do
		end if
	loop
	set re=Nothing
	REUBB_SIZE=strContent
end function

function REUBB_QUOTE(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[QUOTE\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/QUOTE\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[QUOTE\]"
			strContent=re.replace(strContent, chr(1) & "QUOTE" & chr(2))
			re.Pattern="\[\/QUOTE\]"
			strContent=re.replace(strContent, chr(1) & "/QUOTE" & chr(2))
			do
				re.Pattern="\x01QUOTE\x02\x01\/QUOTE\x02"
				strContent=re.Replace(strContent,"")
				re.Pattern="\x01QUOTE\x02(.[^\x01]*)\x01\/QUOTE\x02"
				strContent=re.Replace(strContent,"")
				Test=re.Test(strContent)
			loop while(Test)
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_QUOTE=strContent
end function

function REUBB_MONEY(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[MONEY=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/MONEY\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[MONEY=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "MONEY=$1" & chr(2))
			re.Pattern="\[\/MONEY\]"
			strContent=re.replace(strContent, chr(1) & "/MONEY" & chr(2))
			re.Pattern="(\x01MONEY=*([0-9]*)\x02)(\x01\/MONEY\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01MONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/MONEY\x02)"
			strContent=re.Replace(strContent,"����Ǯ������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_MONEY=strContent
end function

function REUBB_POINT(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[POINT=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/POINT\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[POINT=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "POINT=$1" & chr(2))
			re.Pattern="\[\/POINT\]"
			strContent=re.replace(strContent, chr(1) & "/POINT" & chr(2))
			re.Pattern="(\x01POINT=*([0-9]*)\x02)(\x01\/POINT\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01POINT=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POINT\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_POINT=strContent
end function

function REUBB_USERCP(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[USERCP=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USERCP\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USERCP=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "USERCP=$1" & chr(2))
			re.Pattern="\[\/USERCP\]"
			strContent=re.replace(strContent, chr(1) & "/USERCP" & chr(2))
			re.Pattern="(\x01USERCP=*([0-9]*)\x02)(\x01\/USERCP\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USERCP=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USERCP\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_USERCP=strContent
end function

function REUBB_POWER(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[POWER=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/POWER\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[POWER=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "POWER=$1" & chr(2))
			re.Pattern="\[\/POWER\]"
			strContent=re.replace(strContent, chr(1) & "/POWER" & chr(2))
			re.Pattern="(\x01POWER=*([0-9]*)\x02)(\x01\/POWER\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01POWER=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POWER\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_POWER=strContent
end function

function REUBB_POST(strText)
	dim strContent
	dim re,Test
	dim ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[POST=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/POST\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[POST=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "POST=$1" & chr(2))
			re.Pattern="\[\/POST\]"
			strContent=re.replace(strContent, chr(1) & "/POST" & chr(2))
			re.Pattern="(\x01POST=*([0-9]*)\x02)(\x01\/POST\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01POST=*([0-9]*)\x02)(.[^\x01]*)(\x01\/POST\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_POST=strContent
end function

function REUBB_REPLYVIEW(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[REPLYVIEW\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/REPLYVIEW\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[REPLYVIEW\]"
			strContent=re.replace(strContent, chr(1) & "REPLYVIEW" & chr(2))
			re.Pattern="\[\/REPLYVIEW\]"
			strContent=re.replace(strContent, chr(1) & "/REPLYVIEW" & chr(2))
			re.Pattern="(\x01REPLYVIEW\x02)(\x01\/REPLYVIEW\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01REPLYVIEW\x02)(.[^\x01]*)(\x01\/REPLYVIEW\x02)"
			strContent=re.Replace(strContent,"���ظ�������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_REPLYVIEW=strContent
end function

function REUBB_USEMONEY(strText)
	dim strContent
	dim re,Test
	dim ii,iii
	dim SplitBuyUser,iPostBuyUser
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[USEMONEY=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USEMONEY\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USEMONEY=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "USEMONEY=$1" & chr(2))
			re.Pattern="\[\/USEMONEY\]"
			strContent=re.replace(strContent, chr(1) & "/USEMONEY" & chr(2))
			re.Pattern="(\x01USEMONEY=*([0-9]*)\x02)(\x01\/USEMONEY\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USEMONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	re.Pattern="\[USEMONEY=*([0-9]*),([0|1])\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USEMONEY\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USEMONEY=*([0-9]*),([0|1])\]"
			strContent=re.replace(strContent, chr(1) & "USEMONEY=$1,$2" & chr(2))
			re.Pattern="\[\/USEMONEY\]"
			strContent=re.replace(strContent, chr(1) & "/USEMONEY" & chr(2))
			re.Pattern="(\x01USEMONEY=*([0-9]*),([0|1])\x02)(\x01\/USEMONEY\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USEMONEY=*([0-9]*),([0|1])\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_USEMONEY=strContent
end function

function REUBB_SEC(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SEC=(.[^\[]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SEC\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SEC=(.[^\[]*)\]"
			strContent=re.replace(strContent, chr(1) & "SEC=$1" & chr(2))
			re.Pattern="\[\/SEC\]"
			strContent=re.replace(strContent, chr(1) & "/SEC" & chr(2))
			re.Pattern="(\x01SEC=(.[^\x01]*)\x02)(\x01\/SEC\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01SEC=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/SEC\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_SEC=strContent
end function

function REUBB_USERMEM(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USERMEM\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USERMEM\]"
			strContent=re.replace(strContent, chr(1) & "USERMEM" & chr(2))
			re.Pattern="\[\/USERMEM\]"
			strContent=re.replace(strContent, chr(1) & "/USERMEM" & chr(2))
			re.Pattern="(\x01USERMEM\x02)(\x01\/USERMEM\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USERMEM\x02)(.[^\x01]*)(\x01\/USERMEM\x02)"
			strContent=re.Replace(strContent,"����Ա������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_USERMEM=strContent
end function

function REUBB_SPEC(strText)
	dim strContent
	dim re,Test
		
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SPEC=(.[^\[]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SPEC\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SPEC=(.[^\[]*)\]"
			strContent=re.replace(strContent, chr(1) & "SPEC=$1" & chr(2))
			re.Pattern="\[\/SPEC\]"
			strContent=re.replace(strContent, chr(1) & "/SPEC" & chr(2))
			re.Pattern="(\x01SPEC=(.[^\x01]*)\x02)(\x01\/SPEC\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01SPEC=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/SPEC\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_SPEC=strContent
end function

function REUBB_SEX(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[SEX=([0|1])\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/SEX\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[SEX=([0|1])\]"
			strContent=re.replace(strContent, chr(1) & "SEX=$1" & chr(2))
			re.Pattern="\[\/SEX\]"
			strContent=re.replace(strContent, chr(1) & "/SEX" & chr(2))
			re.Pattern="(\x01SEX=([0|1])\x02)(\x01\/SEX\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01SEX=([0|1])\x02)(.[^\x01]*)(\x01\/SEX\x02)"
			strContent=re.Replace(strContent,"���Ա�������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_SEX=strContent
end function

function REUBB_USERVIP(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[USERVIP\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USERVIP\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USERVIP\]"
			strContent=re.replace(strContent, chr(1) & "USERVIP" & chr(2))
			re.Pattern="\[\/USERVIP\]"
			strContent=re.replace(strContent, chr(1) & "/USERVIP" & chr(2))
			re.Pattern="(\x01USERVIP\x02)(\x01\/USERVIP\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USERVIP\x02)(.[^\x01]*)(\x01\/USERVIP\x02)"
			strContent=re.Replace(strContent,"���߼�������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_USERVIP=strContent
end function

function REUBB_USERAGE(strText)
	dim strContent
	dim re,Test
	dim ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[USERAGE=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USERAGE\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USERAGE=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "USERAGE=$1" & chr(2))
			re.Pattern="\[\/USERAGE\]"
			strContent=re.replace(strContent, chr(1) & "/USERAGE" & chr(2))
			re.Pattern="(\x01USERAGE=*([0-9]*)\x02)(\x01\/USERAGE\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USERAGE=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USERAGE\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_USERAGE=strContent
end function

function REUBB_USERGROUP(strText)
	dim strContent
	dim re,Test
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[USERGROUP=(.[^\[]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/USERGROUP\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[USERGROUP=(.[^\[]*)\]"
			strContent=re.replace(strContent, chr(1) & "USERGROUP=$1" & chr(2))
			re.Pattern="\[\/USERGROUP\]"
			strContent=re.replace(strContent, chr(1) & "/USERGROUP" & chr(2))
			re.Pattern="(\x01USERGROUP=(.[^\x01]*)\x02)(\x01\/USERGROUP\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01USERGROUP=(.[^\x01]*)\x02)(.[^\x01]*)(\x01\/USERGROUP\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_USERGROUP=strContent
end function

function REUBB_TIMEL(strText)
	dim strContent
	dim re,Test
	dim ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[TIMEL=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/TIMEL\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[TIMEL=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "TIMEL=$1" & chr(2))
			re.Pattern="\[\/TIMEL\]"
			strContent=re.replace(strContent, chr(1) & "/TIMEL" & chr(2))
			re.Pattern="(\x01TIMEL=*([0-9]*)\x02)(\x01\/TIMEL\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01TIMEL=*([0-9]*)\x02)(.[^\x01]*)(\x01\/TIMEL\x02)"
			strContent=re.Replace(strContent,"����ʾ������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_TIMEL=strContent
end function

function REUBB_TIMEX(strText)
	dim strContent
	dim re,Test
	dim ii
	
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="\[TIMEX=*([0-9]*)\]"
	Test=re.Test(strContent)
	if Test then
		re.Pattern="\[\/TIMEX\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[TIMEX=*([0-9]*)\]"
			strContent=re.replace(strContent, chr(1) & "TIMEX=$1" & chr(2))
			re.Pattern="\[\/TIMEX\]"
			strContent=re.replace(strContent, chr(1) & "/TIMEX" & chr(2))
			re.Pattern="(\x01TIMEX=*([0-9]*)\x02)(\x01\/TIMEX\x02)"
			strContent=re.Replace(strContent,"")
			re.Pattern="(\x01TIMEX=*([0-9]*)\x02)(.[^\x01]*)(\x01\/TIMEX\x02)"
			strContent=re.Replace(strContent,"������������ֹԤ����")
			re.Pattern="\x02"
			strContent=re.replace(strContent, "]")
			re.Pattern="\x01"
			strContent=re.replace(strContent, "[")
		end if
	end if
	set re=Nothing
	REUBB_TIMEX=strContent
end function

function UBB_ACTION(strText)
	dim strContent
	dim re
	dim AlertFontColor

	AlertFontColor="GRAY"
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�����İ����ҳ�Ǯ��˭Ҫ���ҹ�ˮ��</b></font><img src=images/z_dongzuo/WuLiao.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�������ˣ��ء��������Ҫ���Ҹ���ѽ���ء���</b></font><img src=images/z_dongzuo/FuZe.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�������ˣ�ѽ��ѽ��ѽ��</b></font><img src=images/z_dongzuo/ShengQi.gif border=0 align=middle>")
	re.Pattern="(\/�к�)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��Һã��������Ҵ�ϲ�����ӣ�����ǣ� ����ǣ�</b></font><img src=images/z_dongzuo/ZhaoHu.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��ѽ����ѽ��</b></font><img src=images/z_dongzuo/GuZhang.gif border=0 align=middle>")
	re.Pattern="(\/�ȴ�)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��֪��֪������֪��֪�����ҵȵ�����Ҳл�ˡ���</b></font><img src=images/z_dongzuo/DengDai.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���У��Ҽ�����ԣ�</b></font><img src=images/z_dongzuo/FanDui.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�����죬̫��������</b></font><img src=images/z_dongzuo/ChangGe.gif border=0 align=middle>")
	re.Pattern="(\/��Ҫ)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��Ҫѽ����Ҫ��</b></font><img src=images/z_dongzuo/BuYao.gif border=0 align=middle>")
	re.Pattern="(\/ȥ��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�ٺ٣���ȥ���ɣ�</b></font><img src=images/z_dongzuo/QuSi.gif border=0 align=middle>")
	re.Pattern="(\/��Ц)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>����֮�£���Ȼû���ҵĶ��֡���</b></font><img src=images/z_dongzuo/KuangXiao.gif border=0 align=middle>")
	re.Pattern="(\/ʹ��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�Һ�����ѽ���ء������ء���</b></font><img src=images/z_dongzuo/TongKu.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>С�����и��ˣ�������ڡ�</b></font><img src=images/z_dongzuo/DaoBie.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���һ��������ɡ�</b></font><img src=images/z_dongzuo/TiaoWu.gif border=0 align=middle>")
	re.Pattern="(\/ʧ��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>����û����̫û���ˣ�</b></font><img src=images/z_dongzuo/ShiWang.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�����ĵ���������۾�������²�����˭��</b></font><img src=images/z_dongzuo/LangMan.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���������˼һ᲻����˼�ģ�</b></font><img src=images/z_dongzuo/HaiXiu.gif border=0 align=middle>")
	re.Pattern="(\/�ȿ�)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�㡢�㡢��û�ҿᡭ��</b></font><img src=images/z_dongzuo/BiKu.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�����ˣ������ˣ�����������������</b></font><img src=images/z_dongzuo/JiuMing.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���Ҷ����㻹�������������ɵ��޽��ȣ����������ȣ�</b></font><img src=images/z_dongzuo/ZhaoSi.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>������С�������㣬�����ҵĶ��֣�</b></font><img src=images/z_dongzuo/KuangWang.gif border=0 align=middle>")
	re.Pattern="(\/ȭ��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��λ������ĺ�����Ŷ��������ȭ����</b></font><img src=images/z_dongzuo/QuanJi.gif border=0 align=middle>")
	re.Pattern="(\/�Ҳ�)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�Ҳȣ��Ҳȣ�����¥�ϵģ��ȱ�¥�µģ�</b></font><img src=images/z_dongzuo/WoCai.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>¥�ϵģ���������Ұɣ�</b></font><img src=images/z_dongzuo/RaoMing.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���ߣ����ߣ��������㣡</b></font><img src=images/z_dongzuo/WoTi.gif border=0 align=middle>")
	re.Pattern="(\/ɵЦ)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���������롫�� </b></font><img src=images/z_dongzuo/ShaXiao.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�㵱����˭ѽ������Ϊ��������ѽ��</b></font><img src=images/z_dongzuo/ChouMei.gif border=0 align=middle>")
	re.Pattern="(\/��ˮ)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���������ģ������۸�Ŷ����ˮ�ȣ�</b></font><img src=images/z_dongzuo/GuanShui.gif border=0 align=middle>")
	re.Pattern="(\/��̬)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>����ĺñ�̬Ŷ</b></font><img src=images/z_dongzuo/BianTai.gif border=0 align=middle>")
	re.Pattern="(\/ƴ��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>ʲôʱ��զ��ƴƴ�ư���</b></font><img src=images/z_dongzuo/PinJiu.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>������Ϊ�����£��Űپ�ʮ�Ŷ�õ�塭��</b></font><img src=images/z_dongzuo/ShenQing.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�ۣ��������ҵ�Ż��ѽ���ۡ���</b></font><img src=images/z_dongzuo/EXin.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��ʵ����̫��������</b></font><img src=images/z_dongzuo/GaoXing.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���������</b></font><img src=images/z_dongzuo/HuaiYi.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�������˹��Ͳ������ˡ�</b></font><img src=images/z_dongzuo/LaGou.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�þ�û������ô����Ķ����ˡ�</b></font><img src=images/z_dongzuo/JingDian.gif border=0 align=middle>")
	re.Pattern="(\/ͬ��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���Ƕ��ԣ����˰ɣ�</b></font><img src=images/z_dongzuo/TongYi.gif border=0 align=middle>")
	re.Pattern="(\/KISS)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>С���Ӳ�Ҫ͵����</b></font><img src=images/z_dongzuo/KISS.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>ף�����տ��֣�</b></font><img src=images/z_dongzuo/ShengRi.gif border=0 align=middle>")
	re.Pattern="(\/�ε�)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��Ҫ���ˣ�</b></font><img src=images/z_dongzuo/YunDao.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�������������㡣</b></font><img src=images/z_dongzuo/QiNi.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�ָ������</b></font><img src=images/z_dongzuo/CuoLa.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���ͣ����ͣ�</b></font><img src=images/z_dongzuo/JiaYou.gif border=0 align=middle>")
	re.Pattern="(\/��ϲ)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��ϲ����ϲ��ʲôʱ����Ծư���</b></font><img src=images/z_dongzuo/GongXi.gif border=0 align=middle>")
	re.Pattern="(\/��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>ร�ԭ����ô��ѽ����</b></font><img src=images/z_dongzuo/JianDan.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>ͬ־�������ˣ�</b></font><img src=images/z_dongzuo/GuLi.gif border=0 align=middle>")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>������������</b></font><img src=images/z_dongzuo/GuoJiang.gif border=0 align=middle>")
	re.Pattern="(\/ԭ��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���Ǵ����ԭ����һ�ΰɣ�</b></font><img src=images/z_dongzuo/YuanLiang.gif border=0 align=middle>")
	re.Pattern="(\/�ж�)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�Һøж�Ӵ��</b></font><img src=images/z_dongzuo/GanDong.gif border=0 align=middle>")
	re.Pattern="(\/��л)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>лл��лл��</b></font><img src=images/z_dongzuo/DaoXie.gif border=0 align=middle>")
	re.Pattern="(\/ҡͷ)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>����������ѽ��</b></font><img src=images/z_dongzuo/YaoTou.gif border=0 align=middle>")
	re.Pattern="(\/��ͷ)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>û��û��������ô���£�</b></font><img src=images/z_dongzuo/DianTou.gif border=0 align=middle>")
	re.Pattern="(\/ӵ��)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>���ɣ�С�������������屧��</b></font><img src=images/z_dongzuo/YongBao.gif border=0 align=middle>")
	re.Pattern="(\/�ϣ�)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>�ϣ�������</b></font><img src=images/z_dongzuo/OK.gif border=0 align=middle>")
	re.Pattern="(\/��ӭ)"
	strContent=re.Replace(strContent,"<br><font color="&AlertFontColor&"><b>��ӭ����ӭ�����һ�ӭ����</b></font><img src=images/z_dongzuo/HuanYing.gif border=0 align=middle>")
	set re=Nothing
	UBB_ACTION=strContent
end function

function REUBB_ACTION(strText)
	dim strContent
	dim re
	dim AlertFontColor

	AlertFontColor="GRAY"
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	strContent=strText
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/�к�)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/�ȴ�)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��Ҫ)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ȥ��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��Ц)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ʹ��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ʧ��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/�ȿ�)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ȭ��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/�Ҳ�)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ɵЦ)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��ˮ)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��̬)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ƴ��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ͬ��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/KISS)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/�ε�)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��ϲ)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/����)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ԭ��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/�ж�)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��л)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ҡͷ)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��ͷ)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/ӵ��)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/�ϣ�)"
	strContent=re.Replace(strContent,"")
	re.Pattern="(\/��ӭ)"
	strContent=re.Replace(strContent,"")
	set re=Nothing
	REUBB_ACTION=strContent
end function
%>