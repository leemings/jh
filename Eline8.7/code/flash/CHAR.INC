<%
function htmlencode(str)
    dim result
    dim l
    if isNULL(str) then 
       htmlencode=""
       exit function
    end if
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case "<"
	                result=result+"&lt;"
	           case ">"
	                result=result+"&gt;"
              case chr(13)
	                result=result+"<br>"
	           case chr(34)
	                result=result+"&quot;"
	           case "&"
	                result=result+"&amp;"
              case chr(32)	           
	                'result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   end if
	                else
	                   result=result+"&nbsp;"	                    
	                end if
	           case chr(9)
	                result=result+"    "
	           case else
	                result=result+mid(str,i,1)
         end select
       next 
       htmlencode=result
   end function


function sustainhtml(str)
    dim result
    dim l
    if isNULL(str) then 
       sustainhtml=""
       exit function
    end if
    l=len(str)
    result=""
	dim i
	for i = 1 to l
	    select case mid(str,i,1)
	           case chr(13)
	                result=result+"<br>"
	           case chr(34)
	                result=result+"&quot;"
                   case chr(32)	           
	                'result=result+"&nbsp;"
	                if i+1<=l and i-1>0 then
	                   if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then	                      
	                      result=result+"&nbsp;"
	                   else
	                      result=result+" "
	                   end if
	                else
	                   result=result+"&nbsp;"	                    
	                end if
	           case "&"
	                result=result+"&amp;"
	           case chr(9)
	                result=result+"    "
	           case else
	                result=result+mid(str,i,1)
         end select
       next 
       sustainhtml=result
   end function

 ' 检查sql字符串中是否有单引号，有则进行转化
   function CheckStr(str)
       dim tstr,l,i,ch
	   l=len(str)
	   for i=1 to l
	       ch=mid(str,i,1)
	       if ch="'" then
		      tstr=tstr+"'"
		   end if
		   tstr=tstr+ch
	   next
	   CheckStr=tstr
   end function

'email有效性检查
function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if

end function

'将一文字中的HTML的标识去掉
Function FilterHTML(strToFilter)
Dim strTemp
strTemp = strToFilter
While Instr(1,strTemp,"<") AND Instr(1, strTemp, ">")
strTemp = Left(strTemp, Instr(1, strTemp, "<")-1) & Right(strTemp, Len(strTemp)-Instr(1,strTemp, ">"))
WEnd
FilterHTML = strTemp
End Function

'UBB代码
function UBBCode(strContent)
	on error resume next
	strContent = HTMLEncode(strContent)
	dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =true
	objRegExp.Global=True

	objRegExp.Pattern="(\[URL\])(.*)(\[\/URL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")

	objRegExp.Pattern="(\[URL=(.*)\])(.*)(\[\/URL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")

	objRegExp.Pattern="(\[EMAIL\])(.*)(\[\/EMAIL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""mailto:$2"">$2</A>")
	objRegExp.Pattern="(\[EMAIL=(.*)\])(.*)(\[\/EMAIL\])"
	strContent= objRegExp.Replace(strContent,"<A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

	if strflash= "1" then
	objRegExp.Pattern="(\[FLASH\])(.*)(\[\/FLASH\])"
	strContent= objRegExp.Replace(strContent,"<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")
	end if

	objRegExp.Pattern="(\[IMG\])(.*)(\[\/IMG\])"
	strContent=objRegExp.Replace(strContent,"<IMG SRC=""$2"" border=1>")

        objRegExp.Pattern="(\[HTML\])(.*)(\[\/HTML\])"
	strContent=objRegExp.Replace(strContent,"<SPAN><IMG src=pic/code.gif align=absBottom> HTML 代码片段如下:<BR><TEXTAREA style=""WIDTH: 94%; BACKGROUND-COLOR: #f7f7f7"" name=textfield rows=10>$2</TEXTAREA><BR><INPUT onclick=runEx() type=button value=运行此代码 name=Button> [Ctrl+A 全部选择   提示:你可先修改部分代码，再按运行]</SPAN><BR>")

	objRegExp.Pattern="(\[color=(.*)\])(.*)(\[\/color\])"
	strContent=objRegExp.Replace(strContent,"<font color=$2>$3</font>")
	objRegExp.Pattern="(\[face=(.*)\])(.*)(\[\/face\])"
	strContent=objRegExp.Replace(strContent,"<font face=$2>$3</font>")
	objRegExp.Pattern="(\[align=(.*)\])(.*)(\[\/align\])"
	strContent=objRegExp.Replace(strContent,"<div align=$2>$3</div>")

	objRegExp.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
	strContent=objRegExp.Replace(strContent,"<BLOCKQUOTE><font size=1 face=""Verdana, Arial"">quote:</font><HR>$2<HR></BLOCKQUOTE>")
	objRegExp.Pattern="(\[fly\])(.*)(\[\/fly\])"
	strContent=objRegExp.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$2</marquee>")
	objRegExp.Pattern="(\[move\])(.*)(\[\/move\])"
	strContent=objRegExp.Replace(strContent,"<MARQUEE scrollamount=3>$2</marquee>")
	objRegExp.Pattern="(\[glow=(.*),(.*),(.*)\])(.*)(\[\/glow\])"
	strContent=objRegExp.Replace(strContent,"<table width=$2 style=""filter:glow(color=$3, strength=$4)"">$5</table>")
	objRegExp.Pattern="(\[SHADOW=(.*),(.*),(.*)\])(.*)(\[\/SHADOW\])"
	strContent=objRegExp.Replace(strContent,"<table width=$2 style=""filter:shadow(color=$3, direction=$4)"">$5</table>")
    
	objRegExp.Pattern="(\[i\])(.*)(\[\/i\])"
	strContent=objRegExp.Replace(strContent,"<i>$2</i>")
	objRegExp.Pattern="(\[u\])(.*)(\[\/u\])"
	strContent=objRegExp.Replace(strContent,"<u>$2</u>")
	objRegExp.Pattern="(\[b\])(.*)(\[\/b\])"
	strContent=objRegExp.Replace(strContent,"<b>$2</b>")
	objRegExp.Pattern="(\[fly\])(.*)(\[\/fly\])"
	strContent=objRegExp.Replace(strContent,"<marquee>$2</marquee>")

	objRegExp.Pattern="(\[size=1\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=1>$2</font>")
	objRegExp.Pattern="(\[size=2\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=2>$2</font>")
	objRegExp.Pattern="(\[size=3\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=3>$2</font>")
	objRegExp.Pattern="(\[size=4\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=4>$2</font>")

	strContent = doCode(strContent, "[list]", "[/list]", "<ul>", "</ul>")
	strContent = doCode(strContent, "[list=1]", "[/list]", "<ol type=1>", "</ol id=1>")
	strContent = doCode(strContent, "[list=a]", "[/list]", "<ol type=a>", "</ol id=a>")
	strContent = doCode(strContent, "[*]", "[/*]", "<li>", "</li>")
	strContent = doCode(strContent, "[code]", "[/code]", "<pre id=code><font size=1 face=""Verdana, Arial"" id=code>", "</font id=code></pre id=code>")

	set objRegExp=Nothing
	UBBCode=strContent
end function
'VBB代码
Function vbcheck(message)
            vbcheck = replace(message, "<","&lt",1,-1,1)
            vbcheck = replace(vbcheck, ">","&gt",1,-1,1)
			vbcheck = replace(vbcheck,"'","''",1,-1,1) 
			vbcheck = replace(vbcheck,vbcrlf,"<br>"&vbcrlf)
            vbcheck = replace(vbcheck, "[b]","<b>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/b]","</b>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[s]", "<s>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/s]", "</s>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[u]","<u>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/u]","</u>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[i]","<i>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/i]","</i>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[red]", "<font color=red>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/red]", "</font id=red>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[green]", "<font color=green>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/green]", "</font id=green>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[blue]", "<font color=blue>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/blue]", "</font id=blue>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[white]", "<font color=white>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/white]", "</font id=white>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[purple]", "<font color=purple>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/purple]", "</font id=purple>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[yellow]", "<font color=yellow>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/yellow]", "</font id=yellow>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[violet]", "<font color=violet>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/violet]", "</font id=violet>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[brown]", "<font color=brown>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/brown]", "</font id=brown>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[black]", "<font color=black>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/black]", "</font id=black>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[pink]", "<font color=pink>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/pink]", "</font id=pink>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[orange]", "<font color=orange>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/orange]", "</font id=orange>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[gold]", "<font color=gold>", 1, -1, 1)
		    vbcheck = replace(vbcheck, "[/gold]", "</font id=gold>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[size=1]", "<font size=1>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/size=1]", "</font id=size1>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[size=2]", "<font size=2>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/size=2]", "</font id=size2>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[size=3]", "<font size=3>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/size=3]", "</font id=size3>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[size=4]", "<font size=4>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/size=4]", "</font id=size4>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[size=5]", "<font size=5>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/size=5]", "</font id=size5>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[size=6]", "<font size=6>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/size=6]", "</font id=size6>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[marquee]", "<marquee direction=left>", 1, -1, 1)
			vbcheck = replace(vbcheck, "[/marquee]","</marquee>", 1, -1, 1)
            vbcheck = replace(vbcheck, "[img]","<img src=""",1,-1,1)
			vbcheck = replace(vbcheck, "[/img]",""" border=0>", 1, -1, 1)
			vbcheck=replace(vbcheck,"[url=","<a href=""",1,-1,1)
            vbcheck=replace(vbcheck,"[/url]","</a>",1,-1,1)
			vbcheck=replace(vbcheck,"[email=","<a href=""mailto:",1,-1,1)
            vbcheck=replace(vbcheck,"[/email]","</a>",1,-1,1)
			vbcheck=replace(vbcheck,"]",""" target=_blank>",1,-1,1)
End Function
'vb代码反编译
Function unvbcheck(message)
            unvbcheck = replace(message, "&lt","<",1,-1,1)
            unvbcheck = replace(unvbcheck, "&gt",">",1,-1,1)
			unvbcheck = replace(unvbcheck,"''","'",1,-1,1) 
			unvbcheck = replace(unvbcheck,"<br>",vbcrlf)
            unvbcheck = replace(unvbcheck, "<b>","[b]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</b>","[/b]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<s>", "[s]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</s>", "[/s]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<u>","[u]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</u>","[/u]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<i>","[i]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</i>","[/i]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=red>", "[red]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=red>", "[/red]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=green>", "[green]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=green>", "[/green]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=blue>", "[blue]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=blue>", "[/blue]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=white>", "[white]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=white>", "[/white]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=purple>", "[purple]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=purple>", "[/purple]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=yellow>", "[yellow]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=yellow>", "[/yellow]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=violet>", "[violet]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=violet>", "[/violet]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=brown>", "[brown]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=brown>", "[/brown]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=black>", "[black]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=black>", "[/black]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=pink>", "[pink]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=pink>", "[/pink]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=orange>", "[orange]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=orange>", "[/orange]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "<font color=gold>", "[gold]", 1, -1, 1)
		    unvbcheck = replace(unvbcheck, "</font id=gold>", "[/gold]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<font size=1>", "[size=1]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</font id=size1>", "[/size=1]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<font size=2>", "[size=2]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</font id=size2>", "[/size=2]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<font size=3>", "[size=3]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</font id=size3>", "[/size=3]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<font size=4>", "[size=4]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</font id=size4>", "[/size=4]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<font size=5>", "[size=5]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</font id=size5>", "[/size=5]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<font size=6>", "[size=6]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</font id=size6>", "[/size=6]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "<marquee direction=left>", "[marquee]", 1, -1, 1)
			unvbcheck = replace(unvbcheck, "</marquee>","[/marquee]", 1, -1, 1)
            unvbcheck = replace(unvbcheck, "<img src=""","[img]",1,-1,1)
			unvbcheck = replace(unvbcheck, """ border=0>","[/img]", 1, -1, 1)
			unvbcheck=replace(unvbcheck,"<a href=""","[url=",1,-1,1)
            unvbcheck=replace(unvbcheck,"</a>","[/url]",1,-1,1)
			unvbcheck=replace(unvbcheck,"<a href=""mailto:","[email=",1,-1,1)
            unvbcheck=replace(unvbcheck,"</a>","[/email]",1,-1,1)
			unvbcheck=replace(unvbcheck,""" target=_blank>","]",1,-1,1)		
End Function
%>