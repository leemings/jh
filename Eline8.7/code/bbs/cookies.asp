<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
response.buffer=true
select case request("action")
case "hidden"
	call hidden(1)
case "hidden3"
	call hidden(3)
case "hidden4"
	call hidden(4)
case "hidden5"
	call hidden(5)
case "hidden6"
	call hidden(6)
case "hidden7"
	call hidden(7)
case "hidden8"
	call hidden(8)
case "online"
	call online()
case "stylemod"
	call stylemod()
case "show0"
	call showboardlist()
case "show1"
	call hideboardlist()
case "mylist"
	call mylist()
case "resetre"
	call resetre()
case "hasview"
	call hasview()
case else
	Errmsg=Errmsg+"<br>"+"<li>请指定正确的参数。"
	founderr=true
end select

call nav()
call head_var(2,0,"","")
if founderr then
	call dvbbs_error()
else
	if isnull(Request.ServerVariables("HTTP_REFERER")) or Request.ServerVariables("HTTP_REFERER")="" then
	response.redirect "index.asp"
	else
	response.redirect Request.ServerVariables("HTTP_REFERER")
	end if
end if
call footer()

sub hidden(curstatus)
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>请登录后进行操作。"
	founderr=true
	exit sub
end if

conn.execute("update online set userhidden="&curstatus&" where userid="&userid)
dim usercookies
usercookies=request.cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies="0"
select case usercookies
case "0"
case 1
 	Response.Cookies("aspsky").Expires=Date+1
case 2
	Response.Cookies("aspsky").Expires=Date+31
case 3
	Response.Cookies("aspsky").Expires=Date+365
end select
Response.Cookies("aspsky")("usercookies") = usercookies
Response.Cookies("aspsky")("userhidden") = curstatus
Response.Cookies("aspsky").path=cookiepath
end sub

sub online()
if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>请登录后进行操作。"
	founderr=true
	exit sub
end if

conn.execute("update online set userhidden=2 where userid="&userid)
dim usercookies
usercookies=request.cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies="0"
select case usercookies
case "0"
case 1
  Response.Cookies("aspsky").Expires=Date+1
case 2
	Response.Cookies("aspsky").Expires=Date+31
case 3
	Response.Cookies("aspsky").Expires=Date+365
end select
Response.Cookies("aspsky")("usercookies") = usercookies
Response.Cookies("aspsky")("userhidden") = 2
Response.Cookies("aspsky").path=cookiepath
end sub

sub stylemod()
dim usercookies
usercookies=request.cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies="0"
select case usercookies
case "0"
	Response.Cookies("skin").Expires=Date+7
case 1
	Response.Cookies("skin").Expires=Date+1
case 2
	Response.Cookies("skin").Expires=Date+31
case 3
	Response.Cookies("skin").Expires=Date+365
end select
if not isnumeric(request("skinid")) then
	Errmsg=Errmsg+"<br>"+"<li>非法的参数。"
	founderr=true
	exit sub
end if
if Cint(request("skinid"))>0 then
Response.Cookies("skin")("skinid_"&boardid)=request("skinid")
else
Response.Cookies("skin")("skinid_"&boardid)=""
end if
end sub

sub showboardlist()
	dim usercookies
	usercookies=request.cookies("aspsky")("usercookies")
	if isnull(usercookies) or usercookies="" then usercookies="0"
	select case usercookies
	case "0"
		Response.Cookies("bbslist").Expires=Date+7
	case 1
		Response.Cookies("bbslist").Expires=Date+1
	case 2
		Response.Cookies("bbslist").Expires=Date+31
	case 3
		Response.Cookies("bbslist").Expires=Date+365
	end select
	Response.Cookies("bbslist")("Isshow")="1"
	response.redirect "index.asp"
end sub

sub hideboardlist()
	dim usercookies
	usercookies=request.cookies("aspsky")("usercookies")
	if isnull(usercookies) or usercookies="" then usercookies="0"
	select case usercookies
	case "0"
		Response.Cookies("bbslist").Expires=Date+7
	case 1
		Response.Cookies("bbslist").Expires=Date+1
	case 2
		Response.Cookies("bbslist").Expires=Date+31
	case 3
		Response.Cookies("bbslist").Expires=Date+365
	end select
	Response.Cookies("bbslist")("Isshow")="0"
	response.redirect "index.asp"
end sub

sub mylist()
	dim usercookies
	usercookies=request.cookies("aspsky")("usercookies")
	if isnull(usercookies) or usercookies="" then usercookies="0"
	select case usercookies
	case "0"
		Response.Cookies("bbslist").Expires=Date+7
	case 1
		Response.Cookies("bbslist").Expires=Date+1
	case 2
		Response.Cookies("bbslist").Expires=Date+31
	case 3
		Response.Cookies("bbslist").Expires=Date+365
	end select
	Response.Cookies("bbslist")("mylist")=replace(replace(request("mylist"),"'","")," ","")
	Response.Cookies("bbslist")("Isshow")=replace(request("IsShow"),"'","")
	response.redirect "index.asp"
end sub

sub hasview()
	dim usercookies
	usercookies=request.cookies("aspsky")("usercookies")
	if isnull(usercookies) or usercookies="" then usercookies="0"
	if Boardid=0 then
		Response.Cookies("LastView").Expires=Date+7
		response.cookies("LastView")=""
		conn.execute("update [user] set LastViewTime=now() where userid="&userid)
	else
		Response.Cookies("LastView").Expires=Date+7
		dim key,tempsplit
		for each key in Request.Cookies("LastView")
			if Key<>"" then
				TempSplit=Split(Key,"_")
				if TempSplit(0)="Topic" then
					if CLng(TempSplit(2))=BoardID then
						response.cookies("LastView")(key)=""
					end if
				end if
			end if
		next
		Response.Cookies("LastView")("Board_"&BoardID)=now()
		response.redirect "list.asp?boardid="&BoardID
	end if
end sub

sub resetre()
	conn.execute("update [user] set reann='' where userid="&userid)
	response.redirect "index.asp"
end sub
%>
