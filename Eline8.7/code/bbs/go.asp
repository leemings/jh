<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<%
	response.buffer=true
	dim times
	stats="跳转主题"
	if request("sid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
	elseif not isInteger(request("sid")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>非法的帖子参数。"
	else
		times=request("sid")
	end if
	if founderr then
		call nav()
		call head_var(2,0,"","")
		call dvbbs_error()
	else
		call nav()
		call head_var(1,BoardDepth,0,0)
		call main()
		if founderr then call dvbbs_error()
	end if
	call footer()
	sub main()
	if request("action")="next" then
	set rs=conn.execute("select top 1 topicid from topic where boardid="&boardid&" and topicid>"&times&" and not locktopic=2 order by LastPostTime")
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>没有找到下一篇帖子，请<a href=list.asp?boardid="&boardid&">返回</a>。"
	else
		response.redirect "dispbbs.asp?boardid="&boardid&"&ID="&rs(0)&""
	end if
	else
	set rs=conn.execute("select top 1 topicid from topic where boardid="&boardid&" and  topicid<"&times&" and not locktopic=2 order by LastPostTime desc")
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>没有找到下一篇帖子，请<a href=list.asp?boardid="&boardid&">返回</a>。"
	else
		response.redirect "dispbbs.asp?boardid="&boardid&"&ID="&rs(0)&""
	end if
	end if
	end sub
	set rs=nothing
%>