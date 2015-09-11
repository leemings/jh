<!--#include file="../showchat.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
if Application("jhshowurl")<>"http://qqshow-item.tencent.com" then
	showgif=Application("jhshowurl") & "gif/pic/[img].gif"
	showgif1="src="& Application("jhshowurl") & "gif/'+j+'/'+showArray[i]+'.gif"
else
	showgif=Application("jhshowurl") & "/[img]/0/00/00.gif"
	showgif1="src="& Application("jhshowurl") &"/'+showArray[i]+'/'+j+'/00/00.gif"
end if
aqjh_name=Session("aqjh_name")
ndate=Day(date())
if ndate>20 then
		Response.Write "<script Language=Javascript>alert('提示：江湖秀比赛报名只限每月前20天，现在时间已过请下月再参加！');window.close();</script>"
		response.end
end if
myok=trim(lcase(cstr(server.HTMLEncode(Request("ok")))))
%>
<html><title>江湖秀大赛报名</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0">
<%if myok="ok" then
	sm=trim(lcase(cstr(server.HTMLEncode(Request.form("sm")))))
	sm=replace(sm,"&#","")
	sm=Replace(sm,"&amp;","&")
	sm=Replace(sm,"&amp;","&")
	sm=Replace(sm,"'","&#" & asc("'"))
	sm=Replace(sm,"\","")
	sm=Replace(sm,",","&#" & asc(","))
	sm=Replace(sm,"""","&#" & asc(""""))
	sm=Replace(sm,"<","&#" & asc("<"))
	sm=Replace(sm,">","&#" & asc(">"))
	sm=replace(sm,"..","")
	sm=Replace(sm,"\x3c","")
	sm=Replace(sm,"\x3e","")
	sm=Replace(sm,"\074","")
	sm=Replace(sm,"\74","")
	sm=Replace(sm,"\75","")
	sm=Replace(sm,"\76","")
	sm=Replace(sm,"&lt","")
	sm=Replace(sm,"&gt","")
	sm=Replace(sm,"\076","")
	if len(sm)<3 then
		Response.Write "<script Language=Javascript>alert('提示：说明不可以小于3个汉字！');location.href = 'javascript:history.go(-1)';</script>"
		response.end
	end if
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("aqjh_usermdb")
	rs.open "select 金币 from [用户] where 姓名='" & Session("aqjh_name") & "'",conn,1,3
	if rs("金币")<10 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	 	Response.Write "<script language=JavaScript>{alert('提示：你没有10个金币，不能进行报名操作！');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
	end if
	rs("金币")=rs("金币")-10
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("showmdb")
	rs.Open "select * from use where a='"& Session("aqjh_name") &"'", conn, 1,3
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你还没有注册江湖秀，不能够参加比赛！');window.close();</script>"
		response.end
	end if
	if rs("f")="||||||14|13|12||11||10|9||||8|||||||" or rs("f")="||||||7|6|5||4||3|2||||1|||||||" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：您的秀图为初始化图像，请您更改装饰后再来！');window.close();</script>"
		response.end
	end if
	myshow=rs("f")
	if rs("c")="M" then
		sex="男"
	else
		sex="女"
	end if
	rs.close
	newbao=0
	rs.Open "select * from sai where s_姓名='"& Session("aqjh_name") &"'", conn, 1,3
	if rs.eof and rs.bof then
		newbao=1
		rs.AddNew
	end if
	rs("s_姓名")=Session("aqjh_name")
	rs("s_数据")=myshow
	rs("s_票数")=0
	rs("s_加入时间")=now()
	rs("s_说明")=sm
	rs("s_性别")=sex
	rs.update
	if newbao=1 then
		conn.execute "update pool set p_金币=p_金币+10,p_报名人数=p_报名人数+1,p_男第一名='',p_男第二名='',p_女第一名='',p_女第二名='',p_男第一领奖=false,p_男第二领奖=false,p_女第一领奖=false,p_女第二领奖=false"
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	says="<font color=#ff0000>【消息】</font>["&session("aqjh_name")&"]<font color=green>参加了江湖秀选美大赛，此大赛将在21日进行公开投票，距投票还有<font color=blue><b>"& (21-ndate) &"</b></font>天！<font class=t>(" & time() & ")</font></font>"
	call showchat(says)
	Response.Write "<script Language=Javascript>alert('提示：报名成功，请你回到江湖秀页面刷新即可看到你的报名资料！');window.close();</script>"
	response.end
else%>
<table width="100%" height="100%" border="0" align="center" bgcolor="#FF9900">
     <form name="form" method="post" action="biaoming.asp?ok=ok">
     <tr> 
      <td height="83" align="center"> 
        <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("showmdb")
	rs.Open "select * from use where a='"& Session("aqjh_name") &"'", conn, 1,3
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：你还没有注册江湖秀，不能够参加比赛！');window.close();</script>"
		response.end
	end if
	if rs("f")="||||||14|13|12||11||10|9||||8|||||||" or rs("f")="||||||7|6|5||4||3|2||||1|||||||" then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：您的秀图为初始化图像，请您更改装饰后再来！');window.close();</script>"
		response.end
	end if
	Response.Write "<script language=JavaScript>var jhshow='"& rs("f") &"';"
	Response.Write "document.write ("& chr(34) &"<div id=show"& myrnd &" style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; LEFT: 0px; PADDING-BOTTOM: 0px; WIDTH: 140px; PADDING-TOP: 0px; POSITION: relative; TOP: 0px; HEIGHT: 226px;'></div>"& chr(34) &");"
	Response.Write "var showArray = jhshow.split('|');"
	Response.Write "var s='';"
	Response.Write "for (var i=0; i<25; i++){if(showArray[i] != ''){j = i + 1;"
	Response.Write "s+='<IMG id=s'+j+' "& showgif1 &" style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:'+j+';"& chr(34) &">';"
	Response.Write "}}"
	Response.Write "s+='<IMG src=face/blank.gif style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:50;"& chr(34) &">';"
	Response.Write "show"& myrnd &".innerHTML=s;</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
      </td>
  </tr>
  <tr> 
      <td><font color="#FFFF00">确定后将使用当前形像进行报名，报名后系统将扣除您10个金币作为报名费用，一旦报名后不能撤消，如果您要修改形像在21号前你可以重新报名，重新报名后将使用新形像，但系统仍然要再扣除10个金币，所以请您慎重考虑。</font></td>
  </tr>
  <tr>
    <td align="center">形像说明(40汉字)： 
      <input name="sm" type="text" size="20" maxlength="40"><br><br> 
      <input type="submit" name="Submit" value="确定我要报名">
    </td>
  </tr></form>
</table>
<%end if%>
</body></html>