<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('提示：必须进入聊天室才可以操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
id=request("id")
if not(isnumeric(id)) then
	Response.Write "<script language=JavaScript>{alert('滚！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")

rs.open "select * from fw where id="&id,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('情人小筑中没有这个房间出租！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
yhid=rs("yhid")
fwm=rs("a")
rs.close
if yhid=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('此房还未出租！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.open "select id,姓名,配偶,状态,性别 from 用户 where id="&yhid,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('江湖中找不到此租房人的资料，不能进入！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
fwyhid=rs("id")
fwyhxm=rs("姓名")
fwyhpo=rs("配偶")
fwyhzt=rs("状态")
fwyhsex=rs("性别")
rs.close
if fwyhpo<>aqjh_name then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('此房间不是为你开设的，不能进入！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
randomize timer
r=int(rnd*14)+1
if r=3 or r=6 or r=9 or r=11 or r=12 or r=15 then
	if fwyhsex="女" then
		conn.execute "update 用户 set 保留='孕|'&now()&'|100|正常' where id="&yhid
		t="妻子怀孕了。"
	else
		conn.execute "update 用户 set 保留='孕|'&now()&'|100|正常' where 姓名='"&aqjh_name&"'"
		t="怀孕了。"
	end if
end if
conn.execute "update fw set yhid=0 where id="&id
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>【情人小筑】</font><b><font color=#8000FF>"&aqjh_name&"轻轻走进了"&fwyhxm&"在情人小筑租下的<font color=red>※"&fwm&"※</font>房间……</font></b>"		'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
addmsg saystr
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
	Application.Lock()
	Application("SayCount")=Application("SayCount")+1
	i="SayStr"&YuShu(Application("SayCount"))
	Application(i)=Str
	Application.UnLock()
End Sub
%>
<HTML>
<HEAD>
<TITLE>江湖情人小筑</TITLE>
<META http-equiv=""Content-Type"" content=""text/html"">
</HEAD>

<BODY BGCOLOR="#8800FF">
<div align="center">
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH=512 HEIGHT=482>
<TR>
<TD ROWSPAN=1 COLSPAN=3 WIDTH=512 HEIGHT=77><IMG SRC="images\h11.gif" WIDTH=512 HEIGHT=77 BORDER=0 ALT="h11.gif"></TD>
</TR>
<TR>
<TD ROWSPAN=1 COLSPAN=1 WIDTH=104 HEIGHT=306><IMG SRC="images\h21.gif" WIDTH=104 HEIGHT=306 BORDER=0 ALT="h21.gif"></TD>
      <TD ROWSPAN=1 COLSPAN=1 WIDTH=309 HEIGHT=306 bgcolor="#FFFFFF" background="images/h22.gif"> 
        <div align="center"><MARQUEE onmouseover=this.stop() onmouseout=this.start() scrollAmount=1 scrollDelay=60 direction=up width=150 height=306>
          <font color=blue><script src="xzm.js"></script></font>
          </marquee></div>
      </TD>
<TD ROWSPAN=1 COLSPAN=1 WIDTH=99 HEIGHT=306><IMG SRC="images\h23.gif" WIDTH=99 HEIGHT=306 BORDER=0 ALT="h23.gif"></TD>
</TR>
<TR>
<TD ROWSPAN=1 COLSPAN=3 WIDTH=512 HEIGHT=99><IMG SRC="images\h31.gif" WIDTH=512 HEIGHT=99 BORDER=0 ALT="h31.gif"></TD>
</TR>
</TABLE></div>
</BODY>
</HTML>
