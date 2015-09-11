<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="roomconfig.asp"-->
<!--#include file="../../chk.asp"-->
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
ChkPost()
if sfkf=0 then
	response.write "门派房间创建功能已关闭。"
	response.end
end if
if session("aqjh_inthechat")<>"1" then
	Response.Write "<script language=JavaScript>{alert('你必须在进入聊天室后才可以修改本门聚义厅的设置！！');window.close();}</script>"
	Response.End
end if
function jc(z)
	if not isnumeric(z) or (z<>0 and z<>1) then
		jc=true
	else
		jc=false
	end if
end function
h=clng(request.form("bh"))	'保护h
f=clng(request.form("fzxz"))	'发招f
g=clng(request.form("sjxz"))	'事件g
i=clng(request.form("kp"))	'卡片i
j=clng(request.form("db"))	'赌博j
c=clng(request.form("xz"))	'进入c
if jc(bh) or jc(fzxz) or jc(sjxz) or jc(kp) or jc(db) or jc(xz) then
	Response.Write "<script language=JavaScript>{alert('严重警告，提交数据错误！！');window.close();}</script>"
	Response.End 
end if

'固定值
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,身份,grade from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
mymp=rs("门派")
mysf=rs("身份")
rs.close
if mysf<>"掌门" or aqjh_grade<5 or mymp="天网" or mymp="出家" or mymp="游侠" or mymp="官府" then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "你不是掌门，不要乱跑！"
	response.end
end if
if xz=1 then
	d="只有"&mymp&"门派弟子方可进入"
	k=mymp&"重地，非本派人员擅闯者，杀无赫！！！"
else
	d="任何人都可以进入"
	k=mymp&"聚义厅，广交天下朋友，欢迎各们光临。"
end if
rs.open "select * from r where a='"&mymp&"'",conn,2,2
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你们派的聚义厅还未创建，不能修改！');window.close();}</script>"
	Response.End
end if
if rs("c")=c and rs("f")=f and rs("g")=g and rs("h")=h and rs("i")=i and rs("j")=j then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你没有改变任何设置，不需要修改！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
conn.execute "update r set c="&c&",d='"&d&"',f="&f&",g="&g&",h="&h&",i="&i&",j="&j&",k='"&k&"' where a='"&mymp&"'"
rs.open "select * from r order by id",conn,2,2
application.lock()
roomming=""
do while Not rs.Eof
	rooming=rooming&rs("a")&"|"&rs("b")&"|"&rs("c")&"|"&rs("d")&"|"&rs("e")&"|"&rs("f")&"|"&rs("g")&"|"&rs("h")&"|"&rs("i")&"|"&rs("j")&";"
	rs.MoveNext
loop
application("aqjh_room")=rooming
application.unlock()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../hc3w_Admin/setup.css">
<title>门派聚义大厅创建</title></head>
<body text="#FFFFFF" background="../../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
  <font color="#000000" size="2">
  <p>江湖各门派聚义厅</p>
  </font></div>
<table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="75%">
  <tr> 
    <td width="67"> 
      <div align="center"><font color="#FFFFFF" size=2>房间ＩＤ</font></div>
    </td>
    <td width="197"><font size="2"> <%=id%></font></td>
    <td width="81"> 
      <div align="center"><font size="2" color="#FFFFFF">房 间 名</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2"><%=mymp%></font></td>
  </tr>
  <tr> 
    <td width="67"> 
      <div align="center"><font color="#FFFFFF" size="2">发招限制</font></div>
    </td>
    <td width="197">
      <%if f=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
    </td>
    <td width="81" nowrap> 
      <div align="center"><font color="#FFFFFF" size="2">事件限制</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2">
      <%if g=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
      </font></td>
  </tr>
  <tr> 
    <td width="67" nowrap> 
      <div align="center"><font color="#FFFFFF" size="2">保&nbsp;&nbsp;护</font></div>
    </td>
    <td width="197"><font color="#FFFFFF" size="2">
      <%if h=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
      </font></td>
    <td width="81"> 
      <div align="center"><font color="#FFFFFF" size="2">卡&nbsp;&nbsp;片</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2">
      <%if h=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
      </font></td>
  </tr>
  <tr> 
    <td width="67"> 
      <div align="center"><font size="2" color="#FFFFFF">赌&nbsp;&nbsp;博</font></div>
    </td>
    <td width="197"><font color="#FFFFFF" size="2">
      <%if j=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
      </font></td>
    <td width="81"> 
      <div align="center"><font color="#FFFFFF" size="2">限&nbsp;&nbsp;制</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2"><%=d%></font></td>
  </tr>
  <tr> 
    <td colspan="4">
      <div align="center"> <font size="2"> <br>
        恭喜，<font color=red><b><%=mymp%></b></font>聚义厅设置修改完成。</font><br>
        <br>
        <input  onClick="javascript:location.href = 'javascript:history.go(-1)';" value="返 回" type=button name="back">
        -------- 
        <input  onClick="javascript:window.close()" value="关 闭" type=button name="close">
      </div>
    </td>
  </tr>
</table>