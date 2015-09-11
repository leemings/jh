<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="roomconfig.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=210"
if sfkf=0 then
	response.write "门派房间创建功能已关闭。"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,身份,grade from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
mymp=rs("门派")
mysf=rs("身份")
rs.close
if mysf<>"掌门" or aqjh_grade<5 or mymp="天网" or mymp="出家" or mymp="游侠" then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "你不是掌门，不要乱跑！"
	response.end
end if
rs.open "select count(姓名) as 数量 from 用户 where 门派='"&mymp&"' and 等级>="&dzdj,conn
zs=rs("数量")
rs.close
rs.open "select * from r where a='"&mymp&"'",conn,1,2
if (rs.eof or rs.bof) and zs<dzrs then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "你门派弟子中"&dzdj&"级以上者不足"&dzrs&"人，不能创建房间！"
	response.end
end if
if not(rs.eof or rs.bof) then
	rooma=rs("a")	'房间名'
	roomid=rs("id")		'房间ID
	roomb=rs("b")		'最高在线多少人
	roomc=rs("c")		'是否有限制，0为无限制，1为有限制
	roomd=rs("d")		'进入房间条件说明
	roome=rs("e")		'进入房间条件
	roomf=rs("f")		'房间内是否允许打架，0为允许，1为禁止
	roomg=rs("g")		'房间内是否允许产生随机事件，0为允许，1为禁止
	roomh=rs("h")		'房间内是否允许保护，0为允许，1为禁止
	roomi=rs("i")		'房间内是否允许使用卡片，0为允许，1为禁止
	roomj=rs("j")		'房间内是否允许赌博，0为允许，1为禁止
	tjasp="modiroomok.asp"
	an="修 改"
	s="你是"&mymp&" "&mysf&"，你派弟子中<font color=blue>"&dzdj&"</font>级以上的人共有"&zs&"人"
	if zs<100 then
    		s1="，<font color=red><b不足"&dzrs&"人</b></font>，如在月底时本门"&dzdj&"级以上的人数不足"&dzrs&"以上，站长会删除此门派。"
	else
		s1="。"
    	end if
	s=s&s1
else
	rooma=mymp	'房间名'
	roomid="创建本派房间"		'房间ID
	roomb=200		'最高在线多少人
	roomc=0		'是否有限制，0为无限制，1为有限制
	roomd="允许任何人进入"		'进入房间条件说明
	roome=""		'进入房间条件
	roomf=0		'房间内是否允许打架，0为允许，1为禁止
	roomg=0		'房间内是否允许产生随机事件，0为允许，1为禁止
	roomh=0		'房间内是否允许保护，0为允许，1为禁止
	roomi=0		'房间内是否允许使用卡片，0为允许，1为禁止
	roomj=0		'房间内是否允许赌博，0为允许，1为禁止
	tjasp="addroomok.asp"
	an="创 建"
	s="你是<font color=blue>"&mymp&"</font> <font color=red>"&mysf&"</font>，你门派弟子中"&dzdj&"级以上的已达到"&dzrs&"人以上，现在你可以创建一个本派的房间。"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>门派房间创建管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body bgcolor="#FFFFFF" text="#FFFFFF" background="../../JHIMG/bk_Hc3w.gif">
<div align="center">
  <p><font size="5"><b><font color="#0000FF">江湖门派聚义厅</font></b></font></p>
  <p><font size="2" color="#000000"><%=s%></font></p>
</div>
<form name="form1" method="post" action="<%=tjasp%>">
  <table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="82%">
    <tr> 
      <td width="102"> 
        <div align="center"><font color="#FFFFFF" size="2">ID</font></div>
      </td>
      <td width="75"><font color="#FFFFFF" size="2"> <%=roomid%> </font> 
        <input type="hidden" name="id" value="<%=roomid%>" size="15" maxlength="20">
      </td>
      <td width="81"> <font size="2" color="#FFFFFF">房 间 名</font></td>
      <td width="353"><font color="#FFFFFF" size="2"> 
        <input type="text" name="name" readonly value="<%=rooma%>" size="15" maxlength="20">
        房间名为门派名，不可更改</font> </td>
    </tr>
    <tr> 
      <td width="102"> 
        <div align="center"><font color="#FFFFFF" size="2">是否可以保护</font></div>
      </td>
      <td width="75"><font color="#FFFFFF" size="2"> 
        <select name=bh size=1 >
          <option value='0' <%if roomh=0 then%>selected<%end if%>><font color="#FFFFFF" size="2">允许</font></option>
          <option value='1' <%if roomh=1 then%>selected<%end if%>><font color="#FFFFFF" size="2">禁止</font></option>
        </select>
        </font></td>
      <td width="81" nowrap> <font color="#FFFFFF" size="2">是否可以杀人</font></td>
      <td width="353"><font color="#FFFFFF" size="2"> 
        <select name=fzxz size=1 >
          <option value='0' <%if roomf=0 then%>selected<%end if%>>允许</option>
          <option value='1' <%if roomf=1 then%>selected<%end if%>>禁止</option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="102" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">是否有随机事件</font></div>
      </td>
      <td width="75"><font color="#FFFFFF" size="2"> 
        <select name=sjxz size=1 >
          <option value='0' <%if roomg=0 then%>selected<%end if%>>允许</option>
          <option value='1' <%if roomg=1 then%>selected<%end if%>>禁止</option>
        </select>
        </font></td>
      <td width="81"> <font color="#FFFFFF" size="2">是否可以用卡</font></td>
      <td width="353"><font color="#FFFFFF" size="2"> 
        <select name=kp size=1 >
          <option value='0' <%if roomi=0 then%>selected<%end if%>><font color="#FFFFFF" size="2">允许</font></option>
          <option value='1' <%if roomi=1 then%>selected<%end if%>><font color="#FFFFFF" size="2">禁止</font></option>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td width="102"> 
        <div align="center"><font size="2" color="#FFFFFF">是否可以赌博</font></div>
      </td>
      <td width="75"><font color="#FFFFFF" size="2"> 
        <select name=db size=1 >
          <option value='0' <%if roomj=0 then%>selected<%end if%>><font color="#FFFFFF" size="2">允许</font></option>
          <option value='1' <%if roomj=1 then%>selected<%end if%>><font color="#FFFFFF" size="2">禁止</font></option>
        </select>
        </font></td>
      <td width="81"> 
        <div align="center"><font color="#FFFFFF" size="2">是否只有本门弟子才可进入</font></div>
      </td>
      <td width="353"><font color="#FFFFFF" size="2"> 
        <select name=xz size=1 >
          <option value='0' <%if roomc=0 then%>selected<%end if%>><font color="#FFFFFF" size="2">允许所有人进入</font></option>
          <option value='1' <%if roomc=1 then%>selected<%end if%>><font color="#FFFFFF" size="2">只允许本门进入</font></option>
        </select>
        此限制对官府人员无效 </font></td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center">创建门派聚义厅需要花费银两：<font color="#FF0000"><%=yinliang%></font>、金币：<font color="#FF0000"><%=jinbi%></font>、会员金卡：<font color="#FF0000"><%=jinka%></font><br>
          <input type="submit" value="<%=an%>" name="submit">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.close()" value="关 闭" type=button name="close">
        </div>
      </td>
    </tr>
  </table>
</form>
</body>
</html>