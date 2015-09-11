<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
%>
<!--#include file="config.asp"-->
<%
wupinid=clng(Request("wupinid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="select * from r where id="&wupinid
rs.open sqlstr,conn
if rs.EOF or rs.BOF then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：没有你要操作的房间，请刷新！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href=css/css.css type=text/css rel=stylesheet>
<title>修改房间资料</title></head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<div align="center"><font color="#000000">房间修改程序！ </div>
<form method="post" action="modiroomok.asp"><center>
  <table border="0" cellspacing="1" cellpadding="4" bgcolor="f2f2ea" cellspacing=0 cellpadding=0  width="75%">
    <tr> 
      <td width="67"> 
        <div align="center">ID</div>
      </td>
      <td width="197"> <%=rs("id")%>  
        <input type="hidden" name="id" value="<%=rs("id")%>" size="15" maxlength="20">
      </td>
      <td width="81"> 
        <div align="center">房间名</div>
      </td>
      <td width="224"> 
        <input type="text" name="name"
value="<%=rs("a")%>" size="15" maxlength="20" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">最大人数</div>
      </td>
      <td width="197"> 
        <input type="text" name="zdrs"
value="<%=rs("b")%>" size="15" maxlength="3" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
      <td width="81" nowrap> 
        <div align="center">发招限制</div>
      </td>
      <td width="224"> 
        <select name=fzxz size=1 >
          <option value='0' <%if rs("f")=0 then%>selected<%end if%>>允许</option>
          <option value='1' <%if rs("f")=1 then%>selected<%end if%>>禁止</option>
        </select>
        </td>
    </tr>
 <tr> 
      <td width="67"> 
        <div align="center">打架时间</div>
      </td>
      <td width="197"> 
        <input type="text" name="djsj"
value="<%=rs("l")%>" size="2" maxlength="2" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid"> [设置60为不限制]
        </td>
      <td width="81" nowrap> 
        <div align="center">泡点</div>
      </td>
      <td width="224"> 
        <select name=pd size=1 >
          <option value='1' <%if rs("m")=1 then%>selected<%end if%>>允许</option>
          <option value='0' <%if rs("m")=0 then%>selected<%end if%>>禁止</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center">事件限制</div>
      </td>
      <td width="197"> 
        <select name=sjxz size=1 >
          <option value='0' <%if rs("g")=0 then%>selected<%end if%>>允许</option>
          <option value='1' <%if rs("g")=1 then%>selected<%end if%>>禁止</option>
        </select>
        </td>
      <td width="81"> 
        <div align="center">保护</div>
      </td>
      <td width="224"> 
        <select name=bh size=1 >
          <option value='0' <%if rs("h")=0 then%>selected<%end if%>>允许</option>
          <option value='1' <%if rs("h")=1 then%>selected<%end if%>>禁止</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">卡片</div>
      </td>
      <td width="197"> 
        <select name=kp size=1 >
          <option value='0' <%if rs("i")=0 then%>selected<%end if%>>允许</option>
          <option value='1' <%if rs("i")=1 then%>selected<%end if%>>禁止</option>
        </select>
        </td>
      <td width="81"> 
        <div align="center">赌博</div>
      </td>
      <td width="224"> 
        <select name=db size=1 >
          <option value='0' <%if rs("j")=0 then%>selected<%end if%>>允许</option>
          <option value='1' <%if rs("j")=1 then%>selected<%end if%>>禁止</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">限制</div>
      </td>
      <td colspan="3"> 
        <select name=xz size=1 >
          <option value='0' <%if rs("c")=0 then%>selected<%end if%>>允许[加入此房间无限制]</option>
          <option value='1' <%if rs("c")=1 then%>selected<%end if%>>禁止[加入此房间有限制]</option>
        </select>
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">限制说明</div>
      </td>
      <td colspan="3"> 
        <input type="text" name="xzsm"
value="<%=rs("d")%>" size="50" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        </td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center">表达式</div>
      </td>
      <td colspan="3"> 
        <input type="text" name="bds"
value="<%=rs("e")%>" size="50" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        <br>
        可以使用：and、or、not、+、-、=等表达式请对应条件连接！<br>
        对于编程不了解者不可随意修改！(ID为1主聊天不要加限制！) </td>
    </tr>
    <tr> 
      <td width="67" height="2"> 
        <div align="center">今日话题</div>
      </td>
      <td colspan="3" height="2"> 
        <input type="text" name="jrht"
value="<%=rs("k")%>" size="50" maxlength="150" class="form" style="BORDER-BOTTOM:#B8AF86 1px solid">
        (在进入江湖时是会显示的)</td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"> 
          <input type="submit" value="确 定" name="submit">
          <font color="#CCCCCC">-------  
          <input  onClick="javascript:window.document.location.href='roomlist.asp'" value="返 回" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<%
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<div align="center"> 数据库更新程序因为时间有限，没有加入为空值时的检测，请大家更改时注意没有值的地方填0</div>