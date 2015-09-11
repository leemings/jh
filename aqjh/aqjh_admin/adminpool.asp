<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
if aqjh_name<>Application("aqjh_user") then Response.Redirect "../chat/readonly/bomb.htm"
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
<link rel="stylesheet" href="setup.css">
<title>修改房间资料</title></head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font color="#000000" size="2">江湖房间修改程序！</font> </div>
<form method="post" action="modiroomok.asp"><table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="75%">
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">ID</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> <%=rs("id")%> </font>
          <input type="hidden" name="id" value="<%=rs("id")%>" size="15" maxlength="20">
      </td>
      <td width="81"> 
        <div align="center"><font size="2" color="#FFFFFF">房间名</font></div>
      </td>
      <td width="224"><font color="#FFFFFF" size="2"> 
        <input type="text" name="name"
value="<%=rs("a")%>" size="15" maxlength="20">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">最大人数</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <input type="text" name="zdrs"
value="<%=rs("b")%>" size="15" maxlength="3">
        </font></td>
      <td width="81" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">发招限制</font></div>
      </td>
      <td width="224"><font color="#FFFFFF" size="2"> 
        <input type="text" name="fzxz"
value="<%=rs("f")%>" size="15" maxlength="1">
        </font></td>
    </tr>
    <tr> 
      <td width="67" nowrap> 
        <div align="center"><font color="#FFFFFF" size="2">事件限制</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <input type="text" name="sjxz"
value="<%=rs("g")%>" size="15" maxlength="1">
        </font></td>
      <td width="81"> 
        <div align="center"><font color="#FFFFFF" size="2">保护</font></div>
      </td>
      <td width="224"><font color="#FFFFFF" size="2"> 
        <input type="text" name="bh"
value="<%=rs("h")%>" size="15" maxlength="1">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">卡片</font></div>
      </td>
      <td width="197"><font color="#FFFFFF" size="2"> 
        <input type="text" name="kp"
value="<%=rs("i")%>" size="15" maxlength="1">
        </font></td>
      <td width="81"> 
        <div align="center"><font size="2" color="#FFFFFF">赌博</font></div>
      </td>
      <td width="224"><font color="#FFFFFF" size="2"> 
        <input type="text" name="db"
value="<%=rs("j")%>" size="15" maxlength="1">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">限制</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <input type="text" name="xz"
value="<%=rs("c")%>" size="5" maxlength="1">
        限制选择为0时允许，为1时禁止！请大家根据选择</font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font color="#FFFFFF" size="2">限制说明</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <input type="text" name="xzsm"
value="<%=rs("d")%>" size="50">
        </font></td>
    </tr>
    <tr> 
      <td width="67"> 
        <div align="center"><font size="2" color="#FFFFFF">表达式</font></div>
      </td>
      <td colspan="3"><font color="#FFFFFF" size="2"> 
        <input type="text" name="bds"
value="<%=rs("e")%>" size="50">
        <br>
        可以使用：and、or、not、+、-、=等表达式请对应条件连接！<br>
        对于编程不了解者不可随意修改！(ID为1主聊天不要加限制！) </font></td>
    </tr>
    <tr> 
      <td colspan="4"> 
        <div align="center"> 
          <input type="submit" value="确 定">
          <font color="#CCCCCC">------- </font> 
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
<div align="center"> <font color="#FFFFFF" size="2">数据库更新程序因为时间有限，没有加入为空值时的检测，请大家更改时注意没有值的地方填0</font></div>