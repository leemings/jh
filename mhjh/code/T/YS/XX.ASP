<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
un=session("yx8_mhjh_username")
if un="" then Response.Redirect "../../error.asp?id=016"
%><!--#include file="data.asp"--><%
sql="select id,等级,攻击,防御,速度,图,名称,特效 from mj where 所有者='"&un&"' and 类型='轩'"
set rs=connb.execute(sql)
%>
<html>
<head>
<title>要塞武器</title>
<LINK href="../../style.css" rel=stylesheet>
</head>
<body background='../../chatroom/bg1.gif'>
<div align="center">　
</div>
<div align="center"><b><font size="3">要塞武器</font></b>
</div>
<div align="center"><br>
  <a href="add.asp"><b>增加武器</b></a><br>
</div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="570" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
    <td height="25" width="67"> 
      <div align="center">招式</div>
    </td>
    <td height="25" width="81"> 
      <div align="center">攻击</div>
    </td>
    <td height="25" width="87"> 
      <div align="center">防御</div>
    </td>
    <td height="25" width="61"> 
      <div align="center">等级</div>
    </td>
    <td height="25" width="61"> 
      <div align="center">速度</div>
    </td>
        <td height="25" width="61"> 
      <div align="center">特效</div>
    </td>
        <td height="25" width="62"> 
      <div align="center">图片</div>
    </td>
    <td height="25" width="56"> 
      <div align="center"> 操作 </div>
    </td>
  </tr>
  <%
do while not (rs.EOF or rs.BOF)
gj=rs("攻击")
fy=rs("防御")
if gj=-100000 then gj=0
if fy=-100000 then fy=0
%>
    <tr>  
    <td height="2" width="67"> 
      <div align="center"> <font color="#000000"><%=rs("名称")%></font> <font color="#000000"> 
          <input type="hidden" name="id" size="10" value="<%=rs("id")%>" maxlength="10">
          </font></div>
      </td>   
    <td height="2" width="81"> 
      <div align="center"> <font color="#000000"><%=gj%></font></div>
      </td>   
    <td height="2" width="87"> 
      <div align="center"> <font color="#000000"><%=fy%></font></div>
      </td> 
    <td height="2" width="61"> 
      <div align="center"><font color="#000000"><%=rs("等级")%> 
          </font></div>
      </td>     
          <td height="2" width="61"> 
      <div align="center"><font color="#000000"><%=rs("速度")%> 
          </font></div>
      </td>
                <td height="2" width="61"> 
      <div align="center"><font color="#000000"><%=rs("特效")%> 
          </font></div>
      </td>
                <td height="2" width="62"> 
      <div align="center"><font color="#000000"><%if rs("图")<>"随机" then%>
      <img src="gif/<%=rs("图")%>">
      <%else%>随机图片<%end if%> 
          </font></div>
      </td>
      <td height="2" width="56"> 
        <div align="center"> <a href="mod.asp?id=<%=rs("id")%>">修改</a>
        <a href="del.asp?id=<%=rs("id")%>">删除</a>
        </div>
      </td>
    </tr>
  <%
rs.movenext
loop
rs.close
%>
</table>
<p align="center">
要塞武器是称霸江湖、攻打敌人要塞的必备武器，只要够条件就可以创立，但只限三种，武器攻击消耗干净之后，要重新设置武器并重新修炼才能继续使用。
<br>
<%
set rs=nothing
connb.close
set connb=nothing
%>
<br><br><a href="xl.asp"><font color="#800000">修炼要塞武器</font></a><br><a href="../../chatroom/yaobu/longeat.asp"><font color="#800000">购买要塞物资</font></a>

</body>

</html>