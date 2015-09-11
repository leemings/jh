<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if session("myroom")<>"h_拥有者" then
	Response.Write "<script Language=Javascript>alert('提示：你非此房间主人无权设置！');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
%>
<html><title>小屋设置操作</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" align="center">
<tr>
<td align="center"> <p><img src="IMAGES/home1.gif" width="196" height="163"></p></td>
</tr>
<tr>
<td align="center">小屋设置操作</td>
</tr>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT h_拥有者2,h_参观 FROM house WHERE h_拥有者='" & aqjh_name &"'",conn,1,1
if rs.Eof or rs.Bof then
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing
	Response.Write "<script Language=Javascript>alert('提示：你还没有房屋或者你无权设置！');parent.email.style.visibility='hidden';</script>"
	Response.End
end if
peiou=rs("h_拥有者2")
guan=rs("h_参观")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<tr><form name="form" method="post" action="setok.asp" >
<td align="center">
<table width="96%" border="0">
          <tr> 
            <td align="center"> 是否允许配偶进入： 
              <select name="peiou" id="peiou">
                <option value="1" <%if peiou<>"无" then%>selected<%end if%>>允许</option>
                <option value="0" <%if peiou="无" then%>selected<%end if%>>禁止</option>
              </select> <br>
              是否允许客人参观：              
              <select name="guan" id="guan">
                <option value="1" <%if guan=1 then%>selected<%end if%>>允许</option>
                <option value="0" <%if guan=0 then%>selected<%end if%>>禁止</option>
              </select>
              <br>
              <br>
              <br></td>
          </tr>
          <tr>
            <td align="center">
<input type="submit" name="Submit" value="确定">
            </td>
          </tr>
        </table>
</td>
</form>
</tr>

</table>

