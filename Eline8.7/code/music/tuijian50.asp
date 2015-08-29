<!--#include file="conn.asp"-->
<!--#include file="Star.INC"-->


<!--#include file="Top.asp"-->
<table width="766" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#E7F3FF">
  <tr> 
    <td width="587" height="20">　<a href="index.asp"><font color="#000000">一线视听&gt;</font></a><font color="#000000"> 
      &gt; <b>推荐单曲</b></font></td>
    <td width="179">&nbsp;</td>
  </tr>
  <tr > 
    <td colspan="2" background="map/fix.gif"><img src="map/fix.gif" width="3" height="1"></td>
  </tr>
</table>
<div align="center"> 
  <table width="766" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#FFFFFF">
    <tr> 
      <td width="129" valign="top" bgcolor="#E7F3FF"> 
        <!--#include file="leftmenu.asp"-->
      </td>
      <td width="1" bgcolor="#FF0000" background="map/fix2.gif"></td>
      <td width="750">
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
          <tr bgcolor="#F7FBFF"> 
            <td width="75%" height="15"><b><font color="#009900">推荐单曲</font></b></td>
            <td width="25%" height="2">&nbsp;</td>
          </tr>
          <tr bgcolor="#F7FBFF"> 
            <td colspan="2" height="15">
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td width="100%" colspan="2"><img border="0" src="images/xiaod.gif" width="1" height="1"></td>
          </tr>
          <tr > 
            <td width="100%" colspan="2"background="map/fix.gif"> 
              <p><img border="0" src="images/xiaod.gif" width="1" height="1"></p>
            </td>
          </tr>
          <tr> 
            <td width="100%" height="0" colspan="2"></td>
          </tr>
          <tr> 
            <td width="100%" colspan="2"> 
              <div align="center"> 
                <center>
                </center>
              </div>
              <div align="center"> 
                <center>
                  <table cellpadding="2" cellspacing="1" width="95%" bordercolordark="#FFFFFF" bgcolor="#000000">
                    <%
set rs=server.createobject("adodb.recordset")

set rs=conn.execute("SELECT top 50 id,Wma,Singer,ClassID,SClassID,NClassID,MusicName,hits FROM MusicList where IsGood=1  order by id desc") 

if not Rs.eof then


%>
                    <form name="form" onsubmit="javascript:return lbsong();" target="_blank" action="Yxplaylist.asp">
                      <tr> 
                        <td width="7%" height=22 align=center bgcolor="#B4DEF8">排名</td>
                        <td width="38%" height=22 align=center bgcolor="#B4DEF8">歌曲名字</td>
                        <td width="11%" height=22 align=center bgcolor="#B4DEF8">人气</td>
                        <td width="11%" height=22 align=center bgcolor="#B4DEF8">Wma试听</td>
                        <td width="10%" height=22 align=center bgcolor="#B4DEF8">点歌</td>
                        <td width="11%" height=22 align=center bgcolor="#B4DEF8">音乐盒</td>
                      </tr>
                      <%
i=0
do while not rs.eof
i=i+1
%>
                      <tr bgcolor="#F7FBFF"> 
                        <td align="center" width="7%"><%=i%></td>
                        <td width="38%"> 　<%=i%>.<a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','Ting88','scrollbars=no,resizable=no,width=300,height=300,menubar=no,top=168,left=168')" ><%=rs("MusicName")%></a></td>
                        <td align=center width="11%"><%=rs("hits")%></td>
                        <td align=center width="11%"> <%if rs("Wma")<>"" then%> <a href="#" onClick="window.open('YxPlay.asp?id=<%=rs("id")%>','Ting88','scrollbars=no,resizable=no,width=300,height=300,menubar=no,top=168,left=168')" >试听</a> 
                          <%else%>
                          暂无 
                          <%end if%> </td>
                        <td align=center width="10%"><a href="#" onClick="window.open('Mailsong.asp?id=<%=rs("id")%>','','scrollbars=no,resizable=no,width=419,height=151,menubar=no,top=98,left=198')">送友</a></td>
                        <td align=center width="11%"><a href="Box.asp?action=add&id=<%=rs("id")%>" target="_blank" title="会员才能使用音乐盒">放入</a></td>
                      </tr>
                      <% 
if m>=50 then exit do
rs.movenext 
loop
else
response.write "尚无收录"
end if
rs.close 
%>
                      <tr bgcolor="#FFFFFF"> 
                        <td align="center" colspan="6" height="22">&nbsp; </td>
                      </tr>
                    </form>
                  </table>
                </center>
              </div>
            </td>
          </tr>
          <tr> 
            <td width="100%" height="10" colspan="2"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</div>
<div align="center"> </div>
<%
set rs=nothing
conn.close
set conn=nothing
%>
<!--#include file="Bottom.asp"-->
