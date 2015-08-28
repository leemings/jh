<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"%>
<HTML>
<HEAD>
<TITLE>魔导寻龙</TITLE>
<link rel=stylesheet href='../../style.css'>
</HEAD>
<BODY topmargin="20" leftmargin="0" rightmargin="0" background='../../chatroom/bg1.gif'>
<CENTER>
<TABLE border="0" width="550" cellpadding="0" cellspacing="0">
  <TBODY>
    <TR>
      <TD></TD>
    </TR>
    <TR>
      <TD>
      <TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
        <TBODY>
          <TR>
            <TD align="center">
            <TABLE border="0" width="100%" cellpadding="0" cellspacing="0" height="2">
              <TBODY>
                <TR>
                  <TD height="32">
                    <p align="center">&nbsp;<B>现有神龙列表</B></p>
                  </TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE width="100%" border="0" cellpadding="0" cellspacing="0">
              <TBODY>
                <TR>
                  <TD></TD>
                </TR>
</CENTER>
                <TR>
                  <TD>
                    <TABLE border="1" bordercolor="#000000" cellpadding="1" cellspacing="1" width="100%">
                      <TBODY> 
                      <TR> 
                        <TD width="13%" height="20" valign="bottom"> 
                          <DIV align="center">
                            <p align="left"><B>神龙名</B></DIV>
                        </TD>
<CENTER>
                        <TD width="37%" height="20" valign="bottom"> 
                          <DIV align="center"><B>功能</B></DIV>
                        </TD>
                        <TD width="26%" height="20" valign="bottom"> 
                          <DIV align="center"><B>售价</B></DIV>
                        </TD>
                        <TD width="11%" height="20" valign="bottom"> 
                          <div align="center"><b>限制</b></div>
                        </TD>
                        <TD width="13%" height="20" valign="bottom"> 
                          <DIV align="center"><B>操作</B></DIV>
                        </TD>
                      </TR>
<!--#include file="data.asp"--><% 
sqlstr="SELECT 物品名,功能,精力,会员,id,银两 FROM 商店 where  类型='龙'"
rs.open sqlstr,connb,1,2
do while not rs.bof and not rs.eof
tai=rs("会员")
if tai="是" then
msg="会员购买"
else
msg="没有限制"
end if
%>

                      <TR> 
                        <TD width="13%" height="20" valign="bottom"> 
                          
                         <%=rs("物品名")%></TD>
                        <TD width="37%" height="20" valign="bottom"> 
                          <div align="left"><%=rs("功能")%></div>
                        </TD>
                    </center>
                        <TD width="26%"  height="20" valign="bottom"> 
                          <DIV align="center">
                            <p align="left">银两<%=rs("银两")%>,精力<%=rs("精力")%></DIV>
                        </TD>
<CENTER>
                        <TD width="11%" height="20" valign="bottom"> 
                          <div align="center"><%=msg%></div>
                        </TD>
                        <TD width="13%" height="20" valign="bottom"> 
                          <DIV align="center"><a href="long1.asp?id=<%=rs("id")%>">购买</a></DIV>
                        </TD>
                        <%
rs.movenext
loop
%>
                      </TR>
                      </TBODY> 
                    </TABLE>
                    </center>
             </TD>
                </TR>
                <TR>
                  <TD align="center" height="20" valign="bottom"> 
                    <br>
                    <p><A class="w" href="javascript:window.close()"><B><font color="#000000">[关闭窗口]</font></B></A><br>
                    </p>
                  </TD>
                </TR>
              </TBODY>
            </TABLE>
            </TD>
          </TR>
          <TR>
            <TD >　</TD>
          </TR>
        </TBODY>
      </TABLE>
      </TD>
    </TR>
  </TBODY>
</TABLE>
</BODY>
</HTML>
<%
rs.close
set rs=nothing
connb.close
set connb=nothing
%>

