<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"%>
<HTML>
<HEAD>
<TITLE>要塞补给</TITLE>
<link rel=stylesheet href='../../style.css'>
</HEAD>
<BODY topmargin="20" leftmargin="0" rightmargin="0" background='../bg1.gif'> 
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
                    <p align="center"><b>要塞补给</b></p>
                  </TD>
                </TR>
              </TBODY>
            </TABLE>
            <TABLE width="100%" border="0" cellpadding="0" cellspacing="0">
              <TBODY>
                <TR>
                  <TD></TD>
                </TR>
                <TR>
                  <TD>
                    <TABLE  border="1" bordercolor="#000000" cellpadding="1" cellspacing="1" width="100%">
                      <TBODY> 
                      <TR> 
                        <TD  width="15%" height="20" valign="bottom"> 
                          <DIV align="center"><B>物品</B></DIV>
                        </TD>
                        <TD  width="38%" height="20" valign="bottom"> 
                          <DIV align="center"><B>功能</B></DIV>
                        </TD>
                        <TD  width="29%" height="20" valign="bottom"> 
                          <DIV align="center"><B>售价</B></DIV>
                        </TD>
                        <TD  width="18%" height="20" valign="bottom"> 
                          <DIV align="center"><B>操作</B></DIV>
                        </TD>
                      </TR>
<!--#include file="data.asp"--><% 
sql="SELECT 物品名,说明,银两,id FROM 商店 where  类型='龙粮'"
rs.open sql,connb,1,1
do while not rs.bof and not rs.eof
%>
                      <TR> 
                        <TD width="15%"  height="20" valign="bottom"> 
                          <%=rs("物品名")%></TD>
                        <TD width="38%"  height="20" valign="bottom"> 
                          <div align="left"><%=rs("说明")%></div>
                        </TD>
</CENTER>
                        <TD width="29%"  height="20" valign="bottom"> 
                          <DIV align="center">
                            <p align="left">银两:<%=rs("银两")%></DIV>
                        </TD>
<CENTER>
                        <TD width="18%"  height="20" valign="bottom"> 
                          <DIV align="center"><a href="longeat1.asp?mz=<%=rs("id")%>">购买</a></DIV>
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
                  <TD align="center"  height="20" valign="bottom"> 
                    <br>
                    <p align="center"><input type=button value='返回' onclick="javascript:location.href='../../t/ys/xx.asp'"></p>
                  </TD>
                </TR>
              </TBODY>
            </TABLE>
            </TD>
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
