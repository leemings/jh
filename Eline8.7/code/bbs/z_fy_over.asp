<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="z_fy_const.asp"-->
<%
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_over.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
Response.Buffer=True 
stats="状纸一览"
call nav()
call head_var(0,0,"社区法院","z_fy_fayuan.asp")
if not founduser then
  Errmsg=Errmsg+"<br>"+"<li>您没有进入社区法院的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
  call dvbbs_error()
else
  call getfyconfig()
  dim fysql,curpage,StartPageNum,EndPageNum,jgstr,fyrs
  If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
    CurPage = 1
  Else
    CurPage = CINT(Request.QueryString("CurPage"))
  End If
  response.write "<form action=z_fy_over.asp method=get>"
  set rs=server.createobject("adodb.recordset")    
  rs.open "SELECT * FROM fy  ORDER BY dateandtime DESC",connfy,1,1    
  if not rs.eof and not rs.bof then              
    RS.PageSize=15           
    Dim TotalPages              
    TotalPages = RS.PageCount              
    If CurPage>RS.Pagecount Then               
      CurPage=RS.Pagecount              
    end if              
    RS.AbsolutePage=CurPage            
    rs.CacheSize = RS.PageSize              
    Dim Totalcount              
    Totalcount =INT(RS.recordcount)
    response.write "<table cellpadding=3 cellspacing=1 align=center class=tableborder1><tr><th valign=middle colspan=2 align=center height=25 id=tabletitlelink><a href=z_fy_write.asp><b>我要举报</b></a></td></tr><tr><td valign=middle class=tablebody1 height=100><CENTER>"
       response.write "<br><table cellpadding=3 cellspacing=1 align=center class=tableborder1 style='width:97%'>"
    response.write "<tr><th height=25 align=center width='10%'>被  告</td><th width='10%'>日　期</td>"
    response.write "<th width='20%'>标题（点击查看详情）</td><th width='10%'>原  告</td>"
    response.write "<th width='16%'>"
    if master or fymaster then
      response.write "处理"
    else 
      response.write "原告要求"
    end if
    response.write "</td><th width='12%'>审判结果</td><th width='19%'>目前状态</td></tr>"  
    dim p
    I=0                  
    p=RS.PageSize*(Curpage-1)                  
    do while (Not RS.Eof) and (I<RS.PageSize)                  
      p=p+1
      '自动释放囚犯
      if rs(6)<>"7" then
  if (rs(7)="入狱10分钟" and datediff("n",rs(8),now())>10) or (rs(7)="入狱半小时" and datediff("n",rs(8),now())>30) or (rs(7)="入狱一小时" and datediff("n",rs(8),now())>60) or (rs(7)="入狱一天" and datediff("d",rs(8),date())>1) or (rs(7)="入狱三天" and datediff("d",rs(8),date())>3) or (rs(7)="入狱一周" and datediff("d",rs(8),date())>7)  then
          sql="update [user] set lockuser=0 where username='" & rs(2) & "'"
          conn.execute sql
          fysql="update fy set stats='7',dateandtime=now() where id=" & rs(0) & ""
          connfy.execute fysql
          mysign="0|0"
          content="刑期已满，"&rs(2)&" 被释放了"
          call logs("监狱","出狱",membername,rs(2))
        end if
      end if
      response.write "<tr><td height=24 width='10%' align=center class=tablebody1><a href=dispuser.asp?name="&rs(2)&" target=_blank>"&left(rs(2),10)&"</a></td>"
      response.write "<td width='10%' align=center class=tablebody2>"&FormatDateTime(rs(8),2)&"</td>"
      response.write "<td width='20%' class=tablebody1>"
          %>
      <a href=# title='<%=rs(1)%>' onClick="javascript:window.open('z_fy_lookdisp.asp?ID=<%=rs(0)%>','','scrollbars=yes,width=400,height=450,top=10,left=300');">
          <%
    if len(rs(1))>10 then 
       response.write left(rs(1),10)+"..."
    else
       response.write rs(1)
    end if
     response.write "</a></td>"
           response.write "<td width='10%' align=center class=tablebody2><a href=dispuser.asp?name="&rs(4)&" target=_blank>"&left(rs(4),10)&"</a></td>"
      response.write "<td align=center  width='16%' class=tablebody1>"
      if master or fymaster then
        if rs(7)="C" then
          response.write "-----"
        else
          if fymaster and rs(7)<>"N"  then
            response.write "-----"
          else
            response.write "<a href=z_fy_fyaction.asp?id="&rs(0)&">审判</a>"
          end if
        end if
        if pst_set(0)="1" and rs(7)="N" then response.write "|<a href=z_fy_fyaction.asp?faction=peishen&id="&rs(0)&"><font color=blue>送陪审团</font></a>"
     else
        response.write rs(5)
        if membername=rs(4) and rs(7)="N" then
          response.write "|<a href=z_fy_fyaction.asp?faction=ceshu&id="&rs(0)&"><font color=blue>撤诉</font></a>"
        end if
        if membername=rs(2) and rs(7)="N" then
          response.write "|<a href=z_fy_shensu.asp?id="&rs(0)&"><font color=blue>申诉</font></a>"
        end if
      end if
      response.write "</td><td align=center width='12%' class=tablebody2>"
      select case  rs(7)
          case "N" 
              response.write "尚无结果"
          case "B"
               response.write "原告败诉"
          case "P"
               response.write "已平反"
          case "C"
               response.write "原告撤诉"
          case else
               response.write rs(7)
       end select
      response.write "</td><td align=center  width='19%' class=tablebody1>"
      select case rs(6)
       case "N"
        jgstr="等待审理" 
       case "0"
        jgstr="事实不符，驳回上诉"  
       case "1"
        jgstr="事实确凿，已执行" 
       case "2"
        jgstr="经调查，属冤假错案，已平反" 
       case "4"
        jgstr="原告诬告，败诉" 
       case "5"
        jgstr="原告严重诬告，受重罚" 
       case "6"
        jgstr="法官已酌情进行判决" 
       case "7"
        jgstr="刑期已满，已被释放" 
       case "8"
        jgstr="原告撤诉"
      end select
      response.write jgstr & "</td></tr>"
      rs.movenext                  
      I=I+1                  
    loop
    response.write "</table>"     
    StartPageNum=1                 
    do while StartPageNum+15<=CurPage                 
    StartPageNum=StartPageNum+15                 
  Loop                 
  EndPageNum=StartPageNum+14                 
  If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount                 
  response.write "<table width='97%' border=0 cellspacing=0 cellpadding=0 align=center class=tableborder1><tr>" 
  response.write "<td width='47%' align=left class=tablebody1>目前共有<b><font color=#FF0000>"&Totalcount&"</font></b> 条诉状，共<b><font color=#FF0000>"&TotalPages&"</font></b> 页，目前为第<b><font color=#FF0000>"&CurPage&"</font></b>页</td>"
  response.write "<td width='1%' align=center class=tablebody1>&nbsp;</td>"
  response.write "<td width='52%' align=right class=tablebody1>"
  if master then
    response.write "<a href='z_fy_fyaction.asp?faction=alldel'><font color=red>[删除已判]</font></a>&nbsp;"
  end if
  response.write "<a href=z_fy_fayuan.asp>[返回]</a>&nbsp;<a href=z_fy_over.asp>[刷新]</a>&nbsp;<a href=z_fy_over.asp?Curpage=1>[首页]</a>&nbsp;<a href=z_fy_over.asp?Curpage="&TotalPages&"></a><a href=z_fy_over.asp?Curpage="&(Curpage-1)&">[上一页]</a>&nbsp;<a href=z_fy_over.asp?Curpage="&Curpage+1&">[下一页]</a>&nbsp;<a href=z_fy_over.asp?Curpage="&TotalPages&">[尾页]</a></td>"
  response.write "</tr></table></center></td></tr></table>"
 else
  response.write "<br><div align=center><p>暂无状纸，<a href=z_fy_write.asp>递交状纸</a></p></div>"
 end if
 response.write "</form>"
 rs.close  
 set rs=nothing   
 end if  
call footer()  
%> 
