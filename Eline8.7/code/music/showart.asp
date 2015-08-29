<!--#include file="config.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="format.asp"-->
<!--#include file="error.asp"-->
<!--#include file="lib/incjs.asp"-->
<!--#include file="lib/inctop.asp"-->
<!--#include file="lib/articlecat.asp"-->
<!--#include file="lib/showart_body.asp"-->
<!--#include file="lib/nav.asp"-->
<!--#include file="lib/incfooter.asp"-->
<%
stats="技术文章"
dim founderr,errmsg
founderr=false
errmsg=""

if request.querystring("art_id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>请指定操作的对象！"
else
  if not isInteger(request.querystring("art_id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的文章id参数。"
  end if
end if
if request.querystring("cat_id")="" then
  founderr=true
  errmsg=errmsg+"<br>"+"<li>请指定操作的对象！"
else
  if not isInteger(request.querystring("cat_id")) then
    founderr=true
    errmsg=errmsg+"<br>"+"<li>非法的文章分类id参数。"
  end if
end if 

if founderr then
  call diserror()
  response.end
else
call top()
call article_nav()
rcount
call showart_body()
call endpage()
end if

function rcount()
OPENarticle
set rscount=conn.execute("select * from art")
sql="UPDATE art SET art_count = art_count + 1 where art_id="&request.querystring("art_id")
conn.execute (sql)
rscount.close
set rscount=nothing
end function
%>