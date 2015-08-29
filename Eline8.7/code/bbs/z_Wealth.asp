<%
function bank_money(username)
	const dbbank="data/e_bank.asp"
	dim connmoney
	dim rsmoney
	Set connmoney = Server.CreateObject("ADODB.Connection")
	connmoney.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&dbbank&"")
	set rsmoney=connmoney.execute("select savemoney from bank where username='"&username&"'")
	if rsmoney.bof and rsmoney.eof then
		bank_money=0
	else
		bank_money=clng(rsmoney(0))
	end if
	rsmoney.close		
	connmoney.close
	set rsmoney=nothing
	Set connmoney = Nothing
end function

function stock_money(username)
	const dbStock="data/e_stock.asp"
	dim connmoney
	dim rsmoney
	Set connmoney = Server.CreateObject("ADODB.Connection")
	connmoney.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&dbStock&"")
	set rsmoney=connmoney.execute("select ZongZiJin from KeHu where ZhangHao='"&username&"'")
	if rsmoney.bof and rsmoney.eof then
		Stock_money=0
	else
		Stock_money=clng(rsmoney("ZongZiJin"))
	end if
	rsmoney.close		
	connmoney.close
	set rsmoney=nothing
	Set connmoney = Nothing
end function
%>