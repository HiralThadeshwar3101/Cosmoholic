<%
	dim conn,res
	set conn=Server.CreateObject("ADODB.Connection")
	conn.provider="Microsoft.Jet.OLEDB.4.0"
	conn.open "C:\inetpub\wwwroot\Ridesharing\wt\loginpro.mdb"
	set res=Server.CreateObject("ADODB.Recordset")
	res.open "loginpro",conn,0,3,2
	dim uname,pass
	uname=Request.form("uname")
	pass=Request.form("pass")
	Do While not res.EOF
		If(uname=res("username") AND pass=res("password")) then
			Response.redirect("index.html")
		Else
			Response.redirect("login3.html")
		End If
		res.MoveNext
	Loop
%>