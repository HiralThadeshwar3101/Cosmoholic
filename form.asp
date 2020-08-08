<%@  language="vbscript" %>
<%Option Explicit%>
<%
	dim conn,res,a,b,c,d,e,f,g
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.provider="Microsoft.JET.OLEDB.4.0"
	conn.open "C:\inetpub\wwwroot\yogeshnew\login.mdb"
	a=Request.Form("name")
	b=Request.Form("city")
	c=Request.Form("email")
	d=Request.Form("state")
	e=Request.Form("address")
	f=Request.Form("land")
	g=Request.Form("pin")
	set res=Server.CreateObject("ADODB.Recordset")
	res.open "login",conn,0,3,2
	res.AddNew
	res("name1")=a
	res("city")=b
	res("email")=c
	res("state")=d
	res("address")=e
	res("landmark")=f
	res("pincode")=g
	res.Update
	res.Movenext
	Response.write"Your Product will be delivered to you Within 7 working days"
%>
