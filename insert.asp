<%@language="vbscript"%>
<% Dim con,u,p,res,tsql,fn,mn,ln,e,m
Set con=Server.CreateObject("ADODB.Connection")
con.Provider="Microsoft.JET.OLEDB.4.0"
con.open "C:\inetpub\wwwroot\Ridesharing\wt\loginpro.mdb"
Set res=Server.CreateObject("ADODB.Recordset")
res.open "loginpro",con,0,3,2
fn=Request.Form("fname")
mn=Request.Form("mname")
ln=Request.Form("lname")
e=Request.Form("email")
m=Request.Form("mob")
u=Request.Form("uname")
p=Request.Form("pass")

tsql="INSERT INTO loginpro VALUES('"&fn&"','"&mn&"','"&ln&"','"&e&"','"&m&"','"&u&"','"&p&"')"
con.Execute(tsql)
Response.redirect("index.html")
%>