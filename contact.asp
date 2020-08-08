<%@language="vbscript"%>
<% Dim tsql,n,e,f
Set con=Server.CreateObject("ADODB.Connection")
con.Provider="Microsoft.JET.OLEDB.4.0"
con.open "C:\inetpub\wwwroot\Ridesharing\wt\contact.mdb"
Set res=Server.CreateObject("ADODB.Recordset")
res.open "contact",con,0,3,2
n=Request.Form("name")
e=Request.Form("email")
f=Request.Form("feedback")


tsql="INSERT INTO contact VALUES('"&n&"','"&e&"','"&f&"')"
con.Execute(tsql)
Response.redirect("successful.html")
%>