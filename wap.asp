<% 
Response.ContentType = "text/vnd.wap.wml" 
'response.cookies("test") = "mike"
response.cookies("test").expires= "6/6/2000"
%>

<?xml version="1.0"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN"
			"http://www.wapforum.org/DTD/wml_1.1.xml">

<wml>
<card title="SyncIT">
<p><% 
'x = Request.ServerVariables("ALL_HTTP")
'response.write replace(x,"HTTP","<br/><br/>HTTP")
response.write replace(Request.ServerVariables("HTTP_USER_AGENT"),";",";<br/>")
%></p>
</card>   
</wml>
