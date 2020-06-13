<%@ Page Language="VB" %>
<% 

If InStr( UCase(Request.ServerVariables("SERVER_NAME")), UCase("mysalonk.com") ) > 0 Then 
        Response.RedirectPermanent("/sk/") 

ElseIf InStr( UCase(Request.ServerVariables("SERVER_NAME")), UCase("shippychalet.com") ) > 0 Then 
        Response.RedirectPermanent("/sc/")

ElseIf InStr( UCase(Request.ServerVariables("SERVER_NAME")), UCase("shippychateau.com") ) > 0 Then 
        Response.RedirectPermanent("/chateau/")
        
ElseIf InStr( UCase(Request.ServerVariables("SERVER_NAME")), UCase("dakotacreekcrosses.com") ) > 0 Then 
        Response.RedirectPermanent("/dcc/")
       
ElseIf InStr( UCase(Request.ServerVariables("SERVER_NAME")), UCase("tweeteverytrade.com") ) > 0 Then 
        Response.RedirectPermanent("/tw/")

ElseIf InStr( UCase(Request.ServerVariables("SERVER_NAME")), UCase("arizonarhinestone.com") ) > 0 Then 
        Response.RedirectPermanent("/azr/")

ElseIf InStr( UCase(Request.ServerVariables("SERVER_NAME")), UCase("shippysigncompany.com") ) > 0 Then 
        Response.RedirectPermanent("/ssc/")
		
ElseIf InStr( UCase(Request.ServerVariables("SERVER_NAME")), UCase("salonkchandler.com") ) > 0 Then 
        Response.RedirectPermanent("/skwp")
		
Else
        Response.RedirectPermanent("/tpauls.htm") 
End If 
%>