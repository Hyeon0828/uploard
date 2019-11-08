<%
if Session("user_id") = "" then
response.write "<script language=javascript>window.open('d_log.html', '_self');</script>"

else
response.write "<script language=javascript>window.open('login_ok.asp', '_self');</script>"
end if
%>
