<%
Dim addSQL
Dim add_id, add_pw, add_name, add_gender, add_birth, add_email1, add_email2, add_phone1, add_phone2, add_phone3, add_zip1, add_zip2, add_zip3, add_my
 add_id = request("IDtext1")
 add_pw = request("Pw")
 add_name = request("txt_name")
 add_gender = request("gender")
 add_birth = request("birth")
 add_email1 = request("e_mail")
 add_email2 = request("domain")
 add_phone1 = request("phone1")
 add_phone2 = request("phone2")
 add_phone3 = request("phone3")
 add_zip1 = request("home_zip1")
 add_zip2 = request("home_zip2")
 add_zip3 = request("home_add1")
 add_my = request("my")


 if request("gender") = "1" then
  add_gender = "남자"
 else
  add_gender = "여자"
 end if

 if request("domain") = "1" then add_email2 = "naver.com"
 if request("domain") = "2" then add_email2 = "hanmail.net" 
 if request("domain") = "3" then add_email2 = "nate.com"
 if request("domain") = "4" then add_email2 = "gmail.com"




'RESPONSE.WRITE(add_id & "/" &  add_pw & "/" &  add_name & "/" & add_gender & "/" & add_birth & "<br>")
'RESPONSE.WRITE(add_email1 & "/" & add_email2 & "/" & add_phone1 & "/" & add_phone2 & "/" & add_phone3 & "<br>")


Set db = Server.createObject ("ADODB.Connection")
 con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
 db.open(con)

addSQL = "INSERT INTO ex1table (db_id, db_pw, db_name, db_gender, db_birth, db_email1, db_email2, db_phone1, db_phone2, db_phone3, db_zip1, db_zip2, db_zip3, db_my,) values"
addSQL = addSQL & "('" & add_id & "','" & add_pw & "','" & add_name & "','" & add_gender & "','" & add_birth & "','" & add_email1 & "','" & add_email2 & "','" & add_phone1 & "','" & add_phone2 & "','" & add_phone3 & "','" & add_zip1 & "','" & add_zip2 & "','" & add_zip3 & "','" & add_my & "');"

response.write (addSQL)
db.Execute(addSQL)
db.close()
set db = nothing

%>

