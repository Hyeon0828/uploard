<html>

  <%
     db_id = request("login_id")
     db_pw = request("login_pw")

     Dim con, db
     Set db = Server.createObject ("ADODB.Connection")
     con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
     db.Open (con)

     Dim login_db, sql
     sql = "select * from ex1table where db_id ='" & db_id & "'" & " and db_pw ='" & db_pw & "'"
     Set login_db = db.execute(sql)

     if login_db.EOF then
      Response.Write  "<script>alert('ID 또는 PW가 잘못됬습니다.');history.go(-1);</script>"
      db.Close
      login_db.close

     Else
      session("user_id") = login_db("db_id")
      session("user_name") = login_db("db_name")
      session("user_pw") = login_db("db_pw")

      login_db.close
      set login_db = nothing
      db.Close
      set db = nothing
      Response.Write "<script language=javascript>parent.location.reload();</script>"

     End If
  %>
</html>
