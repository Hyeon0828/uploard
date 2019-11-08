<html>
<head>
</head>
<style>
   .button {

    width:48%;

    background-color: black;

    border: none;

    color:white;

    padding: 15px 0;

    text-align: center;

    text-decoration: none;

    display: inline-block;

    font-size: 15px;

    cursor: pointer;

	padding-left:1px;
}
</style>

  <title>ID, PW 찾기</title>

  <script language=javascript>
    function cancle() {
      self.close();
    }


  </script>

 <body onload="ss_main.text_id2.focus()">
     <div class="back">
       <center>

         <a href="find_id.asp" class="button" target="_self">
           <font>ID찾기</font></a>
         <a href="find_pw.asp" class="button" target="_self">
           <font>PW찾기</font></a><br><br>
         <div>
              <form name=ss_main action="find_pw.asp" method="post">
                <input type="text" name="text_id2" maxlength="9" size=10 placeholder="ID">
                <input type="text" name="text_name2" maxlength="6" size=8 placeholder="성명">
                <input type="submit" value="찾기"><p>
              </form>

         <%
           Dim sc_name, sc_birth, sc_id, dbMain
           sc_name = Request("text_name2")

           sc_id = Request("text_id2")


            if(sc_id = "") then
               Response.Write "<b>ID, 성명, 생년월일을 입력해주세요.</b>" &_
              "<p>" &_
              "</div><br><input type='button' value='닫기' OnClick='cancle();'><br>"

            else
             Set db = Server.createObject ("ADODB.Connection")
             dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
             db.Open (dbMain)

             Dim search_pw, stl
             stl = "select * from ex1table where db_name ='" & sc_name & "'" & " and db_id ='" & sc_id & "'"
             Set search_pw = db.execute(stl)


             if search_pw.EOF then
              Response.Write "<font color=red><b>잘못 입력하셨습니다.</b></font>" &_
               "<p><b>다시 입력해주세요.</b><p>" &_
               "</div><br><input type='button' value='닫기' OnClick='cancle();'><br>"

             else
              response.Write "<font color=black><b>PW를 찾았습니다. </b></font><p>"
                do while not search_pw.EOF
                Dim indeX1
                 indeX1 = search_pw("db_pw")
                 response.write "<b><font size=5>" & indeX1 & "</font></b>"
               response.Write "</div><br><input type='button' value='닫기' OnClick='cancle();'><br>"
               search_pw.MoveNext
               Loop
             end if
             db.close()
           set db = nothing
           set search_pw = nothing
           end if
         %>

       </center>
     </div>
</html>
