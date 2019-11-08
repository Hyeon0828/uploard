<html>
<head>
</head>
<script language = javascript>

function ok() {
  if (frm_up.text_title.value.length == 0) {
   alert("제목을 입력해주세요.")
   form_up.txt_title.focus();
      }

  else if (frm_up.text_name.value.length == 0) {
   alert("이름을 적어주세요.")
   form_up.txt_name.focus();
      }

  else if (frm_up.text_area.value.length == 0) {
   alert("내용을 적어주세요.")
   frm_up.text_area.focus();
      }

  else if(frm_up.text_pw.value.length <4) {
   alert("수정을 위한 PW 4자리 입력해주세요")
   frm_up.text_pw.focus();
  }

  else {
    frm_up.submit();
      }
    }

</script>
<body onload="frm_up.text_title.focus();">
<form action="board_input.asp" method="get" name="frm_up">
  <center>
  <table border="0" width="80%">
  <tr>
    <td colspan="2">제목<input type="text" name="text_title" maxlength="40" style="width: 100%;">
    </td>
  </tr>

  <tr>
    <td>작성자 <br><input type="text" name="text_name" maxlength="6" readonly
    value=<% if Session("user_id") = "" then
    response.write "비회원"
    else
    response.write session("user_id")
    end if %>>
    </td>

   <tr>
     <td>비밀번호 <br><input type="text" name="text_pw" maxlength="4" placeholder="비밀번호 4자리 입력">
     </td>
   </tr>

  </tr>
	<td>내용</td>
  </table>
  <textarea name="text_area" rows=15% cols=70%></textarea><p>
  <input type="button" value="뒤로" onclick="history.go(-1);">
  <input type="button" value="입력" onclick="ok();">
  </center>
  </form>
  </body>
</html>
