// onEdit(e) 트리거이므로 시트 수정시 자동으로 실행됨, 수동 디버깅 불가



function onEdit(e) {
  var activeSheet = e.source.getActiveSheet();
/*
  2차 모니터링에서 변경이 일어난 경우
  1) '보고서 해당'을 체크하면 '보고서 Dashboard'의 맨 마지막 줄에 값을 복사하고 
  2) numbering formula를 기입한다.
*/
  var range = e.range, s = range.getSheet(), row = range.getRow();
  if (s.getName() === '2차 모니터링' && range.getColumn() === 15) {
    var sheetTarget = e.source.getSheetByName('Dashboard(임시)');
    
    // 2차 모니터링 정보를 Dashboard로 복사
    // slice 등 써서 간략하게 수정 필요
    var srcRange0 = s.getRange(row, 2, 1, 2);  // 월, 일자
    var srcRange1 = s.getRange(row, 4, 1, 6);  // 구분 ~ GICS LEVEL
    var srcRange2 = s.getRange(row, 16, 1, 1); // 담당자

    var trgrow = sheetTarget.getLastRow()+1 // Dashboard 시트의 마지막 행 + 1 (새로 추가되는 행 번호)
    var trgRange0 = sheetTarget.getRange(trgrow,1,1,2);
    var trgRange1 = sheetTarget.getRange(trgrow,3,1,8);
    var trgRange2 = sheetTarget.getRange(trgrow,9,1,1);
    srcRange0.copyTo(trgRange0);  // 월, 일자는 값 복사 아닌 전체복사 (값 복사 하면 형식 깨짐)
    srcRange1.copyTo(trgRange1, {contentsOnly:true});
    srcRange2.copyTo(trgRange2, {contentsOnly:true});

    // 각 열에 numbering formula 기입
    formulaM = '=IF($L'+trgrow+'>1,MAX(Q:Q)+RANK($L'+trgrow+',L:L,1),"")';
    formulaN = '=IF($M'+trgrow+'="","","Cn-2021-"&RIGHT("00"&M'+trgrow+',2))';
    formulaQ = '=IF($P'+trgrow.toStirng()+'="","",VALUE(RIGHT($P'+trgrow.toStirng()+',2)))';

    sheetTarget.getRange(trgrow,13,1,1).setFormula(formulaM);
    sheetTarget.getRange(trgrow,14,1,1).setFormula(formulaN);
    sheetTarget.getRange(trgrow,17,1,1).setFormula(formulaQ);
    

  }

// https://stackoverflow.com/questions/50368481/google-script-copy-row-to-another-sheet-on-edit-almost-there
// https://stackoverflow.com/questions/19982060/copy-row-to-separate-spreadsheet-onedit-of-column


/*
  보고서 Dashboard에서 변경이 일어난 경우
  1) '작성완료'를 체크하면 Cn code를 할당한다.
  2) '송부'를 체크하면 Cn code를 확정한다.
*/
  if (s.getName() == "보고서 Dashboard") {                     // 시트 이름
    var aCell = e.source.getActiveCell(), col = aCell.getColumn();
    
    // '작성완료' (K=11) 체크시
    if (col == 11) {                                                // checkbox 위치한 column
    // OFFSET(ROW(- up, + down), COL(- left, + right)) 만큼 이동
    // dateCell: timestamp 입력하는 cell, postCell: 송부버튼 cell
      var dateCell = aCell.offset(0,1), postCell = aCell.offset(0,4);
      if (aCell.getValue() == true) {
        dateCell.setValue(new Date()).setNumberFormat("yyyy-MM-dd hh:mm");
        postCell.insertCheckboxes();
      } 
      else {
        dateCell.setValue("");
        postCell.removeCheckboxes();
      }
    }

    // '송부' (O=15) 체크시
    if (col == 15){
      var copyfrom = aCell.offset(0,-1), copydest = aCell.offset(0,1); // Cn code source/dest
      var numfrom = aCell.offset(0,-2), numdest = aCell.offset(0,2); // numbering source/dest
      if (aCell.getValue() == true){
        copyfrom.copyTo(copydest,{contentsOnly: true}), copyfrom.clearContent();  // Cn Code(확정) 복사 후 삭제
        numfrom.copyTo(numdest, {contentsOnly: true}), numfrom.clearContent();    // numbering 복사(발간 번호 확인용) 후 삭제
        copyfrom.offset(0,-2).clearContent(); // timestamp 삭제
      }
      else{
        copydest.copyTo(copyfrom,{contentsOnly: true});
        copydest.setValue("");
      }
    }
  }
}

/* 
참고: CHECKBOX가 TRUE일 때 옆에 날짜+시간 적는 함수 https://webapps.stackexchange.com/questions/130000/how-do-you-get-current-date-to-be-added-to-a-cell-when-a-check-box-is-checked-in

참고: 날짜+시간에 따른 numbering 함수 https://superuser.com/questions/816129/sort-ranking-functions-in-excel-for-time-and-number-values-in-one-column
*/
