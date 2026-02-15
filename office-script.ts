/**
 * Excel Office Script: 14 сарапшыға арналған бағалау шаблонын құру
 * - Participants: қатысушы тізімі
 * - E01..E14: сарапшылар енгізетін кесте
 * - Summary: тур бойынша қорытынды (орташа Total)
 */
function main(workbook: ExcelScript.Workbook) {
  const EXPERTS = 14;
  const TOURS = 3; // қажет болса өзгертіңіз
  const CRITERIA = 5; // қажет болса өзгертіңіз (C1..C5)
  const SCALE_MIN = 0; // шкала
  const SCALE_MAX = 10;

  // ---------- helpers ----------
  const ensureSheet = (name: string) => {
    let ws = workbook.getWorksheet(name);
    if (!ws) ws = workbook.addWorksheet(name);
    ws.getUsedRange()?.clear();
    return ws;
  };

  const pad2 = (n: number) => n.toString().padStart(2, "0");

  // ---------- Participants ----------
  const wsP = ensureSheet("Participants");
  wsP.getRange("A1").setValue("ParticipantID");
  wsP.getRange("B1").setValue("FullName");
  wsP.getRange("A1:B1").getFormat().getFont().setBold(true);

  // demo қатысушылар (қаласаң өшіріп таста)
  const demo = [
    ["P001", "Қатысушы 1"],
    ["P002", "Қатысушы 2"],
    ["P003", "Қатысушы 3"],
    ["P004", "Қатысушы 4"],
  ];
  wsP.getRangeByIndexes(1, 0, demo.length, 2).setValues(demo);
  wsP.getUsedRange()?.getFormat().autofitColumns();

  // ---------- Expert sheets ----------
  for (let e = 1; e <= EXPERTS; e++) {
    const name = `E${pad2(e)}`;
    const wsE = ensureSheet(name);

    const headers: (string | number)[] = ["ParticipantID", "Tour"];
    for (let c = 1; c <= CRITERIA; c++) headers.push(`C${c}`);
    headers.push("Total", "Comment");

    wsE.getRangeByIndexes(0, 0, 1, headers.length).setValues([headers]);
    wsE
      .getRangeByIndexes(0, 0, 1, headers.length)
      .getFormat()
      .getFont()
      .setBold(true);

    // Total формуласы (әр жолға)
    // C1..Cn бағандары 3-бағаннан басталады (0-index: 2..)
    const totalColIndex = 2 + CRITERIA;
    const firstDataRow = 1;
    const rowsForEntry = 300; // қанша жазба енгізіледі (қаласаң өсір)

    const totalRange = wsE.getRangeByIndexes(
      firstDataRow,
      totalColIndex,
      rowsForEntry,
      1
    );
    // Excel формуласы: =SUM(C2:G2) тәрізді
    totalRange.setFormulaR1C1(
      `=IF(OR(RC[-${totalColIndex}]="",RC[-${totalColIndex - 1}]=""),"",SUM(RC[-${CRITERIA}]:RC[-1]))`
    );

    // Турға Data Validation (1..TOURS)
    const tourRange = wsE.getRangeByIndexes(firstDataRow, 1, rowsForEntry, 1);
    const tourList = Array.from({ length: TOURS }, (_, i) => (i + 1).toString()).join(",");
    tourRange.getDataValidation().setRule({
      list: { inCellDropDown: true, source: tourList },
    });

    // Критерийлерге Data Validation (SCALE_MIN..SCALE_MAX)
    for (let c = 0; c < CRITERIA; c++) {
      const r = wsE.getRangeByIndexes(firstDataRow, 2 + c, rowsForEntry, 1);
      r.getDataValidation().setRule({
        wholeNumber: {
          formula1: SCALE_MIN.toString(),
          formula2: SCALE_MAX.toString(),
          operator: ExcelScript.DataValidationOperator.between,
        },
      });
    }

    wsE.getUsedRange()?.getFormat().autofitColumns();
    wsE.freezePanes.freezeRows(1);
  }

  // ---------- Summary ----------
  const wsS = ensureSheet("Summary");
  wsS.getRange("A1").setValue("ParticipantID");
  wsS.getRange("B1").setValue("FullName");
  wsS.getRange("A1:B1").getFormat().getFont().setBold(true);

  // Тур тақырыптары
  for (let t = 1; t <= TOURS; t++) {
    wsS.getRangeByIndexes(0, 1 + t, 1, 1).setValue(`Tour ${t} AVG`);
    wsS.getRangeByIndexes(0, 1 + t, 1, 1).getFormat().getFont().setBold(true);
  }
  wsS.getRangeByIndexes(0, 2 + TOURS, 1, 1).setValue("Overall AVG");
  wsS.getRangeByIndexes(0, 2 + TOURS, 1, 1).getFormat().getFont().setBold(true);

  // Participants-тан тізімді әкелу (формуламен)
  // A2: =FILTER(Participants!A2:A999,Participants!A2:A999<>"")
  wsS.getRange("A2").setFormula(
    `=FILTER(Participants!A2:A999,Participants!A2:A999<>"")`
  );
  wsS.getRange("B2").setFormula(`=XLOOKUP(A2#,Participants!A:A,Participants!B:B,"")`);

  // Әр турға 14 сарапшыдан орташа Total есептеу:
  // =LET(pid,$A2,t,1, AVERAGE(
  //   IFERROR(AVERAGEIFS(E01!$H:$H,E01!$A:$A,pid,E01!$B:$B,t),""), ... ))
  //
  // Total баған индексі: ParticipantID(1), Tour(2), C1..Cn, Total => 2 + CRITERIA + 1
  // Excel әріппен табу қиын болғандықтан, толық бағандар арқылы AVERAGEIFS жасаймыз:
  // Total бағанын табу үшін E** листінде Total әрқашан (2+CRITERIA+1) позицияда.
  // Бірақ Office Script-та баған әрпін есептеп береміз:

  const totalColLetter = colToLetter(2 + CRITERIA + 1); // 1-based column index
  // ParticipantID = A, Tour = B
  // Total = computed letter

  const firstRow = 2;

  for (let t = 1; t <= TOURS; t++) {
    const col = 2 + t; // C.. etc (A=1, B=2)
    const startCell = wsS.getCell(firstRow - 1, col - 1); // 0-based
    // dynamic spill A2# бар, сондықтан формуланы бір ұяшыққа қоямыз:
    // =MAP(A2#,LAMBDA(pid, AVERAGE( ... )))
    const parts: string[] = [];
    for (let e = 1; e <= EXPERTS; e++) {
      const sh = `E${pad2(e)}`;
      parts.push(
        `IFERROR(AVERAGEIFS(${sh}!$${totalColLetter}:$${totalColLetter},${sh}!$A:$A,pid,${sh}!$B:$B,${t}),"")`
      );
    }
    const formula = `=MAP(A2#,LAMBDA(pid,LET(x,${parts.join(",")},AVERAGEIF(x,"<>"))))`;
    startCell.setFormula(formula);
  }

  // Overall AVG = барлық турдың орташа мәні
  // =MAP(A2#,LAMBDA(pid,AVERAGEIF(HSTACK(t1,t2,t3),"<>")))
  const overallCell = wsS.getCell(firstRow - 1, 1 + TOURS + 1); // 0-based
  // tourCols addresses look like C2 etc for spilled MAP outputs; for HSTACK we need full spilled arrays:
  // we can directly refer to those MAP outputs as ranges like C2#, D2# ...
  const tourSpills = Array.from({ length: TOURS }, (_, i) => `${colToLetter(3 + i)}2#`);
  overallCell.setFormula(`=AVERAGEIF(HSTACK(${tourSpills.join(",")}),"<>")`);

  wsS.getUsedRange()?.getFormat().autofitColumns();
  wsS.freezePanes.freezeRows(1);

  // Done.

  // ---- local helper for column letter ----
  function colToLetter(col: number): string {
    let temp = col;
    let letter = "";
    while (temp > 0) {
      const mod = (temp - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      temp = Math.floor((temp - 1) / 26);
    }
    return letter;
  }
}
