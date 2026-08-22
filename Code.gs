/*************************************************
 POTATOWALA ROYALTY TRACKER
 FULL Code.gs
**************************************************/

const SHEET_ID =
  '1jwz1CkscOaCTezzuXh7bZo80VVdH-ENa2Yi2QEd1tn4';

const DEFAULT_ROYALTY = 5900;


/*************************************************
 WEB APP
**************************************************/

function doGet(e) {
  const api = pwApiHandle_(e);
  if (api) return api;

  setupSheets();

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('POTATOWALA ROYALTY TRACKER')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function getSS() {

  return SpreadsheetApp.openById(
    SHEET_ID
  );

}


/*************************************************
 SHEET SETUP
**************************************************/

function setupSheets() {

  const ss = getSS();


  /******** OUTLETS ********/

  let outlets =
    ss.getSheetByName(
      'Outlets'
    );


  if (!outlets) {

    outlets =
      ss.insertSheet(
        'Outlets'
      );

  }


  if (
    outlets.getLastRow() === 0
  ) {

    outlets.appendRow([
      'Outlet Name',
      'Outlet Type',
      'Opening Date',
      'Royalty Start Date',
      'Monthly Royalty',
      'Status'
    ]);

  }


  outlets
    .getRange(1,1,1,6)
    .setFontWeight('bold');

  outlets.setFrozenRows(1);


  const defaults = [

    [
      'Kathipara',
      'Own',
      '2022-06-22',
      '',
      0,
      'Active'
    ],

    [
      'Besant Nagar',
      'Own',
      '2023-09-14',
      '',
      0,
      'Active'
    ],

    [
      'Akkarai',
      'Franchise',
      '2024-09-15',
      '2025-01-01',
      5900,
      'Active'
    ],

    [
      'Porur',
      'Franchise',
      '2025-01-25',
      '2025-04-01',
      5900,
      'Active'
    ],

    [
      'Pallavaram',
      'Franchise',
      '2025-12-11',
      '2026-03-01',
      5900,
      'Active'
    ],

    [
      'Sithalapakkam',
      'Franchise',
      '2025-12-29',
      '2026-03-01',
      5900,
      'Active'
    ],

    [
      'Ampa',
      'Franchise',
      '2026-06-07',
      '2026-09-01',
      5900,
      'Active'
    ]

  ];


  ensureDefaultOutlets(
    outlets,
    defaults
  );


  /******** MONTHLY DATA ********/

  let monthly =
    ss.getSheetByName(
      'Monthly Data'
    );


  if (!monthly) {

    monthly =
      ss.insertSheet(
        'Monthly Data'
      );

  }


  if (
    monthly.getLastRow() === 0
  ) {

    monthly.appendRow([
      'ID',
      'Entry Date',
      'Month',
      'Outlet',
      'Petpooja Sales',
      'Swiggy Sales',
      'Zomato Sales',
      'Other Sales',
      'Total Sales',
      'Expenses',
      'Profit / Loss',
      'Royalty Amount',
      'Royalty Status',
      'Paid Date',
      'Notes'
    ]);

  }


  monthly
    .getRange(1,1,1,15)
    .setFontWeight('bold');

  monthly.setFrozenRows(1);

}


/*************************************************
 DEFAULT OUTLETS
**************************************************/

function ensureDefaultOutlets(
  sheet,
  defaults
) {

  let existing = [];


  if (
    sheet.getLastRow() >= 2
  ) {

    existing =
      sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          6
        )
        .getValues();

  }


  defaults.forEach(
    function(item) {

      const exists =
        existing.some(
          function(row) {

            return (
              String(row[0])
                .trim()
                .toLowerCase() ===
              String(item[0])
                .trim()
                .toLowerCase()
            );

          }
        );


      if (!exists) {

        sheet.appendRow(
          item
        );

      }

    }
  );


  /*
   Remove old Vellore default
  */

  if (
    sheet.getLastRow() >= 2
  ) {

    const names =
      sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          1
        )
        .getValues();


    for (
      let i = names.length - 1;
      i >= 0;
      i--
    ) {

      if (
        String(names[i][0])
          .trim()
          .toLowerCase() ===
        'vellore'
      ) {

        sheet.deleteRow(
          i + 2
        );

      }

    }

  }

}


/*************************************************
 GET OUTLETS
**************************************************/

function getOutlets() {

  setupSheets();


  const sh =
    getSS()
      .getSheetByName(
        'Outlets'
      );


  if (
    sh.getLastRow() < 2
  ) {

    return [];

  }


  const rows =
    sh
      .getRange(
        2,
        1,
        sh.getLastRow() - 1,
        6
      )
      .getValues();


  return rows

    .filter(
      row => row[0]
    )

    .filter(
      row =>
        String(
          row[5] || 'Active'
        ).toLowerCase()
        !== 'inactive'
    )

    .map(
      row => ({

        name:
          String(row[0]),

        type:
          String(row[1]),

        openingDate:
          formatDateDisplay(
            row[2]
          ),

        royaltyStartDate:
          formatDateDisplay(
            row[3]
          ),

        monthlyRoyalty:
          Number(
            row[4] || 0
          ),

        status:
          String(
            row[5] ||
            'Active'
          )

      })
    );

}


/*************************************************
 ADD OUTLET
**************************************************/

function addOutlet(data) {

  setupSheets();


  if (
    !data ||
    !data.name
  ) {

    throw new Error(
      'Outlet Name enter pannunga'
    );

  }


  if (
    !data.openingDate
  ) {

    throw new Error(
      'Opening Date select pannunga'
    );

  }


  if (
    data.type !== 'Own' &&
    data.type !== 'Franchise'
  ) {

    throw new Error(
      'Outlet Type select pannunga'
    );

  }


  const outlets =
    getOutlets();


  const duplicate =
    outlets.some(
      function(o) {

        return (
          o.name
            .trim()
            .toLowerCase() ===

          data.name
            .trim()
            .toLowerCase()
        );

      }
    );


  if (duplicate) {

    throw new Error(
      'Outlet already exists'
    );

  }


  let royaltyStart = '';
  let royaltyAmount = 0;


  if (
    data.type === 'Franchise'
  ) {

    royaltyStart =
      data.royaltyStartDate ||
      getFourthMonthStart(
        data.openingDate
      );


    royaltyAmount =
      Number(
        data.monthlyRoyalty ||
        DEFAULT_ROYALTY
      );

  }


  const sh =
    getSS()
      .getSheetByName(
        'Outlets'
      );


  sh.appendRow([

    data.name.trim(),

    data.type,

    data.openingDate,

    royaltyStart,

    royaltyAmount,

    'Active'

  ]);


  SpreadsheetApp.flush();


  return {

    success: true,

    message:
      'Outlet Added Successfully'

  };

}


/*************************************************
 MONTHLY ENTRIES
**************************************************/

function getEntries() {

  setupSheets();


  const sh =
    getSS()
      .getSheetByName(
        'Monthly Data'
      );


  if (
    sh.getLastRow() < 2
  ) {

    return [];

  }


  const rows =
    sh
      .getRange(
        2,
        1,
        sh.getLastRow() - 1,
        15
      )
      .getValues();


  return rows.map(
    row => ({

      id:
        String(
          row[0] || ''
        ),

      entryDate:
        formatDateTime(
          row[1]
        ),

      month:
        normalizeMonth(
          row[2]
        ),

      outlet:
        String(
          row[3] || ''
        ),

      petpooja:
        Number(
          row[4] || 0
        ),

      swiggy:
        Number(
          row[5] || 0
        ),

      zomato:
        Number(
          row[6] || 0
        ),

      other:
        Number(
          row[7] || 0
        ),

      totalSales:
        Number(
          row[8] || 0
        ),

      expenses:
        Number(
          row[9] || 0
        ),

      profitLoss:
        Number(
          row[10] || 0
        ),

      royaltyAmount:
        Number(
          row[11] || 0
        ),

      royaltyStatus:
        String(
          row[12] || ''
        ),

      paidDate:
        formatDateDisplay(
          row[13]
        ),

      notes:
        String(
          row[14] || ''
        )

    })
  );

}


/*************************************************
 GET ONE MONTH
**************************************************/

function getMonthEntry(
  outletName,
  month
) {

  const entries =
    getEntries();


  const outlets =
    getOutlets();


  const outlet =
    outlets.find(
      x =>
        x.name ===
        outletName
    );


  const rule =
    outlet
      ? getRoyaltyRule(
          outlet,
          month
        )
      : {
          amount: 0,
          nilRoyalty: true
        };


  const found =
    entries.find(
      function(item) {

        return (
          item.outlet ===
          outletName &&

          item.month ===
          month
        );

      }
    );


  if (!found) {

    return {

      exists: false,

      royaltyAmount:
        rule.amount || 0,

      royaltyStatus:
        rule.nilRoyalty
          ? 'NIL'
          : 'Pending'

    };

  }


  return Object.assign(
    {
      exists: true
    },
    found
  );

}


/*************************************************
 SAVE / UPDATE ENTRY
**************************************************/

function saveMonthlyEntry(
  data
) {

  setupSheets();


  if (
    !data.outlet
  ) {

    throw new Error(
      'Outlet select pannunga'
    );

  }


  if (
    !data.month
  ) {

    throw new Error(
      'Month select pannunga'
    );

  }


  const outlets =
    getOutlets();


  const outlet =
    outlets.find(
      x =>
        x.name ===
        data.outlet
    );


  if (!outlet) {

    throw new Error(
      'Outlet not found'
    );

  }


  const rule =
    getRoyaltyRule(
      outlet,
      data.month
    );


  if (
    rule.beforeOpening
  ) {

    throw new Error(
      'Selected month is before outlet opening date'
    );

  }


  const petpooja =
    Number(
      data.petpooja || 0
    );


  const swiggy =
    Number(
      data.swiggy || 0
    );


  const zomato =
    Number(
      data.zomato || 0
    );


  const other =
    Number(
      data.other || 0
    );


  const expenses =
    Number(
      data.expenses || 0
    );


  const totalSales =
    petpooja +
    swiggy +
    zomato +
    other;


  const profitLoss =
    totalSales -
    expenses;


  let royaltyAmount =
    rule.amount;


  let royaltyStatus =
    data.royaltyStatus ||
    'Pending';


  if (
    rule.nilRoyalty
  ) {

    royaltyAmount = 0;

    royaltyStatus =
      'NIL';

  }


  let paidDate = '';


  if (
    royaltyStatus === 'Paid'
  ) {

    paidDate =
      new Date();

  }


  const sh =
    getSS()
      .getSheetByName(
        'Monthly Data'
      );


  const sheetData =
    sh
      .getDataRange()
      .getValues();


  let existingRow = 0;
  let existingId = '';


  for (
    let i = 1;
    i < sheetData.length;
    i++
  ) {

    if (

      normalizeMonth(
        sheetData[i][2]
      ) === data.month &&

      String(
        sheetData[i][3]
      ) === data.outlet

    ) {

      existingRow =
        i + 1;

      existingId =
        sheetData[i][0];

      break;

    }

  }


  const id =
    existingId ||
    Utilities.getUuid();


  const values = [[

    id,

    new Date(),

    data.month,

    data.outlet,

    petpooja,

    swiggy,

    zomato,

    other,

    totalSales,

    expenses,

    profitLoss,

    royaltyAmount,

    royaltyStatus,

    paidDate,

    data.notes || ''

  ]];


  if (
    existingRow
  ) {

    sh
      .getRange(
        existingRow,
        1,
        1,
        15
      )
      .setValues(
        values
      );

  }

  else {

    sh
      .getRange(
        sh.getLastRow() + 1,
        1,
        1,
        15
      )
      .setValues(
        values
      );

  }


  SpreadsheetApp.flush();


  return {

    success: true,

    totalSales:
      totalSales,

    expenses:
      expenses,

    profitLoss:
      profitLoss,

    royaltyAmount:
      royaltyAmount,

    royaltyStatus:
      royaltyStatus

  };

}


/*************************************************
 DASHBOARD
**************************************************/

function getDashboard(
  selectedOutlet,
  showUpTo
) {

  const outlets =
    getOutlets();


  const entries =
    getEntries();


  const selected =
    outlets.find(
      x =>
        x.name ===
        selectedOutlet
    ) ||
    outlets[0] ||
    null;


  let paid = 0;
  let pending = 0;


  outlets.forEach(
    function(outlet) {

      const rule =
        getRoyaltyRule(
          outlet,
          showUpTo
        );


      if (
        rule.beforeOpening ||
        rule.nilRoyalty
      ) {

        return;

      }


      const entry =
        entries.find(
          x =>
            x.outlet ===
              outlet.name &&
            x.month ===
              showUpTo
        );


      if (
        entry &&
        String(
          entry.royaltyStatus
        ).toLowerCase()
        === 'paid'
      ) {

        paid++;

      }

      else {

        pending++;

      }

    }
  );


  return {

    totalOutlets:
      outlets.length,

    paid:
      paid,

    pending:
      pending,

    selectedOutlet:
      selected,

    selectedMonthRule:
      selected
        ? getRoyaltyRule(
            selected,
            showUpTo
          )
        : null,

    missingMonths:
      selected
        ? getMissingMonthsInternal(
            selected,
            showUpTo,
            entries
          )
        : []

  };

}


/*************************************************
 OVERALL SALES / EXPENSE / PROFIT
**************************************************/

function getOutletOverallSummary(
  outletName,
  showUpTo
) {

  const outlets =
    getOutlets();


  const entries =
    getEntries();


  const outlet =
    outlets.find(
      x =>
        x.name ===
        outletName
    );


  if (!outlet) {

    return {

      totalSales: 0,

      totalExpenses: 0,

      totalProfitLoss: 0,

      openingDate: '',

      showUpTo:
        showUpTo || ''

    };

  }


  const openingMonth =
    monthFromDate(
      outlet.openingDate
    );


  const endMonth =
    showUpTo ||
    currentMonth();


  let sales = 0;
  let expenses = 0;


  entries.forEach(
    function(entry) {

      if (
        entry.outlet !==
        outletName
      ) {

        return;

      }


      if (
        monthToNumber(
          entry.month
        ) <
        monthToNumber(
          openingMonth
        )
      ) {

        return;

      }


      if (
        monthToNumber(
          entry.month
        ) >
        monthToNumber(
          endMonth
        )
      ) {

        return;

      }


      sales +=
        Number(
          entry.totalSales || 0
        );


      expenses +=
        Number(
          entry.expenses || 0
        );

    }
  );


  return {

    totalSales:
      sales,

    totalExpenses:
      expenses,

    totalProfitLoss:
      sales - expenses,

    openingDate:
      outlet.openingDate,

    showUpTo:
      endMonth

  };

}


/*************************************************
 HISTORY
**************************************************/

function getOutletHistory(
  outletName,
  showUpTo
) {

  const outlets =
    getOutlets();


  const entries =
    getEntries();


  const outlet =
    outlets.find(
      x =>
        x.name ===
        outletName
    );


  if (!outlet) {

    return {

      info: null,

      history: []

    };

  }


  const openingMonth =
    monthFromDate(
      outlet.openingDate
    );


  const endMonth =
    showUpTo ||
    currentMonth();


  const months =
    monthsBetween(
      openingMonth,
      endMonth
    );


  const history =
    months.map(
      function(month) {

        const entry =
          entries.find(
            x =>
              x.outlet ===
                outletName &&
              x.month ===
                month
          );


        const rule =
          getRoyaltyRule(
            outlet,
            month
          );


        let status =
          'PENDING';


        if (
          rule.nilRoyalty
        ) {

          status =
            'NIL';

        }

        else if (
          entry &&
          entry.royaltyStatus
        ) {

          status =
            String(
              entry.royaltyStatus
            ).toUpperCase();

        }


        return {

          month:
            month,

          monthName:
            displayMonth(
              month
            ),

          totalSales:
            entry
              ? entry.totalSales
              : 0,

          expenses:
            entry
              ? entry.expenses
              : 0,

          profitLoss:
            entry
              ? entry.profitLoss
              : 0,

          royaltyAmount:
            rule.amount,

          status:
            status

        };

      }
    );


  history.reverse();


  return {

    info:
      outlet,

    history:
      history

  };

}


/*************************************************
 MISSING MONTH REMINDER
**************************************************/

function getMissingMonthsInternal(
  outlet,
  showUpTo,
  entries
) {

  const openingMonth =
    monthFromDate(
      outlet.openingDate
    );


  const previous =
    addMonths(
      showUpTo,
      -1
    );


  if (
    monthToNumber(
      previous
    ) <
    monthToNumber(
      openingMonth
    )
  ) {

    return [];

  }


  const months =
    monthsBetween(
      openingMonth,
      previous
    );


  return months.filter(
    function(month) {

      const found =
        entries.find(
          x =>
            x.outlet ===
              outlet.name &&
            x.month ===
              month
        );


      return !found;

    }
  );

}


/*************************************************
 ROYALTY RULE
**************************************************/

function getRoyaltyRule(
  outlet,
  month
) {

  const openingMonth =
    monthFromDate(
      outlet.openingDate
    );


  if (
    monthToNumber(
      month
    ) <
    monthToNumber(
      openingMonth
    )
  ) {

    return {

      beforeOpening: true,

      nilRoyalty: true,

      amount: 0,

      reason:
        'BEFORE OPENING'

    };

  }


  if (
    String(
      outlet.type
    ).toLowerCase() ===
    'own'
  ) {

    return {

      beforeOpening: false,

      nilRoyalty: true,

      amount: 0,

      reason:
        'OWN OUTLET'

    };

  }


  let start =
    outlet.royaltyStartDate
      ? monthFromDate(
          outlet.royaltyStartDate
        )
      : addMonths(
          openingMonth,
          3
        );


  if (
    monthToNumber(
      month
    ) <
    monthToNumber(
      start
    )
  ) {

    return {

      beforeOpening: false,

      nilRoyalty: true,

      amount: 0,

      reason:
        'NO FRANCHISE FEE'

    };

  }


  return {

    beforeOpening: false,

    nilRoyalty: false,

    amount:
      Number(
        outlet.monthlyRoyalty ||
        DEFAULT_ROYALTY
      ),

    reason:
      'ROYALTY APPLICABLE'

  };

}


/*************************************************
 DATE FUNCTIONS
**************************************************/

function getFourthMonthStart(
  openingDate
) {

  const month =
    monthFromDate(
      openingDate
    );


  return (
    addMonths(
      month,
      3
    ) +
    '-01'
  );

}


function monthFromDate(
  value
) {

  if (!value) {
    return '';
  }


  if (
    typeof value ===
      'string' &&
    /^\d{4}-\d{2}/.test(
      value
    )
  ) {

    return value.substring(
      0,
      7
    );

  }


  if (
    typeof value ===
      'string' &&
    /^\d{2}-\d{2}-\d{4}$/.test(
      value
    )
  ) {

    const p =
      value.split('-');

    return (
      p[2] +
      '-' +
      p[1]
    );

  }


  try {

    return Utilities.formatDate(
      new Date(value),
      Session.getScriptTimeZone(),
      'yyyy-MM'
    );

  }

  catch(e) {

    return '';

  }

}


function normalizeMonth(
  value
) {

  if (!value) {
    return '';
  }


  if (
    typeof value ===
      'string' &&
    /^\d{4}-\d{2}$/.test(
      value
    )
  ) {

    return value;

  }


  if (
    typeof value ===
      'string' &&
    /^\d{4}-\d{2}-\d{2}/.test(
      value
    )
  ) {

    return value.substring(
      0,
      7
    );

  }


  try {

    return Utilities.formatDate(
      new Date(value),
      Session.getScriptTimeZone(),
      'yyyy-MM'
    );

  }

  catch(e) {

    return '';

  }

}


function monthToNumber(
  month
) {

  if (!month) {
    return 0;
  }


  const p =
    month.split('-');


  return (
    Number(p[0]) * 12 +
    Number(p[1])
  );

}


function addMonths(
  month,
  count
) {

  const p =
    month.split('-');


  const d =
    new Date(
      Number(p[0]),
      Number(p[1]) - 1 + count,
      1
    );


  return (
    d.getFullYear() +
    '-' +
    String(
      d.getMonth() + 1
    ).padStart(
      2,
      '0'
    )
  );

}


function monthsBetween(
  start,
  end
) {

  const result = [];

  let current =
    start;

  let safety = 0;


  while (
    monthToNumber(
      current
    ) <=
    monthToNumber(
      end
    )
  ) {

    result.push(
      current
    );


    current =
      addMonths(
        current,
        1
      );


    safety++;


    if (
      safety > 300
    ) {

      break;

    }

  }


  return result;

}


function currentMonth() {

  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM'
  );

}


function displayMonth(
  month
) {

  const names = [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec'
  ];


  const p =
    month.split('-');


  return (
    names[
      Number(p[1]) - 1
    ] +
    ' ' +
    p[0]
  );

}


function formatDateDisplay(
  value
) {

  if (!value) {
    return '';
  }


  if (
    typeof value ===
      'string' &&
    /^\d{4}-\d{2}-\d{2}/.test(
      value
    )
  ) {

    const p =
      value
        .substring(0,10)
        .split('-');


    return (
      p[2] +
      '-' +
      p[1] +
      '-' +
      p[0]
    );

  }


  if (
    typeof value ===
      'string' &&
    /^\d{2}-\d{2}-\d{4}$/.test(
      value
    )
  ) {

    return value;

  }


  try {

    return Utilities.formatDate(
      new Date(value),
      Session.getScriptTimeZone(),
      'dd-MM-yyyy'
    );

  }

  catch(e) {

    return '';

  }

}


function formatDateTime(
  value
) {

  if (!value) {
    return '';
  }


  try {

    return Utilities.formatDate(
      new Date(value),
      Session.getScriptTimeZone(),
      'dd-MM-yyyy hh:mm a'
    );

  }

  catch(e) {

    return '';

  }

}

/************************************************************
 POTATOWALA GITHUB APP API ADD-ON
 PASTE THIS AT THE VERY END OF YOUR EXISTING Code.gs
 DO NOT DELETE YOUR EXISTING CODE.
************************************************************/

function pwApiHandle_(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || "");
  if (!action) return null;

  try {
    if (action === "loadAll") return pwJsonp_(pwLoadAll_(), p.callback);
    if (action === "saveSales") { pwSaveSales_(p); return pwJsonp_({ok:true}, p.callback); }
    if (action === "saveRoyalty") { pwSaveRoyalty_(p); return pwJsonp_({ok:true}, p.callback); }
    return pwJsonp_({ok:false,error:"Unknown action: "+action}, p.callback);
  } catch(err) {
    return pwJsonp_({ok:false,error:String(err)}, p.callback);
  }
}

function pwLoadAll_() {
  setupSheets();
  pwEnsureDataSheets_();

  const outlets = (getOutlets() || []).map(o => ({
    name:String(o.name || ""),
    type:String(o.type || ""),
    openingDate:String(o.openingDate || ""),
    royaltyStartDate:String(o.royaltyStartDate || ""),
    monthlyRoyalty:Number(o.monthlyRoyalty || 0),
    status:String(o.status || "Active")
  }));

  const ss = getSS();
  const salesSh = ss.getSheetByName("SALES");
  const royaltySh = ss.getSheetByName("ROYALTY");

  const sales=[];
  if (salesSh && salesSh.getLastRow() >= 2) {
    salesSh.getRange(2,1,salesSh.getLastRow()-1,10).getValues().forEach(r=>{
      if(!r[0]||!r[1]) return;
      sales.push({
        outlet:String(r[0]),month:pwMonth_(r[1]),
        petpoojaSales:Number(r[2]||0),swiggySales:Number(r[3]||0),
        zomatoSales:Number(r[4]||0),otherSales:Number(r[5]||0),
        totalSales:Number(r[6]||0),totalExpenses:Number(r[7]||0),
        profitLoss:Number(r[8]||0)
      });
    });
  }

  const royalty=[];
  if (royaltySh && royaltySh.getLastRow() >= 2) {
    royaltySh.getRange(2,1,royaltySh.getLastRow()-1,8).getValues().forEach(r=>{
      if(!r[0]||!r[1]) return;
      royalty.push({
        outlet:String(r[0]),month:pwMonth_(r[1]),amount:Number(r[2]||0),
        status:String(r[3]||"").toLowerCase(),billDate:pwDate_(r[4]),
        paymentDate:pwDate_(r[5]),remarks:String(r[6]||"")
      });
    });
  }
  return {ok:true,outlets:outlets,sales:sales,royalty:royalty};
}

function pwEnsureDataSheets_() {
  const ss=getSS();
  let s=ss.getSheetByName("SALES");
  if(!s) s=ss.insertSheet("SALES");
  if(s.getLastRow()===0) s.appendRow(["Outlet","Month","Petpooja Sales","Swiggy Sales","Zomato Sales","Other Sales","Total Sales","Total Expenses","Profit / Loss","Updated At"]);

  let r=ss.getSheetByName("ROYALTY");
  if(!r) r=ss.insertSheet("ROYALTY");
  if(r.getLastRow()===0) r.appendRow(["Outlet","Month","Amount","Status","Bill Date","Payment Date","Remarks / Ref No","Updated At"]);
}

function pwSaveSales_(p){
  pwEnsureDataSheets_();
  const sh=getSS().getSheetByName("SALES");
  const key=String(p.outlet||"").trim()+"|"+String(p.month||"");
  const row=[
    String(p.outlet||"").trim(),String(p.month||""),
    Number(p.petpoojaSales||0),Number(p.swiggySales||0),
    Number(p.zomatoSales||0),Number(p.otherSales||0),
    Number(p.totalSales||0),Number(p.totalExpenses||0),
    Number(p.profitLoss||0),new Date()
  ];
  pwUpsert_(sh,key,r=>String(r[0]).trim()+"|"+pwMonth_(r[1]),row);
}

function pwSaveRoyalty_(p){
  pwEnsureDataSheets_();
  const sh=getSS().getSheetByName("ROYALTY");
  const key=String(p.outlet||"").trim()+"|"+String(p.month||"");
  const row=[
    String(p.outlet||"").trim(),String(p.month||""),Number(p.amount||0),
    String(p.status||"").toLowerCase(),p.billDate||"",
    p.paymentDate||"",p.remarks||"",new Date()
  ];
  pwUpsert_(sh,key,r=>String(r[0]).trim()+"|"+pwMonth_(r[1]),row);
}

function pwUpsert_(sh,key,keyFn,row){
  const vals=sh.getDataRange().getValues();
  for(let i=1;i<vals.length;i++){
    if(keyFn(vals[i])===key){
      sh.getRange(i+1,1,1,row.length).setValues([row]); return;
    }
  }
  sh.appendRow(row);
}

function pwMonth_(v){
  if(!v) return "";
  if(Object.prototype.toString.call(v)==="[object Date]"&&!isNaN(v)){
    return Utilities.formatDate(v,Session.getScriptTimeZone()||"Asia/Kolkata","yyyy-MM");
  }
  const s=String(v).trim(),m=s.match(/^(\d{4})-(\d{2})/);
  return m ? m[1]+"-"+m[2] : s;
}

function pwDate_(v){
  if(!v) return "";
  if(Object.prototype.toString.call(v)==="[object Date]"&&!isNaN(v)){
    return Utilities.formatDate(v,Session.getScriptTimeZone()||"Asia/Kolkata","yyyy-MM-dd");
  }
  return String(v).trim();
}

function pwJsonp_(obj,callback){
  const text=callback ? callback+"("+JSON.stringify(obj)+");" : JSON.stringify(obj);
  return ContentService.createTextOutput(text).setMimeType(callback?ContentService.MimeType.JAVASCRIPT:ContentService.MimeType.JSON);
}
