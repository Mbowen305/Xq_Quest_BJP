/* Excel Quest · Financial Modeling world. Built by integrate_fm.py: content, grader, and app module. */
/* ===================== DATA ===================== */
const GATE = [
  {world:'World 1 · Associate Isles', track:'MOS', ids:['1-1','1-2','1-3','1-4','1-5']},
  {world:'World 2 · Expert Fortress', track:'MOS', ids:['2-1','2-2','2-3','2-4']},
  {world:'Workbook Wharf', track:'TOSA', ids:['E-1','E-2','E-3','E-4']},
  {world:'Functions Forest', track:'TOSA', ids:['F-1','F-2','F-3','F-4','F-5','F-6']},
  {world:'Style Summit', track:'TOSA', ids:['M-1','M-2','M-3','M-4']},
  {world:'Data Delta', track:'TOSA', ids:['D-1','D-2','D-3','D-4']}
];
const GATE_IDS = GATE.flatMap(g=>g.ids);

const DECKS = [
  ['00a','Formula Grammar, Part 1','Foundations'],['00b','Formula Grammar, Part 2','Foundations'],
  ['10','SUM, AVERAGE, MAX, MIN','Math and Counting'],['11','Counting','Math and Counting'],['12','Rounding','Math and Counting'],
  ['20','Logic: IF, AND, OR, IFS, SWITCH','Logic'],['21','Error Traps','Logic'],['22','Conditional Math','Logic'],['23','SUMPRODUCT and SUBTOTAL','Logic'],['24','Database Functions','Logic'],
  ['30','VLOOKUP and HLOOKUP','Lookups'],['31','INDEX and MATCH','Lookups'],['32','XLOOKUP','Lookups'],['33','OFFSET','Lookups'],
  ['40','Cutting Text','Text'],['41','Cleaning Text','Text'],['42','Joining Text','Text'],
  ['50','Dates','Dates'],['51','Date Math','Dates'],
  ['60','Dynamic Arrays','Arrays'],['61','Random Numbers','Arrays'],
  ['70','Finance Functions','Finance'],
  ['80','LET and LAMBDA','LAMBDA'],['81','LAMBDA Helpers','LAMBDA'],['82','REGEX','LAMBDA'],
  ['90','Names and Tables','Names and Tables']
];
const FM_DECKS = [
  ['FM-0','The Three Statements'],['FM-1','Forecast Drivers'],['FM-2','Linking the Model'],
  ['FM-3','Comparable Companies'],['FM-4','Discounted Cash Flow'],['FM-B','Sensitivity and the Football Field']
];

/* Workbook map: shared by every FM level file */
const FM_MAP = [
  ['Mission','The level goal and the quest list, same as every Excel Quest level.'],
  ['Drivers','Every assumption, in blue. Years run across: B to D are FY2023 to FY2025 (actual), E to I are FY2026 to FY2030 (forecast).'],
  ['IS','Income statement. Rows: 4 Revenue, 5 Growth, 6 COGS, 7 Gross profit, 8 Gross margin, 9 SG&A, 10 D&A, 11 EBIT, 12 Interest, 13 Pre-tax income, 14 Taxes, 15 Net income, 16 EBITDA.'],
  ['BS','Balance sheet. Rows: 4 Cash, 5 Receivables, 6 Inventory, 7 PP&E, 8 Total assets, 10 Payables, 11 Loan, 12 Total liabilities, 14 Paid-in capital, 15 Retained earnings, 16 Total equity, 17 Liabilities + equity, 19 Check, 20 Status.'],
  ['CF','Cash flow statement. Rows: 4 Net income, 5 D&A, 6 to 8 working capital, 9 Operations, 11 Capex, 12 Investing, 13 Borrowing, 14 Repayment, 15 Owner draws, 16 Financing, 18 Net change, 19 Beginning cash, 20 Ending cash.'],
  ['Comps','Peer companies and their multiples (FM-3 on).'],
  ['DCF','Free cash flow, discounting, terminal value, and the sensitivity table (FM-4 on).'],
  ['Summary','The investor offer, the football field chart, and the recommendation (Boss level).']
];

const FM_LEVELS = [
{ id:'FM-0', name:'Ledger Gate', topic:'Read the three statements',
  story:'Two years after the expansion loan, Magic City Sneaker Exchange is busy. A local sneaker brand wants to buy 20% of the shop. Before anyone talks price, you need to read the books the way an analyst does.',
  goal:'A checked FY2025 statement pack: income statement subtotals, a rebuilt cash flow statement whose ending cash matches the balance sheet, retained earnings rolled forward, and a balance check that reads OK.',
  items:['SUM','Cross-sheet references','ABS','IF','Custom number formats','Sign conventions'],
  quests:[
    {t:'Income statement subtotals', task:'On IS, fill B7:D7 (Gross profit), B11:D11 (EBIT), B13:D13 (Pre-tax income), and B15:D15 (Net income) for FY2023 to FY2025.',
     steps:['In B7 type `=B4-B6` (Revenue minus COGS).','In B11 type `=B7-B9-B10` (Gross profit minus SG&A minus D&A).','In B13 type `=B11-B12`, and in B15 type `=B13-B14`.','Select B7:D7 and press {{Cmd+R}}. Do the same for B11:D11, B13:D13, and B15:D15. Always start the selection on the formula cell: {{Cmd+R}} copies the leftmost column of the selection to the right. Select only column B and it copies column A into B instead. Select the whole block B7:D15 and it copies the typed numbers in column B over C and D.'],
     answer:'FY2025: Gross profit 365, EBIT 144, Pre-tax income 132, Net income 99 ($000).',
     tip:'Every number in this model is in thousands. 365 means $365,000.'},
    {t:'Margins and EBITDA', task:'Fill IS!B8:D8 with gross margin and IS!B16:D16 with EBITDA (EBIT plus D&A). Format row 8 as a percent with one decimal.',
     steps:['In B8 type `=B7/B4`.','In B16 type `=B11+B10`.','Select B8:D8 and press {{Cmd+R}}, then select B16:D16 and press {{Cmd+R}}. Format B8:D8 with {{Cmd+1}} > Percentage, 1 decimal.'],
     answer:'Gross margin 30.0%, 30.9%, 32.0%. EBITDA 84, 114, 162.',
     why:'Investors compare businesses by margin, not dollars. EBITDA is the number the comps level multiplies.'},
    {t:'Rebuild the FY2025 cash flow', task:'On CF, fill column D using the Notes sheet: net income, D&A, the three working-capital changes, capex, new borrowing, repayment, and owner draws. Then total each section and the ending cash.',
     steps:['D4 `=IS!D15` and D5 `=IS!D10`.','Working capital: D6 `=-(BS!D5-BS!C5)`, D7 `=-(BS!D6-BS!C6)`, D8 `=BS!D10-BS!C10`.','D9 `=SUM(D4:D8)` for cash from operations.','D11 `=-Notes!B4` (capex). D12 `=D11`.','D13 `=Notes!B5`, D14 `=-Notes!B6`, D15 `=-Notes!B7`, and D16 `=SUM(D13:D15)`.','D18 `=D9+D12+D16`, D19 `=BS!C4`, D20 `=D19+D18`.'],
     answer:'Operations 56, investing (128), financing 100, net change 28, ending cash 60. It matches BS!D4.',
     trap:'An increase in inventory is a use of cash, so it goes in as a negative. A rise in payables is cash you kept, so it goes in as a positive.'},
    {t:'Roll retained earnings forward', task:'In BS!D15, calculate FY2025 retained earnings from last year\'s balance, this year\'s net income, and owner draws.',
     steps:['In D15 type `=C15+IS!D15+CF!D15`. Draws are already negative on CF, so you add them.'],
     answer:'81 + 99 - 30 = 150.',
     why:'This is the main bridge between the income statement and the balance sheet.'},
    {t:'Totals and the balance check', task:'Fill the BS totals for FY2024 and FY2025 (C8:D8, C12:D12, C16:D16, C17:D17). In D19 subtract liabilities + equity from total assets. In D20 show OK when the check is zero and CHECK when it is not.',
     steps:['C8 `=SUM(C4:C7)`, C12 `=SUM(C10:C11)`, C16 `=SUM(C14:C15)`, C17 `=C12+C16`. Then select C8:D8 and press {{Cmd+R}}. Do the same for rows 12, 16, and 17.','D19 `=D8-D17`.','D20 `=IF(ABS(D19)<0.01,"OK","CHECK")`.'],
     answer:'FY2025: total assets 400, liabilities 170, equity 230, liabilities + equity 400, check 0, status OK. FY2024: 195, 34, 161, 195. If FY2025 equity is not 230, check BS!D15 first (it should be 150), then IS!D15 net income (99).',
     tip:'ABS with a tiny tolerance stops rounding dust like 0.0000001 from tripping the check.'},
    {t:'Sort the accounts', task:'On the Sort It sheet, type IS, BS, or CF in column B next to each item in column A.',
     steps:['Ask: is it earned or spent this year (IS), owned or owed on one day (BS), or a movement of cash that isn\'t revenue or expense (CF)?'],
     answer:'Revenue IS, Rent IS, Cash BS, Inventory BS, Accounts payable BS, Loan balance BS, Buying display cases CF, Paying back part of the loan CF.'}
  ],
  quiz:[
    {q:'Which number connects the income statement to the balance sheet every year?', c:['Revenue becomes cash','Net income flows into retained earnings','COGS becomes inventory','Taxes become payables'], a:1, why:'Retained earnings = last year\'s retained earnings + net income - owner draws. That is the bridge.'},
    {q:'The shop spent $128K on display cases in FY2025. Where does that show up?', c:['As an expense on the income statement','Only on the balance sheet','In investing cash flow and in PP&E, then slowly on the income statement through depreciation','Nowhere, because the loan paid for it'], a:2, why:'Capex is cash out the door and an asset on the balance sheet. Depreciation spreads its cost over the years it is used.'}
  ]},
{ id:'FM-1', name:'Forecast Foundry', topic:'Forecasting with drivers',
  story:'The investor asks one question: what does this shop look like in five years? A forecast is only as good as the assumptions under it, so build them where anyone can see and test them.',
  goal:'A Drivers sheet with historical ratios, a Bear/Base/Bull switch, a five-year revenue forecast, a trend-line cross-check, and every typed input in blue.',
  items:['Ratios','Data Validation','XLOOKUP','FORECAST.LINEAR','Mixed references','Font color'],
  quests:[
    {t:'Historical growth', task:'On IS, fill C5:D5 with revenue growth for FY2024 and FY2025.',
     steps:['In C5 type `=C4/B4-1`. Select C5:D5 and press {{Cmd+R}}. Format as a percent.'],
     answer:'15.9% and 20.0%.'},
    {t:'Historical ratios', task:'On Drivers, fill B5:D7 with gross margin, SG&A % of revenue, and D&A % of revenue. Fill C9:D11 with receivable days, inventory days, and payable days. Fill D8 with FY2025 capex % of revenue.',
     steps:['B5 `=IS!B7/IS!B4`, B6 `=IS!B9/IS!B4`, B7 `=IS!B10/IS!B4`. Then select B5:D7 and press {{Cmd+R}}.','C9 `=BS!C5/IS!C4*365`. C10 `=BS!C6/IS!C6*365`. C11 `=BS!C10/IS!C6*365`. Then select C9:D11 and press {{Cmd+R}}.','D8 `=Notes!B4/IS!D4`.'],
     answer:'FY2025: margin 32.0%, SG&A 17.8%, D&A 1.6%, AR 3.2 days, inventory 89.5 days, AP 18.8 days, capex 11.2%.',
     trap:'Capex was 11.2% in FY2025 because of the one-time store build-out. Don\'t copy a one-time year into the forecast. The Base case uses 2%.',
     why:'Days turn a balance into a habit: the shop holds about 90 days of shoes. Habits forecast better than dollar amounts.'},
    {t:'Scenario switch', task:'Make Drivers!B18 a dropdown with Bear, Base, Bull. Then fill E4:I4 with the growth row for whichever scenario B18 names, using the table in A20:I22.',
     steps:['Click B18 and choose [[Data > Data Validation]]. Allow: List. Source: `Bear,Base,Bull`. Pick Base.','In E4 type `=XLOOKUP($B$18,$A$20:$A$22,E20:E22)`. Select E4:I4 and press {{Cmd+R}}.'],
     answer:'With Base selected: 16%, 13%, 10%, 7%, 5%.',
     tip:'`=CHOOSE(2,E20,E21,E22)` does the same job with a number instead of a name. Names are easier to read in a meeting.'},
    {t:'Forecast revenue', task:'Fill IS!E4:I4 with revenue for FY2026 to FY2030, growing each year by that year\'s growth driver. Then fill E5:I5 with the growth check.',
     steps:['In E4 type `=D4*(1+Drivers!E4)`. Select E4:I4 and press {{Cmd+R}}.','Select C5:I5 (start on C5, which already has the growth formula) and press {{Cmd+R}} so the growth row checks itself.'],
     answer:'Base case: 1,322.4, 1,494.3, 1,643.7, 1,758.8, 1,846.7.',
     trap:'Typing `=D4*1.16` buries the assumption inside the formula. Nobody can see it or change it. Point at the driver instead.'},
    {t:'Trend-line cross-check', task:'In Drivers!E24:I24, use FORECAST.LINEAR to project revenue from the FY2023 to FY2025 trend. Compare it to your forecast.',
     steps:['Row 3 holds the years as numbers (they only look like FY2026E). In E24 type `=FORECAST.LINEAR(E$3,IS!$B$4:$D$4,$B$3:$D$3)`. Select E24:I24 and press {{Cmd+R}}.'],
     answer:'1,290 in FY2026 and 1,930 in FY2030. A straight line adds the same 160 every year, so it ends up above the Base case, which slows down.',
     why:'Two methods that disagree start a useful argument: is the shop\'s growth slowing or not?'},
    {t:'Color the inputs', task:'Make every typed assumption on Drivers blue (E5:I15, E20:I22, and B18). Leave formulas black. The statements already follow this rule: look at the blue history on IS.',
     steps:['Click E5 and drag to I15, so every row from 5 to 15 is selected (not just rows 5 and 15). Then hold {{Cmd}} and drag over E20:I22, and Cmd-click B18.','Click the arrow next to [[Home > Font Color]] and pick Blue under Standard Colors.'],
     why:'The modeling rule everywhere: blue is something a person typed, black is a formula. Anyone opening the file knows what they are allowed to change.'}
  ],
  quiz:[
    {q:'FY2025 capex was 11.2% of revenue. Why does the forecast use 2%?', c:['2% is the legal limit','FY2025 included a one-time store build-out','Capex always shrinks','The investor asked for it'], a:1, why:'A forecast should reflect a normal year. One-time spending gets called out, not copied.'},
    {q:'A cell holds `=D4*(1+0.16)`. What is the problem?', c:['It should use SUM','The 16% is hidden inside the formula, so no one can see or test it','Growth can\'t be written as a decimal','Nothing'], a:1, why:'Put 16% in a blue input cell and point at it. Then every scenario is one change away.'}
  ]},
{ id:'FM-2', name:'The Linkworks', topic:'The three-statement model',
  story:'Revenue doesn\'t pay investors. Cash does. Link the income statement, balance sheet, and cash flow so that one change in a blue cell ripples through all three.',
  goal:'A five-year, three-statement model where the balance check reads OK in every year and changing any input updates everything.',
  items:['Cross-sheet links','MAX','IF / ABS','Fill right','Sign conventions','Trace Precedents'],
  quests:[
    {t:'Forecast the income statement', task:'Fill IS!E6:I11: COGS, gross profit, gross margin, SG&A, D&A, and EBIT, using the Drivers sheet.',
     steps:['E6 `=E4*(1-Drivers!E5)`, E7 `=E4-E6`, E8 `=E7/E4`.','E9 `=E4*Drivers!E6`, E10 `=E4*Drivers!E7`, E11 `=E7-E9-E10`.','Select E6:I11 and press {{Cmd+R}}. Start on column E, the one with formulas. Selecting only E6:E11 would copy column D into E.'],
     answer:'FY2026 EBIT 164.0. FY2030 EBIT 229.0.'},
    {t:'Interest, taxes, net income', task:'Fill IS!E12:I16. Interest uses LAST year\'s loan balance. Taxes can\'t go below zero.',
     steps:['E12 `=Drivers!E13*BS!D11`.','E13 `=E11-E12`, E14 `=MAX(0,E13)*Drivers!E12`, E15 `=E13-E14`, E16 `=E11+E10`.','Select E12:I16 and press {{Cmd+R}}.'],
     answer:'FY2026: interest 10.4, net income 115.2.',
     trap:'Using this year\'s loan balance makes a loop: interest changes net income, which changes cash, which changes the loan, which changes interest. Excel calls it a circular reference. Beginning balance avoids it.'},
    {t:'Working capital', task:'Fill BS!E5:I6 and BS!E10:I10 from the day drivers.',
     steps:['E5 `=IS!E4*Drivers!E9/365` (receivables).','E6 `=IS!E6*Drivers!E10/365` (inventory runs on COGS, not revenue).','E10 `=IS!E6*Drivers!E11/365` (payables). Then select E5:I6 and press {{Cmd+R}}, and select E10:I10 and press {{Cmd+R}}.'],
     answer:'FY2026: receivables 10.9, inventory 221.7, payables 46.8.'},
    {t:'PP&E, the loan, and equity', task:'Fill BS!E7:I7, E11:I11, E14:I14, and E15:I15.',
     steps:['E7 `=D7+IS!E4*Drivers!E8-IS!E10` (last year + capex - depreciation).','E11 `=MAX(0,D11-Drivers!E14)` (the loan never goes below zero).','E14 `=D14`. E15 `=D15+IS!E15+CF!E15`.','Select E7:I7 and press {{Cmd+R}}. Do the same for E11:I11, E14:I14, and E15:I15. {{Cmd+R}} only works on one block at a time.'],
     answer:'FY2026: PP&E 145.3, loan 110.0, retained earnings 235.2.'},
    {t:'The cash flow statement', task:'Fill CF!E4:I20 for the forecast years, then link BS!E4:I4 to ending cash.',
     steps:['Follow the FY2025 pattern in column D: E4 `=IS!E15`, E5 `=IS!E10`, E6 to E8 use the BS changes, E9 `=SUM(E4:E8)`.','E11 `=-IS!E4*Drivers!E8`, E12 `=E11`.','E13 `=0`, E14 `=BS!E11-BS!D11`, E15 `=-Drivers!E15`, E16 `=SUM(E13:E15)`.','E18 `=E9+E12+E16`, E19 `=D20`, E20 `=E19+E18`.','On BS, E4 `=CF!E20`. Then select CF!E4:I20 and press {{Cmd+R}}, and select BS!E4:I4 and press {{Cmd+R}}.'],
     answer:'FY2026: operations 110.6, investing (26.4), financing (50.0), ending cash 94.1. FY2030 ending cash 401.4.'},
    {t:'Balance every year', task:'Fill BS totals, the check row (19), and the status row (20) for E:I. Every status must read OK.',
     steps:['Select C8:I8 and press {{Cmd+R}}. Do the same for rows 12, 16, and 17. Then select D19:I20 and press {{Cmd+R}}.','If any year says CHECK, use [[Formulas > Trace Precedents]] on that check cell and follow the arrows.'],
     answer:'Check row 0.0 in all five years.',
     tip:'Most broken models fail the check by an amount that names the bug. Off by exactly the D&A number? You forgot to add it back on CF.'},
    {t:'Break it on purpose', note:true, task:'Switch Drivers!B18 to Bear, then Bull. Set owner draws to 150 in FY2027. Watch what happens to cash. Undo when you\'re done.',
     steps:['Bear ends FY2030 with 342 of cash, Bull with 457. Draws of 150 push cash down fast.','If cash goes below zero, a real business would draw on a credit line. That\'s beyond this model.'],
     why:'A model is a tool for asking what if. If changing one blue cell doesn\'t ripple through, something isn\'t linked.'}
  ],
  quiz:[
    {q:'Inventory rises from 190.0 to 221.7. On the cash flow statement this shows up as:', c:['+31.7','-31.7','0, because inventory is not cash','-221.7'], a:1, why:'Shoes sitting on shelves are cash the shop spent and hasn\'t gotten back yet. A rise in inventory uses cash.'},
    {q:'Why is interest calculated on the beginning loan balance?', c:['Banks require it','It avoids a circular reference','It makes interest larger','The ending balance isn\'t known until next year'], a:1, why:'Interest feeds net income, cash, and the loan. Pointing it at last year\'s loan breaks the loop.'}
  ]},
{ id:'FM-3', name:'Comps Colosseum', topic:'Comparable company analysis',
  story:'The investor says sneaker businesses sell for 10 times EBITDA. Maybe. Check the claim against a set of peer companies and see where that number came from.',
  goal:'A comps table with multiples, a cleaned peer set, median and quartile multiples, and an implied value range for the shop.',
  items:['Division','FILTER','MEDIAN','QUARTILE.INC','Absolute references','Units'],
  quests:[
    {t:'Calculate the multiples', task:'On Comps, fill F5:F10 with EV / Revenue and G5:G10 with EV / EBITDA.',
     steps:['F5 `=C5/D5`, G5 `=C5/E5`. Select F5:G10 and press {{Cmd+D}}.','Format both columns as `0.0"x"` with {{Cmd+1}} > Custom.'],
     answer:'SoleVault 1.5x / 12.0x, KickCulture 0.6x / 7.5x, Heatcheck 1.1x / 11.0x, Lace & Grail 0.9x / 9.0x, StreetStep 0.45x / 6.0x, Courtline 3.0x / 15.0x.'},
    {t:'Screen the peers', task:'Column H says whether each company belongs in the peer set. Change Courtline Athletic to N. Every other company stays Y.',
     steps:['Read column B. Courtline designs and manufactures its own shoes. Every other peer resells or retails them.','Type N in H10.'],
     why:'A brand that makes shoes earns different margins and grows differently. Including it inflates the multiples.'},
    {t:'Median and quartiles', task:'In B13:C15, calculate the 25th percentile, median, and 75th percentile of each multiple, using only peers marked Y.',
     steps:['B14 `=MEDIAN(FILTER(F$5:F$10,$H$5:$H$10="Y"))`.','B13 `=QUARTILE.INC(FILTER(F$5:F$10,$H$5:$H$10="Y"),1)` and B15 the same with 3.','Select B13:C15 and press {{Cmd+R}}.'],
     answer:'EV/Revenue 0.60x, 0.90x, 1.10x. EV/EBITDA 7.5x, 9.0x, 11.0x.',
     tip:'Include Courtline and the EBITDA median jumps to 10.0x. That\'s where the investor\'s number came from.'},
    {t:'Implied enterprise value', task:'In B18:C20, multiply each multiple by the shop\'s FY2025 revenue (IS!D4) or EBITDA (IS!D16).',
     steps:['B18 `=B13*IS!$D$4`. Select B18:B20 and press {{Cmd+D}}.','C18 `=C13*IS!$D$16`. Select C18:C20 and press {{Cmd+D}}.'],
     answer:'From revenue: 684, 1,026, 1,254. From EBITDA: 1,215, 1,458, 1,782 ($000).',
     trap:'The peers are in $ millions and the shop is in $ thousands. Multiples have no units, so they travel. Dollar amounts don\'t.'},
    {t:'From enterprise value to equity value', task:'In B27 calculate the shop\'s net debt (loan minus cash, FY2025). Fill B23:C25 with equity value: implied EV minus net debt.',
     steps:['B27 `=BS!D11-BS!D4`.','B23 `=B18-$B$27`. Select B23:B25 and press {{Cmd+D}}, then select B23:C25 and press {{Cmd+R}}.'],
     answer:'Net debt 70. Equity from revenue: 614, 956, 1,184. From EBITDA: 1,145, 1,388, 1,712.',
     why:'Enterprise value is the whole business. The owner\'s slice is what\'s left after paying off the debt, plus the cash in the drawer.'}
  ],
  quiz:[
    {q:'Why remove Courtline Athletic from the peer set?', c:['It is too big','It makes its own shoes, a different business model','Its name starts with C','Its EBITDA is negative'], a:1, why:'Peers should earn money the same way. Size can differ, the business model shouldn\'t.'},
    {q:'Why use the median multiple instead of the average?', c:['Excel can\'t average multiples','One extreme peer can drag the average, while the median stays in the middle','The median is always higher','Investors are required to'], a:1, why:'With five or six peers, one outlier moves the average a lot. The median ignores how extreme the outlier is.'}
  ]},
{ id:'FM-4', name:'DCF Depths', topic:'Discounted cash flow',
  story:'Comps tell you what other people pay. A DCF asks a different question: what is the shop\'s own future cash worth in today\'s dollars?',
  goal:'A DCF that turns five years of free cash flow and a terminal value into an enterprise value and equity value, cross-checked with NPV, plus a sensitivity table.',
  items:['Free cash flow','Exponents','NPV','Data Table','Absolute references','Percent formats'],
  quests:[
    {t:'Unlevered free cash flow', task:'On DCF, fill E4:I10: EBIT, taxes on EBIT, NOPAT, D&A, capex, change in working capital, and free cash flow.',
     steps:['E4 `=IS!E11`, E5 `=-E4*Drivers!E12`, E6 `=E4+E5`.','E7 `=IS!E10`, E8 `=CF!E11`, E9 `=SUM(CF!E6:E8)`.','E10 `=E6+E7+E8+E9`. Select E4:I10 and press {{Cmd+R}}.'],
     answer:'Free cash flow: 91.9, 108.8, 125.3, 140.4, 152.0.',
     why:'Unlevered means before interest. You value the whole business first, then deal with the loan.'},
    {t:'Discount each year', task:'Fill E12:I12 with year numbers 1 to 5, E13:I13 with discount factors, and E14:I14 with present values. The discount rate is in C16.',
     steps:['Type 1 to 5 in E12:I12.','E13 `=1/(1+$C$16)^E12`, E14 `=E10*E13`. Select E13:I14 and press {{Cmd+R}}.'],
     answer:'Present values at 14%: 80.6, 83.7, 84.6, 83.1, 78.9.',
     trap:'Forget the $ signs on C16 and the rate slides off to an empty cell when you fill right.'},
    {t:'Cross-check with NPV', task:'In C19 total the present values. In C18 get the same answer with one NPV formula.',
     steps:['C19 `=SUM(E14:I14)`.','C18 `=NPV(C16,E10:I10)`.'],
     answer:'Both 411.0.',
     trap:'NPV assumes the first cash flow arrives one year from now. Put a today amount inside the range and every year gets pushed back one, which shrinks the answer.'},
    {t:'Terminal value', task:'In C20 calculate the value of all cash flows after FY2030 using the growth formula. In C21 bring it back to today.',
     steps:['C20 `=I10*(1+C17)/(C16-C17)`.','C21 `=C20/(1+C16)^5`.'],
     answer:'Terminal value 1,423.3. Present value 739.2.',
     trap:'If the growth rate is equal to or higher than the discount rate, the formula breaks. A business can\'t outgrow the economy forever.'},
    {t:'Enterprise and equity value', task:'Fill C22 (enterprise value), C23 (net debt), C24 (equity value), and C25 (terminal value share of EV).',
     steps:['C22 `=C19+C21`, C23 `=BS!D11-BS!D4`, C24 `=C22-C23`, C25 `=C21/C22`.'],
     answer:'EV 1,150.2, net debt 70.0, equity 1,080.2, terminal share 64.3%.'},
    {t:'Sensitivity table', task:'Build a Data Table in B28:G33 that shows equity value for discount rates 12% to 16% (down B29:B33) and terminal growth 2% to 4% (across C28:G28).',
     steps:['B28 `=C24`. Type the rates in B29:B33 and the growth rates in C28:G28.','Select B28:G33 and choose [[Data > What-If Analysis > Data Table]].','Row input cell: C17. Column input cell: C16. Click OK.'],
     answer:'Base (14%, 3%) is 1,080. Corners: 848 (16%, 2%) to 1,485 (12%, 4%).',
     trap:'Data Table input cells must be on the same sheet as the table. That\'s why the discount rate and growth live on DCF, not Drivers.'}
  ],
  quiz:[
    {q:'The discount rate goes from 14% to 16%. The DCF value:', c:['Goes up','Goes down','Stays the same','Doubles'], a:1, why:'A higher required return means future cash is worth less today.'},
    {q:'About 64% of enterprise value comes from the terminal value. What does that tell you?', c:['The model is wrong','The answer leans heavily on what you assume after 2030, so test those inputs','Ignore the first five years','Terminal value should be removed'], a:1, why:'That share is normal for a growing business. It\'s why the sensitivity table matters.'}
  ]},
{ id:'FM-B', name:'The Investor\'s Offer', topic:'Boss level: make the call', boss:true,
  story:'The offer is on the table: $250,000 for 20% of the shop. Put every method on one chart, compare the offer, and make a recommendation the owner can defend.',
  goal:'A Summary sheet with the offer\'s implied valuation, a football field chart of every value range, and a recommendation that updates itself when any input changes.',
  items:['Pre- and post-money','MIN / MAX','Stacked bar chart','IF','Chart formatting'],
  quests:[
    {t:'Value the offer', task:'On Summary, B3 holds the investment (250) and B4 the stake (20%). Fill B5 with the post-money valuation and B6 with the pre-money valuation.',
     steps:['B5 `=B3/B4`. B6 `=B5-B3`.'],
     answer:'Post-money 1,250. Pre-money 1,000.',
     why:'The investor\'s cash goes into the company. Pre-money is what they think the shop is worth before their check.'},
    {t:'Football field table', task:'Fill A9:D11 with each method\'s low and high equity value and the spread (high minus low).',
     steps:['Row 9 EV/Revenue: B9 `=Comps!B23`, C9 `=Comps!B25`.','Row 10 EV/EBITDA: B10 `=Comps!C23`, C10 `=Comps!C25`.','Row 11 DCF: B11 `=MIN(DCF!C29:G33)`, C11 `=MAX(DCF!C29:G33)`.','D9 `=C9-B9`. Select D9:D11 and press {{Cmd+D}}.'],
     answer:'614 to 1,184. 1,145 to 1,712. 848 to 1,485.'},
    {t:'Draw the football field', task:'Make a stacked bar chart from A9:B11 and D9:D11. Hide the Low series so each bar floats from low to high. Title it with the offer.',
     steps:['Select A9:B11, hold {{Cmd}}, add D9:D11. Choose [[Insert > Bar > Stacked Bar]].','Right-click the Low bars, choose Format Data Series, and set Fill to No Fill.','Title: Equity value ($000) vs. offer at 1,000.'],
     tip:'If the methods come out upside down, open the vertical axis options and check Categories in reverse order.'},
    {t:'The recommendation', task:'In B14 write a formula that says Counter when the offer\'s pre-money is below the DCF base case, and Accept otherwise.',
     steps:['B14 `=IF(B6<DCF!C24,"Counter: the offer is below our base case","Accept: the offer meets our base case")`.'],
     answer:'Counter. The offer values the shop at 1,000; the base DCF says 1,080.'},
    {t:'Class Showdown', note:true, task:'In class: two teams, one Bear and one Bull. Each team gets three minutes to defend one assumption and its effect on the value. The class votes on a counteroffer.',
     steps:['Use the Model Lab or your own workbook live on screen.','Defend the logic behind an assumption. Nobody can prove a forecast is right.'],
     why:'Analysts don\'t win by having the right number. They win by having the best-defended one.'}
  ],
  quiz:[
    {q:'An investor offers $250K for 20%. What is the pre-money valuation?', c:['$1,250K','$1,000K','$250K','$200K'], a:1, why:'Post-money is 250 / 0.20 = 1,250. Subtract the 250 investment to get 1,000 pre-money.'},
    {q:'Comps say up to $1.7M and the DCF base says $1.08M. The best response is:', c:['Demand $1.7M','Accept whatever is offered','Show the ranges, explain why big public peers trade higher, and counter inside the overlap','Average every number'], a:2, why:'Small private shops usually sell for less than big public peers. A defended range beats a single number.'}
  ]}
];

const FM_CODEX = [
 {id:'cx-fm0', lvl:'FM-0', title:'The Three Statements', sub:'What each one shows and how they connect', blocks:[
  ['p','Every business tells its story three ways. The **income statement** covers a period: what the shop earned and spent this year. The **balance sheet** is a snapshot on one day: what the shop owns and owes. The **cash flow statement** explains why the cash balance changed between two snapshots.'],
  ['table',['Statement','Question it answers','Time'],[['Income statement','Did we make money?','A year'],['Balance sheet','What do we own and owe?','One day'],['Cash flow','Where did the cash go?','A year']]],
  ['h','Profit is not cash'],
  ['p','The shop can earn $99K of net income and still have less cash, because it spent money on inventory and equipment that isn\'t an expense yet. The cash flow statement starts at net income and adjusts for everything that moved cash differently.'],
  ['h','The three links'],
  ['list',['Net income flows into retained earnings on the balance sheet.','Net income is the first line of the cash flow statement.','Ending cash on the cash flow statement becomes cash on the balance sheet.']],
  ['ex','=C15+IS!D15+CF!D15','Retained earnings: last year, plus net income, plus owner draws (already negative).'],
  ['tip','Assets = Liabilities + Equity, every day, no exceptions. A check row that subtracts one side from the other should always show 0.'],
  ['trap','Depreciation is an expense on the income statement but not a cash payment. Forget to add it back on the cash flow statement and cash comes out too low by exactly that amount.']]},
 {id:'cx-fm1', lvl:'FM-1', title:'Forecast Drivers', sub:'Ratios, days, scenarios, and trend lines', blocks:[
  ['p','A driver is a ratio that behaves more steadily than dollars do. Instead of guessing next year\'s rent, you assume SG&A stays about 18% of revenue. Instead of guessing inventory, you assume the shop holds about 90 days of shoes.'],
  ['table',['Driver','Formula','Why it works'],[['Growth','This year / last year - 1','Revenue is the top line everything hangs from'],['Margin','Gross profit / revenue','Pricing power changes slowly'],['% of revenue','Expense / revenue','Costs scale with sales'],['Days','Balance / annual flow x 365','Habits, like how long shoes sit']]],
  ['ex','=XLOOKUP($B$18,$A$20:$A$22,E20:E22)','Pick the growth row for whichever scenario B18 names.'],
  ['h','Two ways to forecast'],
  ['p','**Driver-based:** revenue grows by an assumed rate each year. **Trend-based:** FORECAST.LINEAR fits a straight line through past years. When the two disagree, that disagreement is the conversation to have.'],
  ['ex','=FORECAST.LINEAR(E$3,IS!$B$4:$D$4,$B$3:$D$3)','Project revenue for the year in E3 from the FY2023 to FY2025 line.'],
  ['tip','Blue font for typed inputs, black for formulas. Everyone in finance reads models this way.'],
  ['trap','Never copy a one-time year forward. FY2025 capex was a store build-out.']]},
 {id:'cx-fm2', lvl:'FM-2', title:'Linking the Model', sub:'Make one input ripple through all three statements', blocks:[
  ['p','A linked model has one rule: every number is either a blue input or a formula that points at one. Change growth and revenue moves, then costs, profit, cash, and the balance sheet, all on their own.'],
  ['h','Order of build'],
  ['list',['Income statement down to EBIT.','Interest on last year\'s loan, then taxes and net income.','Balance sheet, everything except cash.','Cash flow statement, from net income to ending cash.','Link cash back to the balance sheet and check that it balances.']],
  ['h','Working capital signs'],
  ['table',['Change','Cash effect','Why'],[['Inventory up','Negative','Cash is sitting on shelves'],['Receivables up','Negative','Customers owe you, you haven\'t collected'],['Payables up','Positive','You owe suppliers, you kept the cash for now']]],
  ['ex','=-(BS!E6-BS!D6)','Change in inventory, with the sign flipped for cash.'],
  ['trap','Interest on this year\'s loan creates a circular reference. Use last year\'s balance.'],
  ['tip','When the check fails, the size of the gap is a clue. Compare it to D&A, capex, and each working capital change.']]},
 {id:'cx-fm3', lvl:'FM-3', title:'Comparable Companies', sub:'Value a business by what similar ones are worth', blocks:[
  ['p','If similar companies are valued at about 9 times their EBITDA, a fair first guess for this shop is 9 times its EBITDA. That ratio is a **multiple**.'],
  ['table',['Term','Meaning'],[['Enterprise value (EV)','The whole business: equity plus debt, minus cash'],['EBITDA','Earnings before interest, taxes, depreciation, and amortization'],['EV / EBITDA','How many years of EBITDA a buyer pays'],['EV / Revenue','Useful when profits are small or uneven']]],
  ['h','Picking peers'],
  ['list',['Same business model (resale, retail, not manufacturing).','Similar customers and growth.','Size can differ. Multiples adjust for it, though smaller companies usually trade lower.']],
  ['ex','=MEDIAN(FILTER(G$5:G$10,$H$5:$H$10="Y"))','Median EV/EBITDA of only the peers marked Y.'],
  ['p','Equity value = enterprise value - debt + cash. That is the part an investor buys a piece of.'],
  ['trap','Peers in $ millions, the shop in $ thousands. Multiply the multiple, never the peer dollars.']]},
 {id:'cx-fm4', lvl:'FM-4', title:'Discounted Cash Flow', sub:'What future cash is worth today', blocks:[
  ['p','$100 a year from now is worth less than $100 today, because today\'s $100 could be invested. A DCF shrinks each future year\'s cash by the return an investor requires: the **discount rate**.'],
  ['ex','=E10/(1+$C$16)^E12','Present value: cash flow divided by (1 + rate) to the power of the year number.'],
  ['h','Free cash flow'],
  ['p','Unlevered free cash flow = EBIT after tax + D&A - capex - increase in working capital. It\'s the cash the business produces before anyone is paid back.'],
  ['h','Terminal value'],
  ['p','The shop doesn\'t close in 2030. The terminal value captures every year after, assuming cash grows at a slow, steady rate forever.'],
  ['ex','=I10*(1+C17)/(C16-C17)','Next year\'s cash flow divided by (rate - growth).'],
  ['trap','NPV in Excel assumes the first cash flow is one year away. Including a today amount in the range pushes every year back one and shrinks the result.'],
  ['tip','Small private businesses use higher discount rates (often 12% to 20%) than big public companies, because they are riskier and harder to sell.']]},
 {id:'cx-fm5', lvl:'FM-B', title:'Sensitivity and the Football Field', sub:'Show a range, then make a call', blocks:[
  ['p','No one knows the right discount rate or growth rate. A **Data Table** recalculates the model for every combination so you can see how much the answer moves.'],
  ['step',['Put the output formula in the corner cell.','Type the rates down the side and the growth values across the top.','Select the block, then [[Data > What-If Analysis > Data Table]].','Row input = growth cell. Column input = rate cell.']],
  ['p','A **football field** chart lines up every method\'s range as floating bars. Where the bars overlap is the most defensible value.'],
  ['h','Pre-money and post-money'],
  ['ex','=B3/B4','Post-money: investment divided by stake. Subtract the investment to get pre-money.'],
  ['tip','Present the range first and your number second. It shows you did the work.']]}
];

const FM_CODES = [
  ['{{Cmd+1}}','Format Cells. Custom `#,##0.0;(#,##0.0)` shows negatives in parentheses, like a real statement.','FM-0'],
  ['{{Cmd+R}}','Fill right. Select the formula cell plus the cells to its right, then press. It copies the leftmost column of the selection, so always start on the formula. Fill down ({{Cmd+D}}) works the same way from the top row.','FM-1, FM-2'],
  ['{{Cmd+T}}','While editing, cycles a reference through $A$1, A$1, $A1, A1.','FM-1, FM-4'],
  ['{{Ctrl+[}}','Jump to the cells this formula reads. The fastest way to chase a broken check.','FM-2'],
  ['{{Ctrl+`}}','Show formulas instead of values. Spot any number typed where a formula should be.','FM-2'],
  ['{{Fn+F9}}','Highlight part of a formula while editing to see what it equals. Press {{Esc}} after.','FM-2, FM-4'],
  ['{{Ctrl+Cmd+V}}','Paste Special. Paste values to freeze a scenario you want to keep.','FM-B'],
  ['{{Cmd+=}}','Calculate now. Useful after a big Data Table.','FM-4'],
  ['Blue / black','Blue font is a typed input. Black font is a formula. Never mix them in one row.','All'],
  ['One row, one formula','Every cell across a forecast row uses the same formula. If one year is different, mark it.','All'],
  ['Check row = 0','Assets minus liabilities and equity. Anything but 0 means a broken link.','FM-0, FM-2']
];


const FM_FILES = {"FM-0": "FM-0_Sneaker_Exchange_Statements.xlsx", "FM-1": "FM-1_Sneaker_Exchange_Forecast.xlsx", "FM-2": "FM-2_Sneaker_Exchange_Model.xlsx", "FM-3": "FM-3_Sneaker_Exchange_Comps.xlsx", "FM-4": "FM-4_Sneaker_Exchange_DCF.xlsx", "FM-B": "FM-B_Sneaker_Exchange_Offer.xlsx"};
/* Levels built but still in teacher playtest: only the teacher preview (#fm-teacher + passcode) shows them. Remove an id to open it to students. */
const FM_TEACHER_ONLY = ["FM-3", "FM-4", "FM-B"];
/* Excel Quest Financial Modeling checks (FM-0 to FM-B). Registered into XQ. Arrays are in quest order (fm-proto/data.js).
   Expected values are recomputed from each file's own inputs, so a student who changes a blue cell still passes. */
(function(){
  const {cell,val,txt,norm,P,F} = XQC;
  const num = v => typeof v==='number' ? v : (v===null||v===''||v===undefined||isNaN(+v) ? NaN : +v);
  const isF = (S,ref) => !!cell(S,ref).f;
  const fx = (S,ref) => cell(S,ref).f ? '='+cell(S,ref).f : (cell(S,ref).v===null||cell(S,ref).v===undefined ? 'nothing' : 'the typed value '+cell(S,ref).v);
  const close = (a,b) => typeof a==='number' && isFinite(b) && Math.abs(a-b) <= Math.max(0.05, Math.abs(b)*0.0005);
  const k1 = n => (Math.round(n*10)/10).toLocaleString('en-US',{minimumFractionDigits:1,maximumFractionDigits:1});
  const pct = n => (n*100).toFixed(1)+'%';
  const prev = c => String.fromCharCode(c.charCodeAt(0)-1);
  const HIST = ['B','C','D'], FC = ['E','F','G','H','I'];
  const isPct = S_ref => /%/.test((S_ref.xf && S_ref.xf.fmt) || '');
  const rgb = h => [parseInt(h.slice(0,2),16), parseInt(h.slice(2,4),16), parseInt(h.slice(4,6),16)];
  const isBlue = c => { const h=(c && c.xf && c.xf.font && c.xf.font.color) || ''; if(h.length!==6) return false; const [r,g,b]=rgb(h); return b>=128 && b>r+60 && b>g+40; };
  /* one expectation: [sheetName, ref, expected, hint]; formula required unless noF */
  function expect(W, list, noF){
    for(const [sn, ref, want, hint, fmtPct] of list){
      const S = W.sheets[sn], c = cell(S,ref);
      if(!noF && !c.f) return F(c.v===null||c.v===undefined||c.v==='' ? `${sn}!${ref} is empty. ${hint}` : `${sn}!${ref} holds a typed number. Make it a formula so it updates: ${hint}`);
      if(typeof want==='string'){ if(txt(S,ref).toUpperCase()!==want.toUpperCase()) return F(`${sn}!${ref} shows ${txt(S,ref)||'nothing'}, not ${want}. ${hint}`); continue; }
      if(!close(num(c.v), want)) return F(`${sn}!${ref} shows ${typeof c.v==='number'?k1(c.v):(c.v||'nothing')}. It should be ${Math.abs(want)<2?pct(want):k1(want)}. ${hint}`);
      if(fmtPct && !isPct(c)) return F(`${sn}!${ref} is right but shows as a decimal. Format it as a percent: Cmd+1, Percentage.`);
    }
    return P();
  }
  const V = (W,sn,ref) => num(val(W.sheets[sn],ref));

  /* ---------- history (FM-0) ---------- */
  function hist(W){
    const o = {};
    HIST.forEach(c=>{ const rev=V(W,'IS',c+'4'), cogs=V(W,'IS',c+'6'), gp=rev-cogs, ebit=gp-V(W,'IS',c+'9')-V(W,'IS',c+'10'), ebt=ebit-V(W,'IS',c+'12'), ni=ebt-V(W,'IS',c+'14');
      o[c] = {rev,cogs,gp,gm:gp/rev,ebit,ebt,ni,ebitda:ebit+V(W,'IS',c+'10'),sga:V(W,'IS',c+'9'),da:V(W,'IS',c+'10')}; });
    const n = r => V(W,'Notes','B'+r);
    const cf = {ni:o.D.ni, da:o.D.da, dar:-(V(W,'BS','D5')-V(W,'BS','C5')), dinv:-(V(W,'BS','D6')-V(W,'BS','C6')), dap:V(W,'BS','D10')-V(W,'BS','C10'),
      capex:-n(4), borrow:n(5), repay:-n(6), draws:-n(7)};
    cf.cfo = cf.ni+cf.da+cf.dar+cf.dinv+cf.dap; cf.cff = cf.borrow+cf.repay+cf.draws; cf.net = cf.cfo+cf.capex+cf.cff; cf.beg = V(W,'BS','C4'); cf.end = cf.beg+cf.net;
    o.cf = cf; o.reD = V(W,'BS','C15') + cf.ni + cf.draws;
    return o;
  }
  const L0 = [
    W => { const h=hist(W); const l=[]; HIST.forEach(c=>{ l.push(['IS',c+'7',h[c].gp,`Revenue minus COGS: =${c}4-${c}6.`], ['IS',c+'11',h[c].ebit,`Gross profit minus SG&A minus D&A: =${c}7-${c}9-${c}10.`],
        ['IS',c+'13',h[c].ebt,`EBIT minus interest: =${c}11-${c}12.`], ['IS',c+'15',h[c].ni,`Pre-tax income minus taxes: =${c}13-${c}14.`]); }); return expect(W,l); },
    W => { const h=hist(W); const l=[]; HIST.forEach(c=>l.push(['IS',c+'8',h[c].gm,`Gross profit divided by revenue: =${c}7/${c}4.`,true])); HIST.forEach(c=>l.push(['IS',c+'16',h[c].ebitda,`EBIT plus D&A: =${c}11+${c}10.`])); return expect(W,l); },
    W => { const x=hist(W).cf; return expect(W,[
      ['CF','D4',x.ni,'Link it to net income: =IS!D15.'], ['CF','D5',x.da,'Add back depreciation: =IS!D10.'],
      ['CF','D6',x.dar,'More receivables uses cash, so flip the sign: =-(BS!D5-BS!C5).'], ['CF','D7',x.dinv,'More inventory uses cash: =-(BS!D6-BS!C6).'],
      ['CF','D8',x.dap,'More payables keeps cash: =BS!D10-BS!C10.'], ['CF','D9',x.cfo,'Total the operating rows: =SUM(D4:D8).'],
      ['CF','D11',x.capex,'Capex is cash going out: =-Notes!B4.'], ['CF','D12',x.capex,'Investing is just capex here: =D11.'],
      ['CF','D13',x.borrow,'Money borrowed comes in: =Notes!B5.'], ['CF','D14',x.repay,'Repaying goes out: =-Notes!B6.'], ['CF','D15',x.draws,'Owner draws go out: =-Notes!B7.'],
      ['CF','D16',x.cff,'Total financing: =SUM(D13:D15).'], ['CF','D18',x.net,'Operations + investing + financing: =D9+D12+D16.'],
      ['CF','D19',x.beg,'Start from last year\'s cash: =BS!C4.'], ['CF','D20',x.end,'Beginning cash plus the change: =D19+D18.']]); },
    W => { const h=hist(W); const r=expect(W,[['BS','D15',h.reD,'Last year plus net income plus draws (already negative on CF): =C15+IS!D15+CF!D15.']]);
      if(!r.pass) return r; return /IS!/.test(norm(cell(W.sheets.BS,'D15').f)) ? P() : F('BS!D15 has the right number but doesn\'t link to the income statement. Use =C15+IS!D15+CF!D15 so it updates.'); },
    W => { const B=W.sheets.BS; const l=[];
      ['C','D'].forEach(c=>{ const s=(a,b)=>{ let t=0; for(let r=a;r<=b;r++) t+=num(val(B,c+r))||0; return t; };
        const ta=s(4,7), tl=s(10,11), te=V(W,'BS',c+'15')+V(W,'BS',c+'14');
        l.push(['BS',c+'8',ta,`Total assets: =SUM(${c}4:${c}7).`], ['BS',c+'12',tl,`Total liabilities: =SUM(${c}10:${c}11).`], ['BS',c+'16',te,`Total equity: =SUM(${c}14:${c}15).`], ['BS',c+'17',tl+te,`Liabilities plus equity: =${c}12+${c}16.`]); });
      const r=expect(W,l); if(!r.pass) return r;
      const r2=expect(W,[['BS','D19',0,'Assets minus liabilities and equity: =D8-D17.']]); if(!r2.pass) return r2;
      if(!isF(B,'D20') || norm(cell(B,'D20').f).indexOf('IF(')<0) return F(`BS!D20 has ${fx(B,'D20')}. Use =IF(ABS(D19)<0.01,"OK","CHECK").`);
      return txt(B,'D20')==='OK' ? P() : F(`BS!D20 says ${txt(B,'D20')||'nothing'}. When the check is 0 it should say OK.`); },
    W => { const S=W.sheets['Sort It'], A=['IS','IS','BS','BS','BS','BS','CF','CF'];
      for(let i=0;i<8;i++){ const g=txt(S,'B'+(4+i)).toUpperCase(); if(!g) return F(`Sort It!B${4+i} is empty. Type IS, BS, or CF next to ${txt(S,'A'+(4+i))}.`);
        if(g!==A[i]) return F(`Look again at ${txt(S,'A'+(4+i))} (row ${4+i}). Is it earned or spent this year, owned or owed on one day, or a cash movement that isn't revenue or expense?`); }
      return P(); }
  ];

  /* ---------- drivers (FM-1) ---------- */
  function linFit(xs, ys, x){ const n=xs.length, mx=xs.reduce((a,b)=>a+b,0)/n, my=ys.reduce((a,b)=>a+b,0)/n; let sxy=0,sxx=0; for(let i=0;i<n;i++){ sxy+=(xs[i]-mx)*(ys[i]-my); sxx+=(xs[i]-mx)**2; } return my + sxy/sxx*(x-mx); }
  const SCEN_ROW = {BEAR:20, BASE:21, BULL:22};
  const L1 = [
    W => expect(W, ['C','D'].map(c=>['IS',c+'5',V(W,'IS',c+'4')/V(W,'IS',prev(c)+'4')-1,`This year divided by last year, minus 1: =${c}4/${prev(c)}4-1.`,true])),
    W => { const l=[]; HIST.forEach(c=>{ l.push(['Drivers',c+'5',V(W,'IS',c+'7')/V(W,'IS',c+'4'),`=IS!${c}7/IS!${c}4.`,true], ['Drivers',c+'6',V(W,'IS',c+'9')/V(W,'IS',c+'4'),`=IS!${c}9/IS!${c}4.`,true], ['Drivers',c+'7',V(W,'IS',c+'10')/V(W,'IS',c+'4'),`=IS!${c}10/IS!${c}4.`,true]); });
      ['C','D'].forEach(c=>{ l.push(['Drivers',c+'9',V(W,'BS',c+'5')/V(W,'IS',c+'4')*365,`=BS!${c}5/IS!${c}4*365.`], ['Drivers',c+'10',V(W,'BS',c+'6')/V(W,'IS',c+'6')*365,`Inventory runs on COGS: =BS!${c}6/IS!${c}6*365.`], ['Drivers',c+'11',V(W,'BS',c+'10')/V(W,'IS',c+'6')*365,`=BS!${c}10/IS!${c}6*365.`]); });
      l.push(['Drivers','D8',V(W,'Notes','B4')/V(W,'IS','D4'),'=Notes!B4/IS!D4.',true]); return expect(W,l); },
    W => { const D=W.sheets.Drivers; const dv=D.dv.find(d=>d.type==='list' && d.sqref.split(/\s+/).filter(Boolean).some(s=>XQC.cellsIn(s).includes('B18')));
      if(!dv) return F('Drivers!B18 has no dropdown. Choose Data, Data Validation, Allow: List, Source: Bear,Base,Bull.');
      const f1=dv.f1.toUpperCase(); if(!(/BEAR/.test(f1)&&/BASE/.test(f1)&&/BULL/.test(f1)) && !/A\$?20:\$?A\$?22/.test(f1)) return F('The dropdown on Drivers!B18 should list Bear, Base, and Bull.');
      const s=txt(D,'B18').toUpperCase(); if(!SCEN_ROW[s]) return F('Pick a scenario in Drivers!B18 from the dropdown.');
      if(s!=='BASE') return F(`Drivers!B18 is set to ${txt(D,'B18')}. Set it back to Base before you turn it in.`);
      for(const c of FC){ const f=norm(cell(D,c+'4').f); if(!f) return F(`Drivers!${c}4 is empty. =XLOOKUP($B$18,$A$20:$A$22,${c}20:${c}22)`);
        if(!/XLOOKUP\(|CHOOSE\(|INDEX\(/.test(f)) return F(`Drivers!${c}4 has ${fx(D,c+'4')}. Look up the scenario row: =XLOOKUP($B$18,$A$20:$A$22,${c}20:${c}22).`);
        if(!close(num(val(D,c+'4')), V(W,'Drivers',c+SCEN_ROW[s]))) return F(`Drivers!${c}4 should show the ${txt(D,'B18')} growth for that year (${pct(V(W,'Drivers',c+SCEN_ROW[s]))}). Check the $ signs on $B$18 and $A$20:$A$22.`); }
      return P(); },
    W => { const l=[]; let p=V(W,'IS','D4'); for(const c of FC){ const g=V(W,'Drivers',c+'4'); if(!isFinite(g)) return F(`Drivers!${c}4 isn't a number yet. Finish quest FM-1.3 first.`);
        const want=p*(1+g); l.push(['IS',c+'4',want,`Last year times (1 + growth): =${prev(c)}4*(1+Drivers!${c}4).`]); p=want; }
      let r=expect(W,l); if(!r.pass) return r;
      for(const c of FC){ if(!/DRIVERS!/.test(norm(cell(W.sheets.IS,c+'4').f))) return F(`IS!${c}4 has ${fx(W.sheets.IS,c+'4')}. Point at the growth driver on the Drivers sheet instead of typing the rate.`); }
      return expect(W, FC.map(c=>['IS',c+'5',V(W,'IS',c+'4')/V(W,'IS',prev(c)+'4')-1,`Fill C5 right: =${c}4/${prev(c)}4-1.`,true])); },
    W => { const D=W.sheets.Drivers, xs=HIST.map(c=>V(W,'Drivers',c+'3')), ys=HIST.map(c=>V(W,'IS',c+'4'));
      for(const c of FC){ const f=norm(cell(D,c+'24').f); if(!f) return F(`Drivers!${c}24 is empty. =FORECAST.LINEAR(${c}$3,IS!$B$4:$D$4,$B$3:$D$3)`);
        if(!/FORECAST|TREND\(/.test(f)) return F(`Drivers!${c}24 has ${fx(D,c+'24')}. Use FORECAST.LINEAR.`);
        const want=linFit(xs,ys,V(W,'Drivers',c+'3')); if(!close(num(val(D,c+'24')),want)) return F(`Drivers!${c}24 should be ${k1(want)}. Lock the known ranges with $ (IS!$B$4:$D$4 and $B$3:$D$3) and use ${c}$3 for the year.`); }
      return P(); },
    W => { const D=W.sheets.Drivers, START={5:.32,6:.18,7:.016,8:.02,9:3,10:90,11:19,12:.25,13:.08,14:20,15:30};
      const hasF=[], notBlue=[];
      for(let r=5;r<=15;r++) for(const c of FC){ const x=cell(D,c+r); if(x.f) hasF.push([c,r]); if(!isBlue(x)) notBlue.push([c,r]); }
      if(hasF.length){ const [c,r]=hasF[0], s=START[r]<1?pct(START[r]):String(START[r]);
        return F(`Drivers!${c}${r} should be a typed assumption, but it holds ${fx(D,c+r)}${hasF.length>1?` (${hasF.length} input cells hold formulas)`:''}. Type ${s} back in${hasF.length>1?' for each one':''}. A formula there usually comes from filling a history formula too far to the right.`); }
      if(notBlue.length){ const rows=[...new Set(notBlue.map(x=>x[1]))];
        return notBlue.length===1 ? F(`Drivers!${notBlue[0][0]}${notBlue[0][1]} is a typed input but isn't blue. Pick Blue under Standard Colors (not Theme Colors).`)
          : F(`${notBlue.length} typed inputs on Drivers aren't blue yet (row${rows.length>1?'s':''} ${rows.join(', ')}). Click E5 and drag to I15 so every row from 5 to 15 is selected, then pick Blue under Standard Colors (not Theme Colors).`); }
      const sc=[]; for(let r=20;r<=22;r++) for(const c of FC){ if(!isBlue(cell(D,c+r))) sc.push(c+r); }
      if(sc.length) return F(`${sc.length===1?'Drivers!'+sc[0]+' (scenario growth) is a typed input':sc.length+' scenario growth cells in E20:I22 are typed inputs'}. Make ${sc.length===1?'it':'them'} blue.`);
      if(!isBlue(cell(D,'B18'))) return F('Drivers!B18 (the scenario) is an input too. Make it blue.');
      for(const c of FC){ if(isBlue(cell(D,c+'4'))) return F(`Drivers!${c}4 is a formula, so it should stay black.`); }
      return P(); }
  ];

  /* ---------- three-statement link (FM-2) ----------
     Each quest is checked against the student's OWN upstream cells, so one mistake fails one quest, not all of them. */
  const L2 = [
    W => { const l=[]; FC.forEach(c=>{ const rev=V(W,'IS',c+'4'), dr=k=>V(W,'Drivers',c+k), cogs=rev*(1-dr('5')), gp=rev-cogs, sga=rev*dr('6'), da=rev*dr('7');
        l.push(['IS',c+'6',cogs,`Revenue times (1 - gross margin): =${c}4*(1-Drivers!${c}5).`], ['IS',c+'7',gp,`=${c}4-${c}6.`], ['IS',c+'8',gp/rev,`=${c}7/${c}4.`,true],
        ['IS',c+'9',sga,`=${c}4*Drivers!${c}6.`], ['IS',c+'10',da,`=${c}4*Drivers!${c}7.`], ['IS',c+'11',gp-sga-da,`=${c}7-${c}9-${c}10.`]); }); return expect(W,l); },
    W => { const l=[]; FC.forEach(c=>{ const p=prev(c), dr=k=>V(W,'Drivers',c+k), loanP=V(W,'BS',p+'11'), int=dr('13')*loanP, ebt=V(W,'IS',c+'11')-int, tax=Math.max(0,ebt)*dr('12');
        l.push(['IS',c+'12',int,`Interest on LAST year's loan: =Drivers!${c}13*BS!${p}11. Using this year's loan makes a circular reference.`],
        ['IS',c+'13',ebt,`=${c}11-${c}12.`], ['IS',c+'14',tax,`Taxes can't go below zero: =MAX(0,${c}13)*Drivers!${c}12.`], ['IS',c+'15',ebt-tax,`=${c}13-${c}14.`], ['IS',c+'16',V(W,'IS',c+'11')+V(W,'IS',c+'10'),`=${c}11+${c}10.`]); });
      return expect(W,l); },
    W => { const l=[]; FC.forEach(c=>{ const dr=k=>V(W,'Drivers',c+k); l.push(['BS',c+'5',V(W,'IS',c+'4')*dr('9')/365,`Revenue x receivable days / 365: =IS!${c}4*Drivers!${c}9/365.`],
        ['BS',c+'6',V(W,'IS',c+'6')*dr('10')/365,`Inventory runs on COGS, not revenue: =IS!${c}6*Drivers!${c}10/365.`], ['BS',c+'10',V(W,'IS',c+'6')*dr('11')/365,`Payables run on COGS: =IS!${c}6*Drivers!${c}11/365.`]); }); return expect(W,l); },
    W => { const l=[]; FC.forEach(c=>{ const p=prev(c), dr=k=>V(W,'Drivers',c+k);
        l.push(['BS',c+'7',V(W,'BS',p+'7')+V(W,'IS',c+'4')*dr('8')-V(W,'IS',c+'10'),`Last year + capex - depreciation: =${p}7+IS!${c}4*Drivers!${c}8-IS!${c}10.`],
        ['BS',c+'11',Math.max(0,V(W,'BS',p+'11')-dr('14')),`=MAX(0,${p}11-Drivers!${c}14).`], ['BS',c+'14',V(W,'BS',p+'14'),`=${p}14.`],
        ['BS',c+'15',V(W,'BS',p+'15')+V(W,'IS',c+'15')-dr('15'),`Last year + net income + draws (negative on CF): =${p}15+IS!${c}15+CF!${c}15.`]); }); return expect(W,l); },
    W => { const l=[]; FC.forEach(c=>{ const p=prev(c), dr=k=>V(W,'Drivers',c+k), b=r=>V(W,'BS',c+r), bp=r=>V(W,'BS',p+r);
        const ni=V(W,'IS',c+'15'), da=V(W,'IS',c+'10'), dar=-(b('5')-bp('5')), dinv=-(b('6')-bp('6')), dap=b('10')-bp('10'), cfo=ni+da+dar+dinv+dap;
        const capex=-V(W,'IS',c+'4')*dr('8'), rep=b('11')-bp('11'), draws=-dr('15'), cff=rep+draws, net=cfo+capex+cff, beg=V(W,'CF',p+'20'), end=beg+net;
        l.push(['CF',c+'4',ni,`=IS!${c}15.`], ['CF',c+'5',da,`Add back depreciation: =IS!${c}10.`], ['CF',c+'6',dar,`=-(BS!${c}5-BS!${p}5).`], ['CF',c+'7',dinv,`More inventory uses cash: =-(BS!${c}6-BS!${p}6).`],
          ['CF',c+'8',dap,`=BS!${c}10-BS!${p}10.`], ['CF',c+'9',cfo,`=SUM(${c}4:${c}8).`], ['CF',c+'11',capex,`=-IS!${c}4*Drivers!${c}8.`], ['CF',c+'12',capex,`=${c}11.`], ['CF',c+'14',rep,`=BS!${c}11-BS!${p}11.`],
          ['CF',c+'15',draws,`=-Drivers!${c}15.`], ['CF',c+'16',cff,`=SUM(${c}13:${c}15).`], ['CF',c+'18',net,`=${c}9+${c}12+${c}16.`], ['CF',c+'19',beg,`=${p}20.`], ['CF',c+'20',end,`=${c}19+${c}18.`],
          ['BS',c+'4',end,`Cash comes from the cash flow statement: =CF!${c}20.`]); });
      return expect(W,l); },
    W => { const B=W.sheets.BS, l=[]; FC.forEach(c=>{ const b=r=>V(W,'BS',c+r), ta=b('4')+b('5')+b('6')+b('7'), tl=b('10')+b('11'), te=b('14')+b('15');
        l.push(['BS',c+'8',ta,`=SUM(${c}4:${c}7).`], ['BS',c+'12',tl,`=SUM(${c}10:${c}11).`], ['BS',c+'16',te,`=SUM(${c}14:${c}15).`], ['BS',c+'17',tl+te,`=${c}12+${c}16.`], ['BS',c+'19',ta-tl-te,`=${c}8-${c}17.`]); });
      const r=expect(W,l); if(!r.pass) return r;
      for(const c of FC){ const gap=V(W,'BS',c+'19');
        if(Math.abs(gap)>=0.01) return F(`The model doesn't balance in ${'FY'+(2023+'BCDEFGHI'.indexOf(c))}: BS!${c}19 is off by ${k1(gap)}. Compare that gap to D&A (${k1(V(W,'IS',c+'10'))}), capex, and each working-capital change, then use Trace Precedents on the cell that matches.`);
        if(norm(cell(B,c+'20').f).indexOf('IF(')<0) return F(`BS!${c}20 has ${fx(B,c+'20')}. Fill D20 right: =IF(ABS(${c}19)<0.01,"OK","CHECK").`);
        if(txt(B,c+'20')!=='OK') return F(`BS!${c}20 says ${txt(B,c+'20')||'nothing'}. It should say OK once the check is 0.`); }
      return P(); },
    null
  ];

  /* ---------- comparable companies (FM-3) ---------- */
  const PEER_ROWS = [5,6,7,8,9,10];
  function qinc(xs, q){ const a=[...xs].sort((x,y)=>x-y); if(!a.length) return NaN; const p=(a.length-1)*q, lo=Math.floor(p), hi=Math.min(lo+1,a.length-1); return a[lo]+(a[hi]-a[lo])*(p-lo); }
  const med = xs => qinc(xs, .5);
  const mult = n => n.toFixed(2)+'x';
  function peerMults(W){ const S=W.sheets.Comps, ys=PEER_ROWS.filter(r=>txt(S,'H'+r).toUpperCase()==='Y');
    return {rev: ys.map(r=>V(W,'Comps','C'+r)/V(W,'Comps','D'+r)), ebitda: ys.map(r=>V(W,'Comps','C'+r)/V(W,'Comps','E'+r))}; }
  const L3 = [
    W => { const S=W.sheets.Comps, l=[]; PEER_ROWS.forEach(r=>{ l.push(['Comps','F'+r,V(W,'Comps','C'+r)/V(W,'Comps','D'+r),`EV divided by revenue: =C${r}/D${r}.`], ['Comps','G'+r,V(W,'Comps','C'+r)/V(W,'Comps','E'+r),`EV divided by EBITDA: =C${r}/E${r}.`]); });
      const r=expect(W,l); if(!r.pass) return r;
      for(const rr of PEER_ROWS) for(const c of ['F','G']){ const fm=(cell(S,c+rr).xf && cell(S,c+rr).xf.fmt)||''; if(!/x/i.test(fm.replace(/\\/g,''))) return F(`Comps!${c}${rr} is right but has no x on it. Select F5:G10, press Cmd+1, choose Custom, and type 0.0"x".`); }
      return P(); },
    W => { const S=W.sheets.Comps;
      if(txt(S,'H10').toUpperCase()!=='N') return F(`Comps!H10 says ${txt(S,'H10')||'nothing'}. Read Courtline Athletic's description in B10: it makes its own shoes, so it isn't a true peer. Type N.`);
      for(const r of [5,6,7,8,9]){ if(txt(S,'H'+r).toUpperCase()!=='Y') return F(`Comps!H${r} (${txt(S,'A'+r)}) should stay Y. Only Courtline Athletic leaves the peer set.`); }
      return P(); },
    W => { const S=W.sheets.Comps, m=peerMults(W), l=[];
      [['B',m.rev,'F'],['C',m.ebitda,'G']].forEach(([c,xs,src])=>{ l.push([ 'Comps',c+'13',qinc(xs,.25),`=QUARTILE.INC(FILTER(${src}$5:${src}$10,$H$5:$H$10="Y"),1).`],
        ['Comps',c+'14',med(xs),`=MEDIAN(FILTER(${src}$5:${src}$10,$H$5:$H$10="Y")).`], ['Comps',c+'15',qinc(xs,.75),`=QUARTILE.INC(FILTER(${src}$5:${src}$10,$H$5:$H$10="Y"),3).`]); });
      for(const [sn,ref,want,hint] of l){ const c=cell(S,ref);
        if(!c.f) return F(c.v===null||c.v===undefined||c.v==='' ? `Comps!${ref} is empty. ${hint}` : `Comps!${ref} holds a typed number. Make it a formula: ${hint}`);
        if(!close(num(c.v),want)) return F(`Comps!${ref} shows ${typeof c.v==='number'?mult(c.v):(c.v||'nothing')}. Using only the peers marked Y, it should be ${mult(want)}. ${hint}`);
        if(!/\$?H\$?5/.test(norm(c.f))) return F(`Comps!${ref} has the right number, but the formula doesn't read the Y/N column. Use FILTER with $H$5:$H$10="Y" so the peer set can change.`); }
      return P(); },
    W => { const l=[], rev=V(W,'IS','D4'), eb=V(W,'IS','D16');
      [0,1,2].forEach(i=>{ l.push(['Comps','B'+(18+i),V(W,'Comps','B'+(13+i))*rev,`Multiple times FY2025 revenue: =B${13+i}*IS!$D$4.`], ['Comps','C'+(18+i),V(W,'Comps','C'+(13+i))*eb,`Multiple times FY2025 EBITDA: =C${13+i}*IS!$D$16.`]); });
      const r=expect(W,l); if(!r.pass) return r;
      for(const ref of ['B18','B19','B20','C18','C19','C20']){ if(!/IS!/.test(norm(cell(W.sheets.Comps,ref).f))) return F(`Comps!${ref} has the right number but doesn't point at the income statement. Link it to IS!$D$4 or IS!$D$16 so it updates.`); }
      return P(); },
    W => { const nd=V(W,'BS','D11')-V(W,'BS','D4'); let r=expect(W,[['Comps','B27',nd,'Loan minus cash in FY2025: =BS!D11-BS!D4.']]); if(!r.pass) return r;
      const l=[]; [0,1,2].forEach(i=>{ l.push(['Comps','B'+(23+i),V(W,'Comps','B'+(18+i))-V(W,'Comps','B27'),`Enterprise value minus net debt: =B${18+i}-$B$27.`], ['Comps','C'+(23+i),V(W,'Comps','C'+(18+i))-V(W,'Comps','B27'),`=C${18+i}-$B$27.`]); });
      return expect(W,l); }
  ];

  /* ---------- discounted cash flow (FM-4) ---------- */
  const fcfOf = W => FC.map(c=>V(W,'DCF',c+'10'));
  function eqAt(W, r, g){ const f=fcfOf(W), nd=V(W,'DCF','C23'); let pv=0; f.forEach((x,i)=>{ pv+=x/Math.pow(1+r,i+1); }); return pv + f[4]*(1+g)/(r-g)/Math.pow(1+r,5) - nd; }
  const L4 = [
    W => { const l=[]; FC.forEach(c=>{ const ebit=V(W,'IS',c+'11'), tax=-ebit*V(W,'Drivers',c+'12'), da=V(W,'IS',c+'10'), cap=V(W,'CF',c+'11'), wc=V(W,'CF',c+'6')+V(W,'CF',c+'7')+V(W,'CF',c+'8');
        l.push(['DCF',c+'4',ebit,`=IS!${c}11.`], ['DCF',c+'5',tax,`Tax on EBIT, as a negative: =-${c}4*Drivers!${c}12.`], ['DCF',c+'6',ebit+tax,`=${c}4+${c}5.`],
          ['DCF',c+'7',da,`=IS!${c}10.`], ['DCF',c+'8',cap,`Capex is already negative on CF: =CF!${c}11.`], ['DCF',c+'9',wc,`The three working-capital rows on CF: =SUM(CF!${c}6:${c}8).`],
          ['DCF',c+'10',ebit+tax+da+cap+wc,`=${c}6+${c}7+${c}8+${c}9.`]); });
      return expect(W,l); },
    W => { const D=W.sheets.DCF, r=V(W,'DCF','C16');
      for(let i=0;i<5;i++){ const c=FC[i]; if(!close(V(W,'DCF',c+'12'),i+1)) return F(`DCF!${c}12 should be ${i+1}. Type 1 to 5 across E12:I12.`); }
      const l=[]; FC.forEach((c,i)=>{ const df=1/Math.pow(1+r,i+1); l.push(['DCF',c+'13',df,`=1/(1+$C$16)^${c}12. Lock C16 with $ signs.`], ['DCF',c+'14',V(W,'DCF',c+'10')*df,`=${c}10*${c}13.`]); });
      return expect(W,l); },
    W => { const r=V(W,'DCF','C16'), f=fcfOf(W); let npv=0; f.forEach((x,i)=>{ npv+=x/Math.pow(1+r,i+1); });
      let res=expect(W,[['DCF','C19',FC.reduce((t,c)=>t+V(W,'DCF',c+'14'),0),'=SUM(E14:I14).'], ['DCF','C18',npv,'=NPV(C16,E10:I10). Start the range at the first forecast year, not a today amount.']]); if(!res.pass) return res;
      return /NPV\(/.test(norm(cell(W.sheets.DCF,'C18').f)) ? P() : F(`DCF!C18 has the right number but not the NPV function. Use =NPV(C16,E10:I10) so you see both methods agree.`); },
    W => { const r=V(W,'DCF','C16'), g=V(W,'DCF','C17'), last=V(W,'DCF','I10');
      if(!(r>g)) return F(`The discount rate in C16 has to be bigger than the growth rate in C17, or the terminal value formula breaks.`);
      const tv=last*(1+g)/(r-g);
      return expect(W,[['DCF','C20',tv,'Next year\'s cash flow divided by (rate - growth): =I10*(1+C17)/(C16-C17).'], ['DCF','C21',V(W,'DCF','C20')/Math.pow(1+r,5),'Discount it 5 years: =C20/(1+C16)^5.']]); },
    W => { const ev=V(W,'DCF','C19')+V(W,'DCF','C21'), nd=V(W,'BS','D11')-V(W,'BS','D4');
      return expect(W,[['DCF','C22',ev,'=C19+C21.'], ['DCF','C23',nd,'Loan minus cash in FY2025: =BS!D11-BS!D4.'], ['DCF','C24',V(W,'DCF','C22')-V(W,'DCF','C23'),'=C22-C23.'],
        ['DCF','C25',V(W,'DCF','C21')/V(W,'DCF','C22'),'=C21/C22.',true]]); },
    W => { const D=W.sheets.DCF;
      let res=expect(W,[['DCF','B28',V(W,'DCF','C24'),'The corner cell points at equity value: =C24.']]); if(!res.pass) return res;
      const rates=[29,30,31,32,33].map(r=>V(W,'DCF','B'+r)), grow=['C','D','E','F','G'].map(c=>V(W,'DCF',c+'28'));
      if(rates.some(x=>!isFinite(x))) return F('Type the discount rates 12%, 13%, 14%, 15%, 16% down B29:B33.');
      if(grow.some(x=>!isFinite(x))) return F('Type the growth rates 2.0%, 2.5%, 3.0%, 3.5%, 4.0% across C28:G28.');
      if(rates.some(x=>x>1) || grow.some(x=>x>1)) return F('The rates look like whole numbers (14 instead of 14%). Type them with a % sign.');
      const dtCell=Object.values(D.cells).find(c=>c.dt);
      const vals=[29,30,31,32,33].map(r=>['C','D','E','F','G'].map(c=>num(val(D,c+r))));
      if(!dtCell){ return vals.flat().every(x=>!isFinite(x)) ? F('C29:G33 is empty. Select B28:G33 and choose Data, What-If Analysis, Data Table.')
          : F('C29:G33 has numbers but no Data Table. Clear C29:G33, select B28:G33, and choose Data, What-If Analysis, Data Table.'); }
      if(!dtCell.dt.dt2D) return F('That Data Table only uses one input. Rebuild it with both: Row input cell C17, Column input cell C16.');
      const ok=(rr,gg)=>vals.every((row,i)=>row.every((x,j)=>close(x,eqAt(W,rr(i,j),gg(i,j)))));
      if(vals.flat().some(x=>!isFinite(x))) return F('The Data Table has no numbers yet. Press Cmd+= to calculate, then save.');
      if(ok((i,j)=>rates[i],(i,j)=>grow[j])) return P();
      if(ok((i,j)=>grow[j],(i,j)=>rates[i])) return F('The inputs are swapped. Rebuild it with Row input cell C17 (growth runs across row 28) and Column input cell C16 (rates run down column B).');
      return F(`The table's numbers don't match the model. Base case at ${pct(V(W,'DCF','C16'))} and ${pct(V(W,'DCF','C17'))} should be ${k1(V(W,'DCF','C24'))}. Check that B28 is =C24, then rebuild with Row input C17 and Column input C16.`); }
  ];

  /* ---------- the investor's offer (FM-B) ---------- */
  const dcfRange = W => { const xs=[]; for(let r=29;r<=33;r++) for(const c of ['C','D','E','F','G']){ const x=num(val(W.sheets.DCF,c+r)); if(isFinite(x)) xs.push(x); } return xs; };
  const LB = [
    W => expect(W,[['Summary','B5',V(W,'Summary','B3')/V(W,'Summary','B4'),'Investment divided by stake: =B3/B4.'], ['Summary','B6',V(W,'Summary','B5')-V(W,'Summary','B3'),'Post-money minus the investment: =B5-B3.']]),
    W => { const S=W.sheets.Summary, d=dcfRange(W); if(!d.length) return F('The DCF sensitivity table (DCF!C29:G33) has no numbers. Open the file in Excel, press Cmd+= and save.');
      const l=[['Summary','B9',V(W,'Comps','B23'),'=Comps!B23.'], ['Summary','C9',V(W,'Comps','B25'),'=Comps!B25.'], ['Summary','B10',V(W,'Comps','C23'),'=Comps!C23.'], ['Summary','C10',V(W,'Comps','C25'),'=Comps!C25.'],
        ['Summary','B11',Math.min(...d),'=MIN(DCF!C29:G33).'], ['Summary','C11',Math.max(...d),'=MAX(DCF!C29:G33).']];
      let r=expect(W,l); if(!r.pass) return r;
      for(const ref of ['B9','C9','B10','C10']){ if(!/COMPS!/.test(norm(cell(S,ref).f))) return F(`Summary!${ref} has the right number but doesn't link to the Comps sheet. Point at it so the chart updates.`); }
      return expect(W,[9,10,11].map(k=>['Summary','D'+k,V(W,'Summary','C'+k)-V(W,'Summary','B'+k),`High minus low: =C${k}-B${k}.`])); },
    W => { const S=W.sheets.Summary, all=(S.charts||[]).concat(...Object.values(W.chartsheets||{}).map(c=>c.charts||[]));
      if(!all.length) return F('There is no chart yet. Select A9:B11, hold Cmd, add D9:D11, and choose Insert, Bar, Stacked Bar.');
      const bar=all.find(ch=>ch.kinds && ch.kinds.some(k=>k.kind==='barChart'));
      if(!bar) return F('The chart isn\'t a bar chart. Delete it and choose Insert, Bar, Stacked Bar.');
      const k=bar.kinds.find(k=>k.kind==='barChart');
      if(k.barDir!=='bar') return F('That\'s a column chart (vertical bars). A football field uses horizontal bars: choose Insert, Bar, Stacked Bar.');
      if(k.grouping!=='stacked') return F('The bars need to be stacked, so the spread sits on top of the low value. Change the chart type to Stacked Bar.');
      const lowS=bar.series.find(s=>/B\$?9:\$?B\$?11/.test(norm(s.val))), sprS=bar.series.find(s=>/D\$?9:\$?D\$?11/.test(norm(s.val)));
      if(!lowS || !sprS) return F('The chart should use two series: the low values (B9:B11) and the spread (D9:D11), with the method names in A9:A11 as categories.');
      if(bar.series.some(s=>/C\$?9:\$?C\$?11/.test(norm(s.val)))) return F('Leave the High column (C9:C11) out of the chart. Low plus spread already reaches the high value.');
      if(!lowS.noFill) return F('The Low bars still show. Right-click them, choose Format Data Series, and set Fill to No Fill so each bar floats.');
      if(!/1,?000|OFFER/i.test(bar.title||'')) return F(`The chart title ${bar.title?'"'+bar.title+'"':'is missing. It'} should name the offer, like Equity value ($000) vs. offer at 1,000.`);
      return P(); },
    W => { const S=W.sheets.Summary, c=cell(S,'B14');
      if(!c.f) return F(c.v ? 'Summary!B14 holds typed text. Make it a formula so the call updates when the model changes.' : 'Summary!B14 is empty. =IF(B6<DCF!C24,"Counter: the offer is below our base case","Accept: the offer meets our base case")');
      const f=norm(c.f); if(f.indexOf('IF(')<0 || !/DCF!\$?C\$?24/.test(f) || !/B\$?6/.test(f)) return F(`Summary!B14 has ${fx(S,'B14')}. Compare the pre-money in B6 with the DCF base case in DCF!C24 inside an IF.`);
      const want=V(W,'Summary','B6')<V(W,'DCF','C24') ? 'COUNTER' : 'ACCEPT';
      return txt(S,'B14').toUpperCase().indexOf(want)===0 ? P() : F(`Summary!B14 says ${txt(S,'B14')||'nothing'}. With pre-money at ${k1(V(W,'Summary','B6'))} and the base case at ${k1(V(W,'DCF','C24'))}, it should start with ${want==='COUNTER'?'Counter':'Accept'}. Check which way the < points.`); },
    null
  ];

  const reg = XQ.register;
  reg('FM-0', {check:L0, sheets:['IS','BS','CF','Notes','Sort It']});
  reg('FM-1', {check:L1, sheets:['Drivers','IS','BS','CF','Notes']});
  reg('FM-2', {check:L2, sheets:['Drivers','IS','BS','CF','Notes']});
  reg('FM-3', {check:L3, sheets:['Drivers','IS','BS','CF','Notes','Comps']});
  reg('FM-4', {check:L4, sheets:['Drivers','IS','BS','CF','Notes','Comps','DCF']});
  reg('FM-B', {check:LB, sheets:['Drivers','IS','BS','CF','Notes','Comps','DCF','Summary']});
})();

/* Excel Quest · Financial Modeling world (live-site module). Loads after the grader scripts and before the main app script.
   The main app calls XQFM.owns(h), XQFM.screen(h, ctx), XQFM.after(h), XQFM.mapCard(ctx), and XQFM.codes().
   Saves under 'excelquest-fm-v1'. Locked until every MOS and TOSA level has a star (or a teacher preview is on for this tab). */
window.XQFM = (function(){
  const FKEY = 'excelquest-fm-v1', PKEY = 'excelquest-fm-preview';
  const PASS_HASH = 'b773713b605f9';
  const QXP = 25, GXP = 100;
  let fs = (()=>{ try { return JSON.parse(localStorage.getItem(FKEY)||'{}') || {}; } catch(e){ return {}; } })();
  fs.g = fs.g || {}; fs.kc = fs.kc || {}; fs.kcPick = fs.kcPick || {};
  const save = () => { try { localStorage.setItem(FKEY, JSON.stringify(fs)); } catch(e){} };
  const preview = () => { try { return sessionStorage.getItem(PKEY)==='1'; } catch(e){ return false; } };
  const NKEY = 'excelquest-student', MKEY = 'excelquest-save-v2', TKEY = 'excelquest-tosa-v1';
  const getLS = k => { try { return JSON.parse(localStorage.getItem(k)||'{}') || {}; } catch(e){ return {}; } };
  const setLS = (k,v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch(e){ return false; } };
  const getName = () => { try { return localStorage.getItem(NKEY) || ''; } catch(e){ return ''; } };
  const setName = n => { try { localStorage.setItem(NKEY, n); } catch(e){} };
  const today = () => new Date().toISOString().slice(0,10);
  let ctx = null;

  const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  function fmt(s){
    let h = esc(s);
    h = h.replace(/`([^`]+)`/g,(m,c)=>'<code>'+c+'</code>');
    h = h.replace(/\[\[(.+?)\]\]/g,(m,p)=>'<span class="menu">'+p.split(' &gt; ').join('<wbr><span class="chev">›</span><wbr>')+'</span>');
    h = h.replace(/\{\{(.+?)\}\}/g,(m,k)=>'<kbd>'+k+'</kbd>');
    h = h.replace(/\*\*(.+?)\*\*/g,'<b>$1</b>');
    return h;
  }
  function hash(str){ let h1=0xdeadbeef^7, h2=0x41c6ce57^7; for(let i=0;i<str.length;i++){ const ch=str.charCodeAt(i); h1=Math.imul(h1^ch,2654435761); h2=Math.imul(h2^ch,1597334677); }
    h1=Math.imul(h1^(h1>>>16),2246822507)^Math.imul(h2^(h2>>>13),3266489909); h2=Math.imul(h2^(h2>>>16),2246822507)^Math.imul(h1^(h1>>>13),3266489909);
    return (4294967296*(2097151&h2)+(h1>>>0)).toString(16); }
  const k1 = n => { if(n==null || isNaN(n)) return ''; const a = Math.abs(n)<0.05?0:n; const s = Math.abs(a).toLocaleString('en-US',{minimumFractionDigits:1,maximumFractionDigits:1}); return a<0?'('+s+')':s; };

  /* ---------- gate ---------- */
  const WNAME = {1:'World 1 · Associate Isles', 2:'World 2 · Expert Fortress', E:'Workbook Wharf', F:'Functions Forest', M:'Style Summit', D:'Data Delta'};
  function groups(c){ const out=[]; const add=(track,lv)=>{ let g=out.find(x=>x.key===track+lv.world); if(!g){ g={key:track+lv.world, track, name:WNAME[lv.world]||('World '+lv.world), lv:[]}; out.push(g); } g.lv.push({id:lv.id, star:c.levelStats(lv).star}); };
    c.MOS.forEach(l=>add('MOS',l)); c.TOSA.forEach(l=>add('TOSA',l)); return out; }
  function progress(c){ const g=groups(c); const all=g.flatMap(x=>x.lv); return {g, done:all.filter(x=>x.star).length, total:all.length}; }
  const earned = c => { const p=progress(c); return p.total>0 && p.done===p.total; };
  const open = c => earned(c) || !!fs.unlocked || preview();
  /* Once a student earns the world it stays open, even if new quests are added to MOS or TOSA later. */
  function stick(c){ if(c && !fs.unlocked && earned(c)){ fs.unlocked = 'earned'; fs.unlockedAt = today(); save(); } }
  const fmStats = l => ({q:l.quiz.length, f:l.quiz.filter((x,i)=>fs.kc[l.id+'.'+i]===true).length, done:(fs.g[l.id]||[]).filter(x=>x&&x.p).length, total:l.quests.filter(q=>!q.note).length});
  const fmXP = () => Object.values(fs.kc).filter(v=>v===true).length*QXP + Object.values(fs.g).reduce((a,arr)=>a+arr.filter(x=>x&&x.p).length,0)*GXP;

  function gateScreen(c){
    const p = progress(c);
    return `<section class="title-screen fm-title"><div class="over">EXCEL QUEST · BONUS WORLD</div><h1>FINANCIAL<br>MODELING</h1>
      <div class="sub">Three statements · Forecasting · Comps · DCF</div></section>
    <section class="frame box fm-gate"><p class="fm-lock">🔒 LOCKED</p>
      <p style="margin:0;max-width:64ch">This world opens when every MOS level and every TOSA level shows ★ CLEAR. Modeling uses nearly everything from both tracks: lookups, FILTER, number formats, Data Tables, and charts.</p>
      <div class="fm-prog"><span>${p.done}/${p.total}</span><span class="bar fm-bar"><i style="width:${p.total?Math.round(100*p.done/p.total):0}%"></i></span><span>${p.total-p.done} to go</span></div>
      <div class="fm-groups">${p.g.map(g=>{ const d=g.lv.filter(x=>x.star).length; return `<div class="fm-gg"><h4>${esc(g.name)} <small>· ${g.track} · ${d}/${g.lv.length}</small></h4><div class="fm-chips">${g.lv.map(x=>`<a class="fm-chip ${x.star?'on':''}" href="#level-${esc(x.id)}">${esc(x.id)}</a>`).join('')}</div></div>`; }).join('')}</div>
      <p style="margin:14px 0 0;color:var(--muted);font-size:15px">Click any level to go finish it. Your progress here comes from the same save files as the MOS and TOSA maps.</p></section>
    <section class="frame box fm-lost"><h3>LOST YOUR PROGRESS?</h3>
      <p style="margin:0 0 12px">Progress lives in this browser only. If it disappeared (new computer, cleared history, a different browser), paste your backup code on the <a href="#backup">Save Backup</a> page. No backup? Ask your teacher for an unlock code.</p>
      <form id="fm-redeem" class="fm-form"><label for="fm-rname">Your full name<input id="fm-rname" type="text" autocomplete="name" value="${esc(getName())}"></label>
        <label for="fm-rcode">Teacher unlock code<input id="fm-rcode" type="text" autocomplete="off" placeholder="XXXX-XXXX" style="text-transform:uppercase"></label>
        <button type="submit" class="press fm-press" style="margin-top:0">UNLOCK</button></form>
      <p class="turnin-msg" id="fm-rmsg" role="status"></p></section>
    <section class="world frame" style="opacity:.75"><div class="world-head"><h2 style="color:var(--muted)">WHAT'S INSIDE</h2><span class="exam">${FM_LEVELS.length} levels</span></div>
      <div class="path">${FM_LEVELS.map(l=>`<div class="node locked"><span class="num">${l.id}</span><span class="nm">${esc(l.name)}</span><span class="dm">${esc(l.topic)}</span></div>`).join('')}</div></section>`;
  }
  const built = id => !!FM_FILES[id] && (FM_TEACHER_ONLY.indexOf(id)<0 || preview());
  function mapScreen(c){
    return `${preview() && !earned(c) && !fs.unlocked?`<p class="fm-banner">TEACHER PREVIEW · Students see this world only after every MOS and TOSA level is starred. <button type="button" class="retry" data-fmexit>Exit preview</button></p>`:''}
    <section class="title-screen fm-title"><div class="over">EXCEL QUEST · BONUS WORLD · UNLOCKED</div><h1>FINANCIAL<br>MODELING</h1>
      <div class="sub">Client: Magic City Sneaker Exchange<br>An investor wants 20%. What is the shop worth?</div>
      <div class="cart"><span>${FM_LEVELS.length} levels</span><span>${FM_LEVELS.reduce((a,l)=>a+l.quests.filter(q=>!q.note).length,0)} quests</span><span>${FM_CODEX.length} Codex chapters</span><span>Modeling XP</span></div>
      <div><a class="press fm-press" href="#fm-FM-0"><span class="blink">▶ PRESS START</span></a></div></section>
    <section class="frame box"><h3>THE STORY</h3><p style="margin:0;max-width:66ch">You helped this shop decide on its expansion loan back in Level 2-3. Two years later it's growing, and a local sneaker brand offers $250,000 for 20% of the business. Each level builds one piece of the answer, and the boss level makes the call.</p>
      <p style="margin:10px 0 0;color:var(--muted);font-size:15px">Modeling XP is its own score: ${GXP} per graded quest, ${QXP} per first-try Knowledge Check answer. It doesn't count toward the XP exchange.</p></section>
    <section class="world frame fmw"><div class="world-head"><h2>FINANCIAL MODELING WORLD</h2><span class="exam">Primer + 4 levels + boss</span></div>
      <p class="blurb">Work in order. Every level file starts with the finished work from the level before, so one rough level never blocks the next.</p>
      <div class="path">${FM_LEVELS.map(l=>{ const s=fmStats(l); return built(l.id)
        ? `<a class="node" href="#fm-${l.id}">${s.done===s.total?'<span class="star">★ CLEAR</span>':''}<span class="num">${l.id}</span><span class="nm">${esc(l.name)}</span><span class="dm">${esc(l.topic)}</span>${FM_TEACHER_ONLY.indexOf(l.id)>=0?'<span class="ct" style="color:var(--gold)">Playtest · hidden from students</span>':''}<span class="ct">${s.done}/${s.total} quests · ${s.f}/${s.q} check</span><span class="bar"><i style="width:${Math.round(100*s.done/s.total)}%"></i></span></a>`
        : `<div class="node locked"><span class="num">${l.id}</span><span class="nm">${esc(l.name)}</span><span class="dm">${esc(l.topic)}</span><span class="ct">Coming soon</span></div>`; }).join('')}</div></section>
    <div class="cards" style="margin-top:28px">
      <a class="frame box fm-card" href="#fm-lab"><h3 style="color:var(--mint)">MODEL LAB</h3><p>The finished model, live. Change any assumption and watch the statements, the DCF, and the football field move. Includes a bug hunt.</p></a>
      <a class="frame box fm-card" href="#fm-codex"><h3>MODELING CODEX</h3><p>${FM_CODEX.length} chapters: how the statements connect, why each formula works, and the traps.</p></a>
      <a class="frame box fm-card" href="#codes"><h3>CHEAT CODES</h3><p>Mac shortcuts and the house rules every analyst follows.</p></a>
    </div>
    <section class="frame box" style="margin-top:28px"><h3>THE WORKBOOK MAP</h3><p style="margin:0 0 10px;color:var(--muted);font-size:15px">Every level file uses the same sheets and row numbers.</p>
      <div class="tbl-wrap"><table class="codes"><thead><tr><th>SHEET</th><th>WHAT'S ON IT</th></tr></thead><tbody>${FM_MAP.map(r=>`<tr><td><b>${esc(r[0])}</b></td><td>${esc(r[1])}</td></tr>`).join('')}</tbody></table></div></section>`;
  }

  /* ---------- level ---------- */
  function levelScreen(id){
    const i = FM_LEVELS.findIndex(l=>l.id===id), l = FM_LEVELS[i]; if(!l || !built(id)) return null;
    const prev = FM_LEVELS[i-1], next = FM_LEVELS[i+1] && built(FM_LEVELS[i+1].id) ? FM_LEVELS[i+1] : null, cx = FM_CODEX.findIndex(c=>c.lvl===l.id), s = fmStats(l);
    return `<div class="fmw"><div class="crumb"><a href="#modeling">Modeling Map</a> › Level ${l.id}</div>
    <section class="lvl-head frame"><div class="lvl-tag">FINANCIAL MODELING · LEVEL ${l.id}</div><h1>${esc(l.name)}</h1>
      <div class="lvl-meta">${esc(l.topic)} · Magic City Sneaker Exchange</div><p class="lvl-story">${esc(l.story)}</p>
      <div class="goal"><b>THE GOAL</b><span>${esc(l.goal)}</span><small>Do the quests in order. Each one builds on the last.</small></div>
      <div class="lvl-prog"><span>${s.done}/${s.total}</span><span class="bar"><i style="width:${Math.round(100*s.done/s.total)}%"></i></span><span>${s.done===s.total?'★ CLEAR':''}</span></div></section>
    <div class="lvl-grid">
      <section class="box frame"><h3>INVENTORY · SKILLS YOU'LL USE</h3><div class="chips">${l.items.map(x=>`<span>${esc(x)}</span>`).join('')}</div></section>
      <section class="box frame"><h3>LOOT · PRACTICE FILE AND READING</h3><div class="loot">
        <a class="fm-dl" href="files/${encodeURIComponent(FM_FILES[l.id])}" download="${esc(FM_FILES[l.id])}"><span class="fname">${esc(FM_FILES[l.id])}</span><small>Practice workbook ↓</small></a>
        ${cx>=0?`<a class="fm-read" href="#fm-cx-${cx}">Codex: ${esc(FM_CODEX[cx].title)} <small>Read ▶</small></a>`:''}</div></section></div>
    ${turnin(l)}
    <div class="quests">${l.quests.map((q,n)=>questCard(l,q,n)).join('')}</div>
    <section class="kc" id="kc"><h2 class="page-h" style="font-size:18px">KNOWLEDGE CHECK</h2>
      <p class="lead"><b>${s.f}/${s.q}</b> right on the first try. Each is worth ${QXP} Modeling XP. You can retry any question to practice, but only your first answer counts.</p>
      <div class="kc-list">${l.quiz.map((x,n)=>quizItem(l,x,n)).join('')}</div></section>
    <nav class="lvl-nav">${prev?`<a href="#fm-${prev.id}">◀ ${prev.id} ${esc(prev.name)}</a>`:'<a href="#modeling">◀ Modeling Map</a>'}${next?`<a href="#fm-${next.id}">${next.id} ${esc(next.name)} ▶</a>`:'<a href="#fm-lab">Model Lab ▶</a>'}</nav></div>`;
  }
  function turnin(l){
    const g = fs.g[l.id], s = fmStats(l);
    const fails = g ? l.quests.map((q,i)=>({q,i,r:g[i]})).filter(x=>x.r && !x.r.p) : [];
    return `<section class="turnin frame" id="fm-turnin"><h3 class="section-h" style="font-size:11px">TURN IN · GRADER</h3>
      <label class="drop" id="fm-drop"><b>DROP YOUR SAVED WORKBOOK HERE</b><span>or click to choose the .xlsx. It's checked right here in your browser and never uploaded. The grader recomputes every answer from your own inputs, so key cells must hold formulas, not typed numbers.</span><input type="file" id="fm-file" accept=".xlsx"></label>
      <p class="turnin-msg ${g&&s.done===s.total?'ok':''}" id="fm-msg" role="status">${g?`Last turn-in: ${s.done} of ${s.total} quests passed.`:'Save your workbook in Excel first (Cmd+S), then turn it in.'}</p>
      ${fails.length?`<div class="fails"><b>STILL TO FIX</b><ul>${fails.map(x=>`<li><button type="button" class="jump" data-jump="fq-${l.id}-${x.i+1}">${l.id}.${x.i+1} ${esc(x.q.t)}</button><span>${esc(x.r.h||'')}</span></li>`).join('')}</ul></div>`:''}</section>`;
  }
  function questCard(l,q,n){
    const k = l.id+'.'+(n+1), g = (fs.g[l.id]||[])[n];
    return `<article class="quest frame ${q.note?'bp':''} ${g&&g.p?'done':''} ${g&&!g.p?'needs':''}" id="fq-${l.id}-${n+1}"><div class="q-top"><span class="q-num">${k}</span>
      <div class="q-title"><h3>${esc(q.t)}</h3><span class="badge ${q.note?'side':'main'}">${q.note?'CLASS ACTIVITY · NOT GRADED':'MAIN QUEST · '+GXP+' XP'}</span></div>
      ${q.note?'':`<label class="clear locked"><input type="checkbox" ${g&&g.p?'checked':''} disabled> ${g&&g.p?'Cleared':'Graded quest'}</label>`}</div>
      <div class="q-body">${q.why?`<p class="why"><b>WHY</b> ${fmt(q.why)}</p>`:''}
      <div class="objective"><b>OBJECTIVE</b>${fmt(q.task)}</div>
      ${g?(g.p?`<div class="grade pass"><span class="lbl">PASS</span><span>The grader found this done correctly.</span></div>`:`<div class="grade fail"><span class="lbl">NOT YET</span><span>${esc(g.h||'Not done yet.')}</span></div>`):''}
      <details class="walk"><summary>WALKTHROUGH (${q.steps.length} STEP${q.steps.length===1?'':'S'})</summary><ol class="steps">${q.steps.map(s=>`<li><span>${fmt(s)}</span></li>`).join('')}</ol></details>
      ${q.tip?`<div class="note tip"><span class="lbl">PRO TIP</span><span>${fmt(q.tip)}</span></div>`:''}
      ${q.trap?`<div class="note trap"><span class="lbl">WATCH OUT</span><span>${fmt(q.trap)}</span></div>`:''}
      ${q.answer?`<div class="spoil"><button type="button" data-spoil aria-expanded="false">Reveal answer</button><div class="ans" hidden>${fmt(q.answer)}</div></div>`:''}
      </div></article>`;
  }
  function quizItem(l,x,n){
    const key = l.id+'.'+n, pick = fs.kcPick[key], first = fs.kc[key], answered = pick!==undefined;
    return `<article class="kc-item frame" id="fkc-${l.id}-${n}"><p class="kc-q"><span class="q-num">Q${n+1}</span> ${fmt(x.q)}${first===undefined?'':` <span class="ft ${first?'ok':'no'}">${first?'FIRST TRY: RIGHT · +'+QXP+' XP':'FIRST TRY: MISSED · NO XP'}</span>`}</p>
      <div class="choices">${x.c.map((c,j)=>`<button type="button" class="choice ${answered&&j===x.a?'right':''} ${answered&&j===pick&&j!==x.a?'wrong':''}" data-fmkc="${key}" data-pick="${j}" ${answered?'disabled':''}><span class="ltr">${'ABCD'[j]}</span><span>${fmt(c)}</span></button>`).join('')}</div>
      ${answered?`<div class="note ${pick===x.a?'tip':'trap'}"><span class="lbl">${pick===x.a?'CORRECT':'NOT QUITE'}</span><span>${fmt(x.why)}</span></div><button type="button" class="retry" data-fmretry="${key}">Try this one again (practice only, no XP)</button>`:''}</article>`;
  }

  /* ---------- codex ---------- */
  function codexIndex(){
    return `<div class="crumb"><a href="#modeling">Modeling Map</a> › Codex</div><h1 class="page-h">MODELING CODEX</h1>
    <p class="lead">The textbook for this world: how the three statements connect, how forecasts and valuations work, and the traps to watch for.</p>
    <div class="path">${FM_CODEX.map((c,i)=>`<a class="node" href="#fm-cx-${i}"><span class="num" style="font-size:12px;color:var(--mint)">${c.lvl}</span><span class="nm">${esc(c.title)}</span><span class="dm">${esc(c.sub)}</span></a>`).join('')}</div>`;
  }
  function codexChapter(i){
    const c = FM_CODEX[i]; if(!c) return codexIndex(); const prev=FM_CODEX[i-1], next=FM_CODEX[i+1];
    const block = b => {
      if(b[0]==='p') return `<p>${fmt(b[1])}</p>`;
      if(b[0]==='h') return `<h2 class="fm-cxh">${esc(b[1].toUpperCase())}</h2>`;
      if(b[0]==='list') return `<ul>${b[1].map(x=>`<li>${fmt(x)}</li>`).join('')}</ul>`;
      if(b[0]==='step') return `<ol>${b[1].map(x=>`<li>${fmt(x)}</li>`).join('')}</ol>`;
      if(b[0]==='ex') return `<div class="fm-anat"><code>${esc(b[1])}</code><span>${fmt(b[2])}</span></div>`;
      if(b[0]==='tip') return `<div class="note tip"><span class="lbl">PRO TIP</span><span>${fmt(b[1])}</span></div>`;
      if(b[0]==='trap') return `<div class="note trap"><span class="lbl">WATCH OUT</span><span>${fmt(b[1])}</span></div>`;
      if(b[0]==='table') return `<div class="tbl-wrap"><table class="codes"><thead><tr>${b[1].map(h=>`<th>${esc(h.toUpperCase())}</th>`).join('')}</tr></thead><tbody>${b[2].map(r=>`<tr>${r.map(x=>`<td>${fmt(x)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
      return ''; };
    return `<div class="fmw"><div class="crumb"><a href="#fm-codex">Modeling Codex</a> › ${esc(c.title)}</div>
    <section class="lvl-head frame"><div class="lvl-tag">CODEX · FINANCIAL MODELING · ${c.lvl}</div><h1>${esc(c.title)}</h1><div class="lvl-meta">${esc(c.sub)}</div></section>
    <section class="frame box fm-cx">${c.blocks.map(block).join('')}${built(c.lvl)?`<p style="margin:14px 0 0"><a href="#fm-${c.lvl}">Practice it in Level ${c.lvl} ▶</a></p>`:''}</section>
    <nav class="lvl-nav">${prev?`<a href="#fm-cx-${i-1}">◀ ${esc(prev.title)}</a>`:'<a href="#fm-codex">◀ Codex home</a>'}${next?`<a href="#fm-cx-${i+1}">${esc(next.title)} ▶</a>`:'<a href="#fm-codex">Codex home ▶</a>'}</nav></div>`;
  }

  /* ---------- model lab ---------- */
  const BASE = {growth:[.16,.13,.10,.07,.05],gm:.32,sga:.18,da:.016,capex:.02,arDays:3,invDays:90,apDays:19,tax:.25,rate:.08,repay:20,draws:30,wacc:.14,tg:.03};
  const SCEN = {Bear:[.08,.06,.05,.04,.03], Base:[.16,.13,.10,.07,.05], Bull:[.24,.18,.14,.10,.07]};
  const freshLab = () => Object.assign({}, BASE, {growth:SCEN.Base.slice(), scen:'Base', bug:'none', invest:250, stake:.20});
  let lab = freshLab(), labTab = 'is';
  function labScreen(){
    const sl = (id,label,min,max,step,val,show) => `<label for="fml-${id}"><span class="fm-row"><span>${label}</span><b id="fmv-${id}">${show(val)}</b></span><input type="range" id="fml-${id}" data-fmlab="${id}" min="${min}" max="${max}" step="${step}" value="${val}"></label>`;
    const pct = v => (v*100).toFixed(1)+'%', days = v => v+' days', kk = v => v+'K';
    return `<div class="crumb"><a href="#modeling">Modeling Map</a> › Model Lab</div><h1 class="page-h">MODEL LAB</h1>
    <p class="lead">The finished model, running live. Every slider is a blue cell. The balance check should stay at 0.0 no matter what you change, unless you switch on a bug.</p>
    <div class="fm-lab">
      <section class="frame box fm-ctl" aria-label="Assumptions"><h3>ASSUMPTIONS</h3>
        <label for="fml-scen">Scenario<select id="fml-scen" data-fmlab="scen">${['Bear','Base','Bull','Custom'].map(s=>`<option ${s===lab.scen?'selected':''}>${s}</option>`).join('')}</select></label>
        <fieldset><legend>REVENUE GROWTH %</legend><div class="fm-grow5">${lab.growth.map((g,i)=>`<label for="fml-g${i}">FY${26+i}<input type="number" id="fml-g${i}" data-fmlab="g${i}" step="1" value="${Math.round(g*100)}"></label>`).join('')}</div></fieldset>
        ${sl('gm','Gross margin',.25,.40,.005,lab.gm,pct)}${sl('sga','SG&A % of revenue',.12,.24,.005,lab.sga,pct)}
        ${sl('invDays','Inventory days',45,150,1,lab.invDays,days)}${sl('apDays','Payable days',5,45,1,lab.apDays,days)}
        ${sl('capex','Capex % of revenue',.005,.08,.005,lab.capex,pct)}${sl('draws','Owner draws per year',0,200,5,lab.draws,kk)}
        <fieldset><legend>VALUATION</legend>${sl('wacc','Discount rate',.10,.20,.005,lab.wacc,pct)}${sl('tg','Terminal growth',0,.05,.005,lab.tg,pct)}
        ${sl('invest','Offer: investment',100,500,10,lab.invest,kk)}${sl('stake','Offer: stake',.05,.40,.01,lab.stake,v=>Math.round(v*100)+'%')}</fieldset>
        <label for="fml-bug">Bug hunt<select id="fml-bug" data-fmlab="bug">${[['none','No bug'],['da','Bug A'],['inv','Bug B'],['capex','Bug C']].map(o=>`<option value="${o[0]}" ${lab.bug===o[0]?'selected':''}>${o[1]}</option>`).join('')}</select></label>
        <p style="margin:0;font-size:14px;color:var(--muted)">Pick a bug, read how far off the check is, and name the mistake.</p>
        <button type="button" class="retry" data-fmlabreset>Reset to Base</button>
      </section>
      <div id="fm-labout" style="display:grid;gap:18px;min-width:0"></div>
    </div>`;
  }
  function runLab(d){
    const r = {rev:[1140],gp:[365],sga:[203],da:[18],ebit:[144],int:[12],tax:[33],ni:[99],ebitda:[162],cash:[60],ar:[10],inv:[190],ppe:[140],ta:[400],ap:[40],loan:[130],pic:[80],re:[150],tle:[400],chk:[0],
      cfo:[56],capex:[-128],borrow:[150],rep:[-20],draws:[-30],dcash:[28],dar:[-2],dinv:[-65],dap:[6],fcf:[null]};
    for(let t=1;t<=5;t++){ const p = k => r[k][t-1];
      const rev=p('rev')*(1+d.growth[t-1]), cogs=rev*(1-d.gm), gp=rev-cogs, sga=rev*d.sga, da=rev*d.da, ebit=gp-sga-da;
      const int=d.rate*p('loan'), ebt=ebit-int, tax=Math.max(0,ebt)*d.tax, ni=ebt-tax;
      const ar=rev*d.arDays/365, inv=cogs*d.invDays/365, ap=cogs*d.apDays/365, capex=rev*d.capex;
      const ppe = d.bug==='capex' ? p('ppe')-da : p('ppe')+capex-da;
      const loan=Math.max(0,p('loan')-d.repay), rep=-(p('loan')-loan), re=p('re')+ni-d.draws, pic=p('pic');
      const dar=-(ar-p('ar')), dinv=d.bug==='inv' ? (inv-p('inv')) : -(inv-p('inv')), dap=ap-p('ap');
      const cfo=ni+(d.bug==='da'?0:da)+dar+dinv+dap, cff=rep-d.draws, dcash=cfo-capex+cff, cash=p('cash')+dcash;
      const ta=cash+ar+inv+ppe, tl=ap+loan, te=pic+re, fcf=ebit*(1-d.tax)+da-capex-(ar-p('ar'))-(inv-p('inv'))+dap;
      const v={rev,gp,sga,da,ebit,int,tax,ni,ebitda:ebit+da,cash,ar,inv,ppe,ta,ap,loan,pic,re,tle:tl+te,chk:ta-tl-te,cfo,capex:-capex,borrow:0,rep,draws:-d.draws,dcash,dar,dinv,dap,fcf};
      for(const k in v) r[k].push(v[k]); }
    return r;
  }
  function dcf(r,w,g){ if(w-g<0.005) return null; let pv=0; for(let t=1;t<=5;t++) pv+=r.fcf[t]/Math.pow(1+w,t);
    const tv=r.fcf[5]*(1+g)/(w-g), pvtv=tv/Math.pow(1+w,5), ev=pv+pvtv; return {ev, eq:ev-70, share:pvtv/ev}; }
  function football(bars, line){
    const W=640, L=200, R=24, top=16, bh=34, gap=18, H=top+bars.length*(bh+gap)+34, lo=400, hi=2000, x=v=>L+(Math.max(lo,Math.min(hi,v))-lo)/(hi-lo)*(W-L-R);
    let s=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Valuation ranges compared with the offer">`;
    for(let t=lo;t<=hi;t+=400) s+=`<line x1="${x(t)}" x2="${x(t)}" y1="${top-6}" y2="${H-28}" stroke="#3a336e"/><text x="${x(t)}" y="${H-8}" fill="#b4abd8" text-anchor="middle">${t.toLocaleString()}</text>`;
    bars.forEach((b,i)=>{ const y=top+i*(bh+gap); s+=`<text x="${L-10}" y="${y+bh/2+5}" fill="#f1ecff" text-anchor="end">${esc(b[0])}</text><rect x="${x(b[1])}" y="${y}" width="${Math.max(2,x(b[2])-x(b[1]))}" height="${bh}" fill="${b[3]}" opacity=".85"/><text x="${x(b[1])-4}" y="${y+bh/2+5}" fill="#b4abd8" text-anchor="end">${Math.round(b[1]).toLocaleString()}</text><text x="${x(b[2])+4}" y="${y+bh/2+5}" fill="#b4abd8">${Math.round(b[2]).toLocaleString()}</text>`; });
    return s+`<line x1="${x(line)}" x2="${x(line)}" y1="${top-10}" y2="${H-28}" stroke="#f7c548" stroke-width="3" stroke-dasharray="6 4"/></svg>`;
  }
  function labOut(){
    const out = document.getElementById('fm-labout'); if(!out) return;
    const r = runLab(lab), v = dcf(r,lab.wacc,lab.tg), maxChk = Math.max(...r.chk.map(Math.abs)), ok = maxChk<0.05;
    const corners=[]; [lab.wacc-.02,lab.wacc+.02].forEach(w=>[lab.tg-.01,lab.tg+.01].forEach(g=>{ const x=dcf(r,w,Math.max(0,g)); if(x) corners.push(x.eq); })); if(v) corners.push(v.eq);
    const pre = lab.invest/lab.stake - lab.invest;
    const rows = {is:[['sec','INCOME STATEMENT'],['Revenue','rev'],['Gross profit','gp'],['SG&A','sga',-1],['D&A','da',-1],['EBIT','ebit',0,1],['Interest','int',-1],['Taxes','tax',-1],['Net income','ni',0,1],['EBITDA','ebitda']],
      bs:[['sec','ASSETS'],['Cash','cash'],['Receivables','ar'],['Inventory','inv'],['PP&E','ppe'],['Total assets','ta',0,1],['sec','LIABILITIES AND EQUITY'],['Payables','ap'],['Loan','loan'],['Paid-in capital','pic'],['Retained earnings','re'],['Liabilities + equity','tle',0,1]],
      cf:[['sec','OPERATIONS'],['Net income','ni'],['D&A','da'],['Receivables','dar'],['Inventory','dinv'],['Payables','dap'],['Cash from operations','cfo',0,1],['sec','INVESTING AND FINANCING'],['Capex','capex'],['Borrowing','borrow'],['Repayment','rep'],['Owner draws','draws'],['Net change in cash','dcash',0,1],['Ending cash','cash',0,1]]};
    const body = rows[labTab].map(x=> x[0]==='sec' ? `<tr class="sec"><td colspan="7">${x[1]}</td></tr>` : `<tr class="${x[3]?'tot':''}"><td>${esc(x[0])}</td>${r[x[1]].map((n,i)=>`<td class="${i===0?'act':''}">${k1(n==null?null:n*(x[2]===-1?-1:1))}</td>`).join('')}</tr>`).join('')
      + `<tr class="chk ${ok?'':'bad'}"><td>Balance check</td>${r.chk.map(n=>`<td>${k1(n)}</td>`).join('')}</tr>`;
    out.innerHTML = `<div class="fm-tiles">
        <div class="fm-tile ${ok?'ok':'bad'}"><span>BALANCE CHECK</span><b>${ok?'OK':'CHECK'}</b><small>${ok?'Every year balances':'Off by up to '+k1(maxChk)}</small></div>
        <div class="fm-tile"><span>FY2030 CASH</span><b>${k1(r.cash[5])}</b><small>${r.cash.some(c=>c<0)?'Goes negative: needs a credit line':'$000'}</small></div>
        <div class="fm-tile"><span>DCF EQUITY VALUE</span><b>${v?k1(v.eq):'n/a'}</b><small>${v?Math.round(v.share*100)+'% from terminal value':'Rate must beat growth'}</small></div>
        <div class="fm-tile"><span>OFFER PRE-MONEY</span><b>${k1(pre)}</b><small>${v?(pre<v.eq?'Below base case: counter':'Meets base case'):''}</small></div></div>
      <section class="frame box"><h3>THE STATEMENTS ($000)</h3>
        <div class="fm-tabs" role="tablist">${[['is','Income'],['bs','Balance sheet'],['cf','Cash flow']].map(t=>`<button type="button" role="tab" data-fmtab="${t[0]}" aria-selected="${t[0]===labTab}">${t[1]}</button>`).join('')}</div>
        <div class="tbl-wrap"><table class="fm-fs"><thead><tr><th></th>${['FY25A','FY26E','FY27E','FY28E','FY29E','FY30E'].map(y=>`<th>${y}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table></div>
        <div class="spoil"><button type="button" data-spoil aria-expanded="false">Reveal answer</button><div class="ans" hidden><b>A:</b> D&amp;A isn't added back on the cash flow statement, so cash falls short by each year's D&amp;A, and the gap keeps adding up. <b>B:</b> The inventory change has the wrong sign, so a build-up adds cash instead of using it. <b>C:</b> Capex leaves cash but never reaches PP&amp;E, so assets come up short by each year's capex.</div></div></section>
      <section class="frame box fm-ff"><h3>FOOTBALL FIELD · EQUITY VALUE ($000)</h3>${football([['Comps: EV/Revenue',614,1184,'#62e0ef'],['Comps: EV/EBITDA',1145,1712,'#a98bff'],['DCF (rate ±2%, growth ±1%)',Math.min(...corners),Math.max(...corners),'#5ef2c0']], pre)}
        <p style="margin:8px 0 0;font-size:15px;color:var(--muted)">Comps use FY2025 actuals, so they don't move with the forecast. The dashed line is the offer's pre-money value.</p></section>`;
  }

  /* ---------- save backup codes ----------
     A code holds every cleared quest (MOS, TOSA, Modeling), first-try quiz answers, practice-test bests, the Modeling unlock, and the name.
     Restoring only ADDS progress; it never removes anything already in this browser. */
  const SALT = 'xq-save-2026';
  const b64u = u => { let s=''; for(let i=0;i<u.length;i++) s+=String.fromCharCode(u[i]); return btoa(s).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,''); };
  const unb64u = t => { const s=atob(t.replace(/-/g,'+').replace(/_/g,'/')); const u=new Uint8Array(s.length); for(let i=0;i<s.length;i++) u[i]=s.charCodeAt(i); return u; };
  async function pipe(u, Kind, fmt){ const st = new Blob([u]).stream().pipeThrough(new Kind(fmt)); return new Uint8Array(await new Response(st).arrayBuffer()); }
  const canZip = () => typeof CompressionStream!=='undefined' && typeof DecompressionStream!=='undefined';
  function masks(sv, levels){ const o={}; levels.forEach(l=>{ let m=0; l.quests.forEach((q,i)=>{ if(sv[l.id+'.'+(i+1)] && i<52) m += Math.pow(2,i); }); if(m) o[l.id]=m.toString(36); }); return o; }
  function testPart(t){ return t && t.runs ? [t.runs, t.best, t.last ? [t.last.score, t.last.right, t.last.n, t.last.at] : null] : null; }
  function payload(c, name){
    const m=getLS(MKEY), t=getLS(TKEY), qf=t._qf||{}, P={v:1, n:name, at:today(), m:masks(m,c.MOS), t:masks(t,c.TOSA),
      qr:Object.keys(qf).filter(k=>qf[k]===true), qw:Object.keys(qf).filter(k=>qf[k]===false), tq:testPart(t._test), tm:testPart(t._mock), f:{}, k:[], u:fs.unlocked?1:0};
    FM_LEVELS.forEach(l=>{ const g=fs.g[l.id]||[]; let mm=0; g.forEach((x,i)=>{ if(x&&x.p) mm += Math.pow(2,i); }); if(mm) P.f[l.id]=mm.toString(36); });
    P.k = Object.keys(fs.kc).map(k=>k+(fs.kc[k]?'+':'-'));
    return P; }
  async function makeCode(c, name){
    const raw = new TextEncoder().encode(JSON.stringify(payload(c, name)));
    const z = canZip(), body = b64u(z ? await pipe(raw, CompressionStream, 'deflate-raw') : raw);
    const head = 'XQ1'+(z?'z':'p');
    return head+'.'+body+'.'+hash(SALT+'|'+head+'|'+body).slice(-6).toUpperCase(); }
  async function readCode(code){
    const parts = (code||'').replace(/\s+/g,'').split('.');
    if(parts.length!==3 || !/^XQ1[zp]$/.test(parts[0])) throw 'That doesn\'t look like an Excel Quest backup code. Copy the whole code, from XQ1 to the end.';
    if(hash(SALT+'|'+parts[0]+'|'+parts[1]).slice(-6).toUpperCase()!==parts[2].toUpperCase()) throw 'That code has been changed or wasn\'t copied completely. Copy it again from your backup.';
    let u; try { u = unb64u(parts[1]); } catch(e){ throw 'That code wasn\'t copied completely. Copy it again from your backup.'; }
    if(parts[0][3]==='z'){ if(!canZip()) throw 'This browser is too old to read that code. Update it, or try Chrome.'; u = await pipe(u, DecompressionStream, 'deflate-raw'); }
    const P = JSON.parse(new TextDecoder().decode(u)); if(!P || P.v!==1) throw 'That backup code is from a different version of Excel Quest.';
    return P; }
  const bits = s => { let n=parseInt(s,36), out=[], i=0; while(n>0){ if(n%2) out.push(i); n=Math.floor(n/2); i++; } return out; };
  function summary(P, c){
    const lvStars = (mm, levels, sv) => levels.filter(l=>{ const set=new Set(bits(mm[l.id]||'0')); return l.quests.every((q,i)=>!XQ.gradable(l.id,i) || set.has(i)); }).length;
    const quests = mm => Object.values(mm).reduce((a,s)=>a+bits(s).length,0);
    return {name:P.n, at:P.at, mos:`${quests(P.m)} quests, ${lvStars(P.m,c.MOS)}/${c.MOS.length} levels starred`, tosa:`${quests(P.t)} quests, ${lvStars(P.t,c.TOSA)}/${c.TOSA.length} levels starred, ${P.qr.length} quiz answers right on the first try`,
      fm:`${P.u?'unlocked':'locked'}${Object.keys(P.f).length?', '+quests(P.f)+' quests':''}`}; }
  function merge(P, c){
    const m=getLS(MKEY), t=getLS(TKEY); let added=0;
    const put = (sv, mm) => Object.keys(mm).forEach(id=>bits(mm[id]).forEach(i=>{ const k=id+'.'+(i+1); if(!sv[k]){ sv[k]=1; added++; } }));
    put(m, P.m); put(t, P.t);
    t._qf = t._qf || {}; P.qr.forEach(id=>{ if(t._qf[id]===undefined){ t._qf[id]=true; added++; } }); P.qw.forEach(id=>{ if(t._qf[id]===undefined) t._qf[id]=false; });
    const tst = (key, a) => { if(!a) return; const cur=t[key]||{runs:0,best:0}; if(!cur.runs || a[1]>cur.best){ t[key]={runs:Math.max(cur.runs||0,a[0]), best:Math.max(cur.best||0,a[1]), last: cur.last || (a[2]?{score:a[2][0],right:a[2][1],n:a[2][2],at:a[2][3]}:null)}; } else cur.runs=Math.max(cur.runs,a[0]); };
    tst('_test', P.tq); tst('_mock', P.tm);
    Object.keys(P.f).forEach(id=>{ fs.g[id]=fs.g[id]||[]; bits(P.f[id]).forEach(i=>{ if(!(fs.g[id][i]&&fs.g[id][i].p)){ fs.g[id][i]={p:1}; added++; } }); for(let i=0;i<fs.g[id].length;i++) if(fs.g[id][i]===undefined) fs.g[id][i]=null; });
    P.k.forEach(x=>{ const key=x.slice(0,-1); if(fs.kc[key]===undefined) fs.kc[key]=x.slice(-1)==='+'; });
    if(P.u && !fs.unlocked){ fs.unlocked='restored'; added++; }
    if(!getName() && P.n) setName(P.n);
    fs.lastBackup = fs.lastBackup || P.at;
    setLS(MKEY, m); setLS(TKEY, t); save();
    return added; }
  function ago(d){ if(!d) return 'never'; const n=Math.round((Date.parse(today())-Date.parse(d))/86400000); return n<=0?'today':n===1?'yesterday':n+' days ago'; }
  function backupScreen(c){
    return `<div class="crumb"><a href="#map">World Map</a> › Save Backup</div><h1 class="page-h">SAVE BACKUP</h1>
    <p class="lead">Your progress is saved in this browser only. A new computer, a different browser, a private window, or clearing your history can erase it. A backup code is a copy of everything: MOS, TOSA, quiz answers, practice tests, and the Modeling world. Make a new one at the end of every session and before every break.</p>
    <section class="frame box fm-form"><h3>MAKE A BACKUP CODE</h3>
      <p style="margin:0;color:var(--muted);font-size:15px">Last backup: <b id="fm-last">${ago(fs.lastBackup)}</b></p>
      <label for="fm-bname">Your full name<input id="fm-bname" type="text" autocomplete="name" value="${esc(getName())}" placeholder="First and last name"></label>
      <button type="button" class="press fm-press" id="fm-make" style="margin-top:0;justify-self:start">MAKE MY CODE</button>
      <div id="fm-codebox" hidden><label for="fm-code-out">Your backup code<textarea id="fm-code-out" rows="4" readonly></textarea></label>
        <p style="margin:0;display:flex;gap:10px;flex-wrap:wrap"><button type="button" class="retry" id="fm-copy" style="margin:0">Copy code</button><a class="retry" id="fm-save" style="margin:0;text-decoration:none" download="Excel_Quest_backup.txt" href="#">Save as a file</a></p>
        <p style="margin:0;color:var(--muted);font-size:15px">Paste it into a Google Doc or email it to yourself. Anyone with your code gets your progress, so keep it to yourself.</p></div>
      <p class="turnin-msg" id="fm-bmsg" role="status"></p></section>
    <section class="frame box fm-form" style="margin-top:22px"><h3>RESTORE FROM A CODE</h3>
      <p style="margin:0;color:var(--muted);font-size:15px">Restoring adds the progress in the code to this browser. It never erases anything already here.</p>
      <label for="fm-code-in">Paste your backup code<textarea id="fm-code-in" rows="4" autocomplete="off" spellcheck="false"></textarea></label>
      <button type="button" class="press fm-press" id="fm-check" style="margin-top:0;justify-self:start">CHECK THE CODE</button>
      <div id="fm-preview"></div><p class="turnin-msg" id="fm-imsg" role="status"></p></section>`;
  }
  const sumHTML = S => `<div class="fm-sum"><p><b>${esc(S.name||'No name')}</b> · backed up ${esc(S.at||'')}</p><ul><li>MOS: ${esc(S.mos)}</li><li>TOSA: ${esc(S.tosa)}</li><li>Modeling: ${esc(S.fm)}</li></ul></div>`;
  function unlockCode(name){ const n=normName(name); if(!n) return ''; const h=hash('xq-fm-unlock-2026|'+n); const A='23456789ABCDEFGHJKMNPQRSTUVWXYZ'; let x=parseInt(h.slice(-10),16), out='';
    for(let i=0;i<8;i++){ out+=A[x%A.length]; x=Math.floor(x/A.length); } return out.slice(0,4)+'-'+out.slice(4); }
  function normName(s){ return (s||'').toLowerCase().replace(/[^a-z' -]/g,'').replace(/\s+/g,' ').trim(); }

  /* ---------- teacher preview ---------- */
  function teacherTools(){
    return `<div class="crumb"><a href="#modeling">Modeling</a> › Teacher tools</div><h1 class="page-h">TEACHER TOOLS</h1>
    <p class="lead">Preview is on for this browser tab. Closing the tab locks it again.</p>
    <div class="cards">
      <section class="frame box fm-form"><h3>PREVIEW THE WORLD</h3><p style="margin:0">See the Financial Modeling world the way an unlocked student does.</p><a class="press fm-press" href="#modeling" style="margin-top:0;justify-self:start">OPEN MODELING</a></section>
      <section class="frame box fm-form"><h3>UNLOCK CODE FOR A STUDENT</h3><p style="margin:0;font-size:15px">Use it when a student's progress was lost and your records show they finished. Each code works only with the name it was made for.</p>
        <label for="fm-tname">Student's full name<input id="fm-tname" type="text" autocomplete="off"></label>
        <p class="fm-ucode" id="fm-ucode" aria-live="polite"></p></section>
    </div>
    <section class="frame box fm-form" style="margin-top:22px"><h3>CHECK A STUDENT'S BACKUP CODE</h3><p style="margin:0;font-size:15px">Paste a code a student sent you to see what it holds. Nothing is restored into this browser.</p>
      <label for="fm-tcode">Backup code<textarea id="fm-tcode" rows="3" spellcheck="false"></textarea></label>
      <button type="button" class="retry" id="fm-tcheck" style="margin:0;justify-self:start">Check it</button><div id="fm-tsum"></div><p class="turnin-msg" id="fm-tmsg" role="status"></p></section>`;
  }
  function teacherScreen(){
    if(preview()) return teacherTools();
    return `<div class="crumb"><a href="#modeling">Modeling</a> › Teacher preview</div><h1 class="page-h">TEACHER PREVIEW</h1>
    <p class="lead">Opens the Financial Modeling world in this browser tab only. Closing the tab locks it again.</p>
    <form class="frame box fm-pass" id="fm-pass"><label for="fm-code">Passcode</label><input type="password" id="fm-code" autocomplete="off"><button type="submit" class="press fm-press" style="margin-top:0">UNLOCK</button><p class="turnin-msg" id="fm-passmsg" role="status"></p></form>`;
  }

  /* ---------- public API ---------- */
  const owns = h => h==='modeling' || h==='backup' || h.startsWith('fm-');
  function screen(h, c){
    ctx = c;
    stick(c);
    if(h==='fm-teacher') return teacherScreen();
    if(h==='backup') return backupScreen(c);
    if(!open(c)) return gateScreen(c);
    if(h==='modeling') return mapScreen(c);
    if(h==='fm-lab') return labScreen();
    if(h==='fm-codex') return codexIndex();
    if(h.startsWith('fm-cx-')) return codexChapter(+h.slice(6));
    return levelScreen(h.slice(3)) || mapScreen(c);
  }
  function after(h, c){
    if(c) ctx = c;
    const nav = document.querySelector('[data-nav="modeling"]');
    if(nav && ctx) nav.textContent = open(ctx) ? 'Modeling' : '🔒 Modeling';
    if(!owns(h)) return;
    const xp = document.getElementById('xp'), cl = document.getElementById('cleared');
    if(ctx && open(ctx) && h!=='fm-teacher' && h!=='backup'){ if(xp) xp.textContent = fmXP().toLocaleString();
      if(cl) cl.textContent = FM_LEVELS.filter(l=>built(l.id) && fmStats(l).done===fmStats(l).total).length+'/'+FM_LEVELS.length; }
    if(h==='fm-lab') labOut();
  }
  function backupCard(){ const old = !fs.lastBackup || ago(fs.lastBackup)==='never' || (Date.parse(today())-Date.parse(fs.lastBackup))/86400000 > 7;
    return `<a class="frame box fm-teaser fm-bk ${old?'old':''}" href="#backup"><span class="fm-k">SAVE BACKUP · LAST BACKUP: ${esc(ago(fs.lastBackup).toUpperCase())}</span><span>Your progress lives in this browser only. Make a backup code so a new computer or cleared history can't erase it. ▶</span></a>`; }
  function mapCard(c){
    ctx = c; stick(c); const p = progress(c);
    if(open(c)) return backupCard() + `<a class="frame box fm-teaser" href="#modeling"><span class="fm-k">BONUS WORLD · UNLOCKED</span><span class="fm-nm">FINANCIAL MODELING</span><span>Three statements, forecasting, comps, and DCF for Magic City Sneaker Exchange. ▶ Enter</span></a>`;
    return backupCard() + `<a class="frame box fm-teaser locked" href="#modeling"><span class="fm-k">🔒 BONUS WORLD · ${p.done}/${p.total} LEVELS STARRED</span><span class="fm-nm">FINANCIAL MODELING</span><span>Star every MOS and TOSA level to unlock it.</span><span class="bar fm-bar"><i style="width:${p.total?Math.round(100*p.done/p.total):0}%"></i></span></a>`;
  }
  function codes(c){
    if(c) ctx = c;
    if(!ctx || !open(ctx)) return `<section class="frame box" style="margin-top:24px"><h3 style="color:var(--muted)">🔒 FINANCIAL MODELING CODES</h3><p style="margin:0">${FM_CODES.length} more codes unlock with the Financial Modeling world. <a href="#modeling">See what's left</a>.</p></section>`;
    return `<h2 class="section-h" style="margin-top:32px;color:var(--mint)">FINANCIAL MODELING</h2><div class="frame box tbl-wrap"><table class="codes"><thead><tr><th>CODE</th><th>WHAT IT DOES</th><th>USED IN</th></tr></thead><tbody>${FM_CODES.map(c=>`<tr><td>${fmt(c[0])}</td><td>${fmt(c[1])}</td><td>${esc(c[2])}</td></tr>`).join('')}</tbody></table></div>`;
  }
  const rerender = () => window.dispatchEvent(new HashChangeEvent('hashchange'));

  async function turnIn(file){
    const lid = (location.hash||'').replace('#fm-',''), msg = document.getElementById('fm-msg'); if(!file || !msg) return;
    if(!/\.xlsx$/i.test(file.name)){ msg.className='turnin-msg err'; msg.textContent="That isn't an Excel workbook. Turn in the .xlsx you saved."; return; }
    if(typeof JSZip==='undefined'){ msg.className='turnin-msg err'; msg.textContent="The grader couldn't load. Check your connection and reload the page."; return; }
    msg.className='turnin-msg'; msg.textContent='Grading '+file.name+'...';
    try {
      const res = await XQ.grade(lid, await file.arrayBuffer());
      fs.g[lid] = res.map(r=>r===null?null:(r.pass?{p:1}:{p:0,h:r.hint})); save(); rerender();
      setTimeout(()=>{ const l=FM_LEVELS.find(x=>x.id===lid), s=fmStats(l), m2=document.getElementById('fm-msg'); if(!m2) return;
        m2.className='turnin-msg '+(s.done===s.total?'ok':''); m2.textContent = s.done===s.total ? `All ${s.total} quests passed. Level cleared!` : `${s.done} of ${s.total} quests passed. The list below shows what to fix. Fix them in Excel, save, and turn it in again.`;
        if(res.calcManual){ const w=document.createElement('p'); w.className='turnin-msg err'; w.textContent='Heads up: this file was saved with Manual calculation, so some numbers may be out of date. In Excel choose Formulas, Calculation Options, Automatic. Then save and turn it in again.'; m2.after(w); }
        document.getElementById('fm-turnin').scrollIntoView({block:'start'}); }, 0);
    } catch(e){ msg.className='turnin-msg err'; msg.textContent = (e && e.msg) || 'Something went wrong reading that file.'; }
  }

  document.addEventListener('click', e=>{
    const kc = e.target.closest('[data-fmkc]');
    if(kc){ const key=kc.dataset.fmkc, pick=+kc.dataset.pick, lid=key.split('.')[0], n=+key.split('.')[1], l=FM_LEVELS.find(x=>x.id===lid), x=l.quiz[n];
      fs.kcPick[key]=pick; if(fs.kc[key]===undefined) fs.kc[key]=(pick===x.a); save();
      document.getElementById('fkc-'+lid+'-'+n).outerHTML = quizItem(l,x,n); after(location.hash.slice(1)); return; }
    const kr = e.target.closest('[data-fmretry]');
    if(kr){ const key=kr.dataset.fmretry, lid=key.split('.')[0], n=+key.split('.')[1], l=FM_LEVELS.find(x=>x.id===lid); delete fs.kcPick[key]; save();
      document.getElementById('fkc-'+lid+'-'+n).outerHTML = quizItem(l,l.quiz[n],n); return; }
    const tb = e.target.closest('[data-fmtab]'); if(tb){ labTab = tb.dataset.fmtab; labOut(); return; }
    if(e.target.closest('[data-fmlabreset]')){ lab = freshLab(); rerender(); return; }
    if(e.target.closest('[data-fmexit]')){ try { sessionStorage.removeItem(PKEY); } catch(err){} rerender(); return; }
  });
  document.addEventListener('input', e=>{
    const el = e.target; if(!el.dataset || !el.dataset.fmlab) return; const k = el.dataset.fmlab;
    if(k==='scen'){ if(SCEN[el.value]){ lab.growth=SCEN[el.value].slice(); lab.scen=el.value; lab.growth.forEach((g,i)=>{ const x=document.getElementById('fml-g'+i); if(x) x.value=Math.round(g*100); }); } else lab.scen='Custom'; }
    else if(k==='bug') lab.bug = el.value;
    else if(/^g\d$/.test(k)){ lab.growth[+k[1]] = (parseFloat(el.value)||0)/100; lab.scen='Custom'; const s=document.getElementById('fml-scen'); if(s) s.value='Custom'; }
    else { lab[k]=parseFloat(el.value); const v=document.getElementById('fmv-'+k); if(v) v.textContent = /Days$/.test(k) ? lab[k]+' days' : (k==='draws'||k==='invest') ? lab[k]+'K' : k==='stake' ? Math.round(lab[k]*100)+'%' : (lab[k]*100).toFixed(1)+'%'; }
    labOut();
  });
  document.addEventListener('change', e=>{ if(e.target.id==='fm-file'){ turnIn(e.target.files[0]); e.target.value=''; } });
  document.addEventListener('dragover', e=>{ const d=e.target.closest && e.target.closest('#fm-drop'); if(d){ e.preventDefault(); d.classList.add('over'); } });
  document.addEventListener('dragleave', e=>{ const d=e.target.closest && e.target.closest('#fm-drop'); if(d) d.classList.remove('over'); });
  document.addEventListener('drop', e=>{ const d=e.target.closest && e.target.closest('#fm-drop'); if(d){ e.preventDefault(); d.classList.remove('over'); turnIn(e.dataTransfer.files[0]); } });
  document.addEventListener('submit', e=>{ if(e.target.id!=='fm-pass') return; e.preventDefault();
    const m=document.getElementById('fm-passmsg'), v=(document.getElementById('fm-code').value||'').trim();
    if(hash('xq-fm|'+v)===PASS_HASH){ try { sessionStorage.setItem(PKEY,'1'); } catch(err){} rerender(); }
    else { m.className='turnin-msg err'; m.textContent='That passcode isn\'t right.'; } });

  let pending = null;
  document.addEventListener('click', async e=>{
    const t = e.target;
    if(t.id==='fm-make'){ const msg=document.getElementById('fm-bmsg'), name=(document.getElementById('fm-bname').value||'').trim().replace(/\s+/g,' ');
      if(name.split(' ').length<2){ msg.className='turnin-msg err'; msg.textContent='Type your first and last name first.'; return; }
      setName(name); const code = await makeCode(ctx, name); fs.lastBackup = today(); save();
      document.getElementById('fm-code-out').value = code; document.getElementById('fm-codebox').hidden = false;
      const a=document.getElementById('fm-save'); a.href = URL.createObjectURL(new Blob(['Excel Quest backup for '+name+' ('+today()+')\n\n'+code+'\n\nTo restore: open Excel Quest, go to Save Backup, and paste the code.\n'],{type:'text/plain'}));
      a.download = 'Excel_Quest_backup_'+name.replace(/[^A-Za-z]+/g,'_')+'_'+today()+'.txt';
      document.getElementById('fm-last').textContent='today'; msg.className='turnin-msg ok'; msg.textContent='Code ready. Copy it or save it as a file now.'; return; }
    if(t.id==='fm-copy'){ const ta=document.getElementById('fm-code-out'), msg=document.getElementById('fm-bmsg');
      try { await navigator.clipboard.writeText(ta.value); msg.className='turnin-msg ok'; msg.textContent='Copied. Paste it somewhere safe.'; }
      catch(err){ ta.focus(); ta.select(); msg.className='turnin-msg'; msg.textContent='Press Cmd+C to copy the selected code.'; } return; }
    if(t.id==='fm-check'){ const msg=document.getElementById('fm-imsg'), box=document.getElementById('fm-preview'); box.innerHTML='';
      try { pending = await readCode(document.getElementById('fm-code-in').value); const S=summary(pending, ctx);
        box.innerHTML = sumHTML(S)+`<button type="button" class="press fm-press" id="fm-restore" style="margin-top:6px">ADD THIS PROGRESS</button>`; msg.className='turnin-msg'; msg.textContent=''; }
      catch(err){ pending=null; msg.className='turnin-msg err'; msg.textContent = typeof err==='string' ? err : 'That code couldn\'t be read. Copy it again from your backup.'; } return; }
    if(t.id==='fm-restore' && pending){ const n = merge(pending, ctx); pending=null; const msg=document.getElementById('fm-imsg');
      msg.className='turnin-msg ok'; msg.textContent = n ? `Restored ${n} item${n===1?'':'s'}. Reloading...` : 'Everything in that code was already here. Nothing to add.';
      if(n) setTimeout(()=>location.reload(), 700); return; }
    if(t.id==='fm-tcheck'){ const msg=document.getElementById('fm-tmsg'), box=document.getElementById('fm-tsum'); box.innerHTML='';
      try { const P = await readCode(document.getElementById('fm-tcode').value); box.innerHTML = sumHTML(summary(P, ctx)); msg.textContent=''; }
      catch(err){ msg.className='turnin-msg err'; msg.textContent = typeof err==='string' ? err : 'That code couldn\'t be read.'; } return; }
  });
  document.addEventListener('input', e=>{ if(e.target.id==='fm-tname'){ const c=unlockCode(e.target.value), o=document.getElementById('fm-ucode'); o.innerHTML = c ? `Code for <b>${esc(e.target.value.trim())}</b>: <span>${c}</span>` : ''; } });
  document.addEventListener('submit', e=>{ if(e.target.id!=='fm-redeem') return; e.preventDefault();
    const name=(document.getElementById('fm-rname').value||'').trim(), code=(document.getElementById('fm-rcode').value||'').toUpperCase().replace(/[^0-9A-Z]/g,''), m=document.getElementById('fm-rmsg');
    if(normName(name).split(' ').length<2){ m.className='turnin-msg err'; m.textContent='Type your first and last name, the same way your teacher did.'; return; }
    if(unlockCode(name).replace('-','')!==code){ m.className='turnin-msg err'; m.textContent='That code doesn\'t match that name. Check both with your teacher.'; return; }
    fs.unlocked='teacher'; fs.unlockedAt=today(); save(); if(!getName()) setName(name.replace(/\s+/g,' ')); rerender(); });

  return {owns, screen, after, mapCard, codes, _hash: hash, _unlockCode: unlockCode};
})();
