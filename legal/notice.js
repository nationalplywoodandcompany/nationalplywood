const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  LevelFormat, convertInchesToTwip
} = require('docx');
const fs = require('fs');

const FONT = 'Calibri';
const BLANK = (t) => new TextRun({ text: t, highlight: 'yellow' });

const p = (text, opts = {}) => new Paragraph({
  spacing: { after: opts.after ?? 120, line: 276 },
  alignment: opts.align,
  indent: opts.indent,
  children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, size: opts.size ?? 22, font: FONT })],
});

const rich = (children, opts = {}) => new Paragraph({
  spacing: { after: opts.after ?? 120, line: 276 },
  alignment: opts.align,
  indent: opts.indent,
  children,
});

const t = (text, o = {}) => new TextRun({ text, bold: o.b, italics: o.i, size: o.size ?? 22, font: FONT, underline: o.u ? {} : undefined });

const head = (text) => new Paragraph({
  spacing: { before: 260, after: 140 },
  children: [new TextRun({ text, bold: true, size: 24, font: FONT })],
});

const rule = () => new Paragraph({
  spacing: { before: 100, after: 160 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '999999', space: 1 } },
  children: [new TextRun({ text: '', size: 2, font: FONT })],
});

// ---- Chronology table -------------------------------------------------
const CHRON = [
  ['19.07.2025', 'In the course of negotiations the Noticee represented in writing on WhatsApp: \u201c40% advance / 12-13 days delivery\u201d.'],
  ['25.07.2025', 'On being asked whether there would be any complaint as to quality, the Noticee represented in writing: \u201cSir humlog 3000-3500 monthly banate hai, aisse toh koi dikkat aata nai.\u201d'],
  ['01.06.2026', 'The Noticee quoted MEMBRANE HW Rs 105, LAMINATION Rs 112, SIDE PINE + Rs 15. No cutting charge, loading charge or any other charge was mentioned.'],
  ['02.06.2026', 'My client and his customer attended the Noticee\u2019s office. Rs 20,000 was paid there in cash to M/s BG Timber. You thereupon wrote out, in your own hand and on a plain sheet of paper bearing no printed particulars of any firm, a memorandum of the terms in the following words: \u201cKolkata Plywood 2/6/26 \u2014 (1) BGL 10 Lamination 140/- (2) BG 26 Membrane Rosewood 135/- \u2014 + Delivery Jhargram \u2014 + 50% Bill \u2014 2/6 20000 Advance \u2014 Total 40% Advance\u201d, and required my client to sign it in acknowledgement of the cash paid. You retained the original in your own custody and did not hand it over, permitting my client only to photograph it. My client transmitted that photograph to you on WhatsApp the same day at 17:19 hrs with the endorsement \u201c20000/- cash Paid (As a Advance) To BG timber From Kolkata Plywood & Co.\u201d You did not then, or at any time in the eleven weeks since, dispute a single word of it. DELIVERY AT JHARGRAM WAS THUS PART OF THE CONTRACT.'],
  ['03.06.2026', 'The Noticee furnished the account of M/s Wood and Wood Products, HDFC Bank A/c 50200095517481, IFSC HDFC0004339, and asked for a transfer of Rs 80,000.'],
  ['04.06.2026', 'My client furnished the complete size list of all 94 doors on WhatsApp \u2014 25 pcs membrane and 69 pcs lamination. That list is the contractual specification.'],
  ['09.06.2026', 'Rs 80,000 transferred to the said account of M/s Wood and Wood Products.'],
  ['22.06.2026', 'My client recorded in writing: \u201c13-Days ho gya hai Payment kiye ?? 10-15din Maximum bole thy\u2026\u201d No delivery was made.'],
  ['08.07.2026', 'The Noticee furnished the details of two savings accounts of third parties \u2014 a PhonePe collect code of Sri Ramkrishna Mondal and the number +91 90621 80929 of Sri Ajay Sharma \u2014 against which Rs 50,000 and Rs 20,000 respectively were paid the same evening.'],
  ['11.07.2026', 'The Noticee again nominated the account of Sri Ajay Sharma and Rs 51,200 was transferred. The goods were despatched. The tax invoice sent was made out in the name of a stranger, \u201cASF PLYWOOD, 163 Ishwaripur, PS Rahara\u201d, and not in the name of my client.'],
  ['12.07.2026', 'The goods reached the customer\u2019s site at Jhargram \u2014 37 days after the order. My client pointed out at 11:18 hrs that the invoice was in the name of ASF Plywood. The Noticee replied that a revised invoice would be sent.'],
  ['13.07.2026', 'Revised tax invoice no. 18/WWP/26-27 dated 13-07-2026 issued, billed to my client and shipped to M/s Kriti Construction, Jhargram, along with e-way bill no. 8517 1317 0306 and motor vehicle no. WB19S0857.'],
  ['16.07.2026', 'On inspection my client recorded the defects in writing: of the 69 pcs lamination ordered (all widths within 32\u2033), 10 pcs had been manufactured 36\u2033/36.5\u2033 wide, and 11 pcs had been supplied in MEMBRANE where LAMINATION had been ordered. Several lamination doors were supplied with one-side design although both-side design had been agreed, and the thickness was not uniform.'],
  ['17.07.2026', 'The Noticee admitted the defect in writing: \u201cAur isko return karwa dijiye, change karke bhejwa denge.\u201d On being asked who would bear the freight and by when the replacement would come, the Noticee replied only \u201cJab ayega usko dhek ke Bata panga\u201d.'],
  ['20.07.2026', 'My client sent five pages of measurements comparing the ordered size against the size actually supplied, recording that the sizes were out by 2 suta to 1 inch throughout, and that two carpenters had spent a full day measuring. He recorded that the customer would bear no extra cost and that the freight of Rs 7,000 each way arose from the Noticee\u2019s default.'],
  ['29.07.2026', 'The 21 defective doors were despatched to the Noticee\u2019s factory at my client\u2019s own cost of Rs 7,000, the Noticee having refused to bear it.'],
  ['30.07.2026', 'The goods reached the factory and were measured and noted by the Noticee\u2019s own man. The Noticee stated: \u201cThis will take time \u2014 Will rectify and send \u2014 Cannot confirm exact time now.\u201d My client called upon the Noticee to refund the price of the 21 pcs if they could not be delivered by Tuesday. No refund was made.'],
  ['31.07.2026', 'My client had set out for the factory to inspect. The Noticee messaged \u201cNo need to come.\u201d'],
  ['07.08.2026 & 11.08.2026', 'My client\u2019s customer abused him and attended his office. On being pressed, the Noticee replied: \u201cThen wait kariye\u201d, \u201cI don\u2019t care. Mereko mat boliye\u201d and \u201cBut koi bakwass nai sunuenga\u201d.'],
  ['12.08.2026 to 21.08.2026', 'The ready-date was shifted from Friday, to \u201cTomorrow evening\u201d, to Monday (confirmed \u201cPakka Final Na??\u201d \u2014 \u201cYes.\u201d), and then abandoned on 17.08.2026 with \u201cNai. Weather dhek rhe hai?\u201d.'],
  ['21.08.2026', 'The Noticee stated the factory opens at 8:30, the manager comes by 9 and loading would be around 10:00\u201310:30. My client confirmed in writing that he would place a vehicle at 10:00 a.m. and asked that the challan be prepared.'],
  ['22.08.2026, morning', 'The Noticee resiled and stated loading would happen only after lunch, at 2:00 p.m., adding \u201cBaad mein hum kuch nai sunega\u201d.'],
  ['22.08.2026, 2:00 p.m.', 'My client attended with a hired vehicle at the appointed hour. Not one of the 21 doors was complete. 18 of them were finished on one side only and were still under manufacture. The vehicle was detained 4 hours 06 minutes, of which 156 minutes were charged as waiting time (Trip CRN113083391386, vehicle WB-01-BB-8397). At 17:03 hrs my client recorded that he had been waiting three hours without lunch and again asked for the challan.'],
  ['22.08.2026, 5:00 p.m. onwards', 'The 18 pieces were completed only at about 5:00 p.m., after which the remaining 3 were taken up. It then emerged that the skin which the customer had selected was exhausted at your factory, so that my client had to photograph an alternative skin, send it to his customer on WhatsApp and obtain his approval on the spot before manufacture could proceed. The goods were finished at about 6:00 p.m.'],
  ['22.08.2026, evening', 'The goods were released at about 6:40 p.m. WITHOUT ANY DOCUMENT \u2014 no delivery challan, no invoice, no e-way bill \u2014 notwithstanding my client\u2019s written request of 21.08.2026 and his repeated requests on the day. On inspection the replacement doors were themselves found defective, the resin and skin not having adhered in several pieces, leaving the core paper visible.'],
];


const chronRows = [
  new TableRow({
    tableHeader: true,
    children: ['Date', 'Event'].map((h, i) => new TableCell({
      width: { size: i === 0 ? 2100 : 7200, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: 'E8E8E8' },
      margins: { top: 80, bottom: 80, left: 110, right: 110 },
      children: [rich([t(h, { b: true })], { after: 0 })],
    })),
  }),
  ...CHRON.map(([d, e]) => new TableRow({
    children: [
      new TableCell({
        width: { size: 2100, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 110, right: 110 },
        children: [rich([t(d, { b: true })], { after: 0 })],
      }),
      new TableCell({
        width: { size: 7200, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 110, right: 110 },
        children: [rich([t(e)], { after: 0 })],
      }),
    ],
  })),
];

// ---- Damages table ----------------------------------------------------
const DAMAGES = [
  ['1.', 'Freight paid on 29.07.2026 for returning the 21 defective doors to the Noticee\u2019s factory', '7,000'],
  ['2.', 'Freight paid on 22.08.2026 for re-transporting the replaced doors (Trip CRN113083391386, vehicle WB-01-BB-8397), inclusive of the waiting charge at item 3', '6,940'],
  ['3.', 'of which vehicle waiting charges \u2014 156 minutes at Rs 3.50 per minute, the vehicle having been detained at the Noticee\u2019s factory for 4 hours 06 minutes', '(546)'],
  ['4.', 'Charges paid to my client\u2019s own carpenter for two days\u2019 re-measurement and segregation of the consignment', '1,000'],
  ['5.', 'Cutting and re-fitting charges for 73 doors at Rs 100 per door, paid to the customer\u2019s carpenters, the customer having otherwise required the entire consignment to be taken back', '7,300'],
  ['6.', 'Cutting charge of Rs 2 per sq. ft. on 1622.907 sq. ft., introduced for the first time on the date of despatch', '3,246'],
  ['7.', 'Loading charge, introduced for the first time on the date of despatch', '1,000'],
  ['8.', 'Differential of Rs 10 per sq. ft. on 228.79 sq. ft. re-rated at Rs 132 in place of Rs 122, introduced for the first time on the date of despatch', '2,288'],
  ['9.', 'Travel, lodging and incidental expenses at Kolkata on 09.07.2026 to 11.07.2026 and on 22.08.2026', '[AMOUNT]'],
];



const dmgRow = (cells, o = {}) => new TableRow({
  children: [
    new TableCell({ width: { size: 700, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 110, right: 110 },
      shading: o.head ? { type: ShadingType.CLEAR, fill: 'E8E8E8' } : undefined,
      children: [rich([t(cells[0], { b: o.head || o.b })], { after: 0 })] }),
    new TableCell({ width: { size: 6800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 110, right: 110 },
      shading: o.head ? { type: ShadingType.CLEAR, fill: 'E8E8E8' } : undefined,
      children: [rich([t(cells[1], { b: o.head || o.b })], { after: 0 })] }),
    new TableCell({ width: { size: 1800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 110, right: 110 },
      shading: o.head ? { type: ShadingType.CLEAR, fill: 'E8E8E8' } : undefined,
      children: [rich([t(cells[2], { b: o.head || o.b })], { after: 0, align: AlignmentType.RIGHT })] }),
  ],
});

const damageRows = [
  dmgRow(['Sl.', 'Particulars', 'Amount (Rs.)'], { head: true }),
  ...DAMAGES.map(r => dmgRow(r)),
  dmgRow(['', 'Total quantified out-of-pocket loss (item 3 being included within item 2, and excluding item 9)', '28,774'], { b: true }),
];


// ---- Payment schedule table -------------------------------------------
const PAYMENTS = [
  ['02.06.2026', 'Cash', 'Paid in hand at the Noticee\u2019s office to M/s BG Timber', '20,000'],
  ['09.06.2026', 'Bank transfer', 'Current account of M/s Wood and Wood Products, HDFC Bank A/c 50200095517481', '80,000'],
  ['08.07.2026', 'Bank transfer', 'Savings account of Sri Ajay Sharma, nominated by the Noticee at 15:41 hrs', '20,000'],
  ['08.07.2026', 'Bank transfer', 'Savings account of Sri Ramkrishna Mondal, nominated by the Noticee at 15:40 hrs', '50,000'],
  ['11.07.2026', 'Bank transfer', 'Savings account of Sri Ajay Sharma, nominated by the Noticee at 14:39 hrs', '51,200'],
];


const payCell = (txt, w, o = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA },
  margins: { top: 80, bottom: 80, left: 110, right: 110 },
  shading: o.head ? { type: ShadingType.CLEAR, fill: 'E8E8E8' } : undefined,
  children: [rich([t(txt, { b: o.head || o.b })], { after: 0, align: o.right ? AlignmentType.RIGHT : undefined })],
});

const payRows = [
  new TableRow({ tableHeader: true, children: [
    payCell('Date', 1600, { head: true }), payCell('Mode', 2100, { head: true }),
    payCell('Paid into', 3900, { head: true }), payCell('Amount (Rs.)', 1700, { head: true, right: true }) ] }),
  ...PAYMENTS.map(r => new TableRow({ children: [
    payCell(r[0], 1600), payCell(r[1], 2100), payCell(r[2], 3900), payCell(r[3], 1700, { right: true }) ] })),
  new TableRow({ children: [
    payCell('', 1600), payCell('', 2100), payCell('Total consideration paid', 3900, { b: true }),
    payCell('2,21,200', 1700, { b: true, right: true }) ] }),
];

// ---- Document ---------------------------------------------------------
const doc = new Document({
  numbering: {
    config: [{
      reference: 'demands',
      levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '(%1)', alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
    }, {
      reference: 'bullets',
      levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
    }],
  },
  styles: { default: { document: { run: { font: FONT, size: 22 } } } },
  sections: [{
    properties: { page: { margin: { top: convertInchesToTwip(1), bottom: convertInchesToTwip(1), left: convertInchesToTwip(1), right: convertInchesToTwip(1) } } },
    children: [
      rich([t('DRAFT FOR ADVOCATE — NOT TO BE SENT AS IS', { b: true, size: 20 })], { align: AlignmentType.CENTER, after: 60 }),
      rich([t('All highlighted fields must be verified and filled in before issue.', { i: true, size: 18 })], { align: AlignmentType.CENTER, after: 200 }),
      rule(),

      rich([t('LEGAL NOTICE', { b: true, size: 28, u: true })], { align: AlignmentType.CENTER, after: 100 }),
      rich([t('(Under instructions from and on behalf of my client)', { i: true, size: 20 })], { align: AlignmentType.CENTER, after: 240 }),

      rich([t('Ref. No.: '), BLANK('__________'), t('    Date: '), BLANK('__________')], { after: 200 }),

      rich([t('BY REGISTERED POST WITH ACKNOWLEDGEMENT DUE, SPEED POST, COURIER, E-MAIL AND WHATSAPP', { b: true, size: 20 })], { after: 240 }),

      head('TO:'),
      rich([t('1.  Sri Akshat Musaddi', { b: true })], { after: 60 }),
      rich([t('     Sri Arun Musaddi'), t('  (Mob. 9831168888)')], { after: 60 }),
      rich([t('     Proprietors / Partners / Persons in management of:')], { after: 120 }),
      rich([t('     M/s WOOD & WOOD PRODUCTS', { b: true })], { after: 60 }),
      rich([t('     4/3, Foreshore Road, Howrah \u2013 711103, West Bengal')], { after: 60 }),
      rich([t('     GSTIN: 19AADFW7733P1ZB    State Code: 19')], { after: 160 }),
      rich([t('     M/s BG TIMBER', { b: true })], { after: 60 }),
      rich([t('     67/25, Strand Road, Nimtalla, Kolkata \u2013 700006, West Bengal')], { after: 60 }),
      rich([t('     GSTIN: '), BLANK('[IF AVAILABLE]')], { after: 60 }),
      rich([t('     (Hereinafter jointly and severally referred to as \u201cthe Noticee\u201d)', { i: true })], { after: 200 }),

      head('FROM:'),
      rich([t('M/s KOLKATA PLYWOOD AND COMPANY', { b: true }), t('  (a Partnership Firm)')], { after: 60 }),
      rich([t('10/3, Kings Road, Golabari, Howrah \u2013 711101, West Bengal')], { after: 60 }),
      rich([t('GSTIN: 19ABCFK2821J1ZB    E-mail: kolkataplywoodcompany@gmail.com')], { after: 60 }),
      rich([t('Through its Partner, Sri '), BLANK('[FULL NAME \u2014 PRINCE SONI?]'), t('    Mob. 8104234567 / 7300119098')], { after: 60 }),
      rich([t('(Hereinafter referred to as \u201cmy client\u201d or \u201cthe Complainant\u201d)', { i: true })], { after: 200 }),

      rule(),

      rich([
        t('SUBJECT: ', { b: true }),
        t('Notice demanding delivery of conforming goods, refund of excess and undisclosed charges, and compensation for loss and damage caused by inordinate delay and by the supply of defective and non-conforming goods against the order dated 02.06.2026 for 94 pcs lamination and membrane doors, invoice no. 18/WWP/26-27 dated 13.07.2026.', { b: true }),
      ], { after: 200 }),

      p('Sir,'),
      p('Under instructions from and on behalf of my client above named, I do hereby serve upon you this notice as follows:'),

      head('1.  PARTIES AND THE CONTRACT'),
      p('1.1  My client carries on business as a dealer in plywood, doors and allied building materials. The Noticee carries on the business of manufacture and sale of flush doors and allied wood products.'),
      p('1.2  On 02.06.2026 my client met you at your office and placed an order for 94 pcs of flush doors, to be manufactured strictly in accordance with the sizes to be furnished by my client. On the same date my client paid you a sum of Rs. 20,000/- in cash by way of advance at your office.'),
      p('1.3  At the time of placing the said order you expressly represented and assured my client that the goods would be manufactured and delivered within 10 to 15 days. You further represented, in writing on WhatsApp, that your factory has a manufacturing capacity of approximately 3,000 doors per month. My client placed the order and made payment relying upon these representations.'),
      p('1.4  On 04.06.2026 my client furnished to you, in writing on WhatsApp, the complete size list of all 94 doors. That list is the contractual specification and is on record.'),
      p('1.5  On your instructions and at the accounts nominated by you from time to time, my client paid to you the entire consideration aggregating to Rs. 2,21,200/- as follows:'),
      new Table({ columnWidths: [1600, 2100, 3900, 1700], width: { size: 9300, type: WidthType.DXA }, rows: payRows }),
      p('', { after: 120 }),
      p('1.6  It is material to record that of the said aggregate sum of Rs. 2,21,200/-, only Rs. 80,000/- was received into the current account maintained in the name of M/s Wood and Wood Products. The balance was taken by you in cash and into savings bank accounts standing in the names of third parties nominated by you, namely Sri Ramkrishna Mandal and Sri Ajay Sharma, neither of whom is a party to the contract. The particulars of each such account were furnished by you to my client in writing on WhatsApp, and those messages are on record.'),
      p('1.7  Despite receipt of the entire consideration, you have to date failed and neglected to issue to my client a tax invoice for the full value of the goods supplied. My client calls upon you to do so forthwith.'),

      head('2.  CHRONOLOGY OF EVENTS'),
      p('2.1  The material events are set out below. Each of them is supported by contemporaneous WhatsApp messages, call recordings, bank statements and transport documents in my client’s possession:'),
      new Table({ columnWidths: [2100, 7200], width: { size: 9300, type: WidthType.DXA }, rows: chronRows }),
      p('', { after: 160 }),

      head('3.  BREACH OF CONTRACT BY THE NOTICEE'),
      p('3.1  DELAY IN DELIVERY. On 19.07.2025 you represented in writing that your terms were \u201c40% advance / 12-13 days delivery\u201d. The order was placed on 02.06.2026, the specification furnished on 04.06.2026, and 40% of the price stood paid by 09.06.2026. Delivery was nevertheless effected only on 11.07.2026 and the goods reached the site on 12.07.2026, being 37 days from the date of the order. Time was of the essence and you were aware throughout that the goods were required for onward supply to a builder at a live construction site.'),
      p('3.2  MISREPRESENTATION AS TO CAPACITY. When my client asked, before placing the order, whether there would be any difficulty as to quality, you represented in writing that \u201chumlog 3000-3500 monthly banate hai, aisse toh koi dikkat aata nai\u201d. On the strength of that representation my client placed the order. An order of 94 pieces then took 37 days, and the replacement of 21 pieces has taken a further six weeks and is still incomplete.'),
      p('3.3  THE PLEA OF WEATHER IS UNTENABLE. From 17.08.2026 onwards you attributed the delay to the absence of sunshine. On 14.08.2026 you had confirmed, in answer to the express question \u201cPakka Final Na ??\u201d, simply \u201cYes.\u201d My client has since observed the manufacturing process at your factory and the setting of resin and skin on a door takes about twenty minutes, three doors having been completed in his presence within half an hour. The delay is not attributable to the weather.'),
      p('3.4  SUPPLY OF NON-CONFORMING GOODS. Against the size list of 04.06.2026, in which every lamination door was within 32 inches in width, you manufactured 10 pcs at 36 inches and 36.5 inches. You further supplied 11 pcs in MEMBRANE where LAMINATION had been ordered. You supplied lamination doors with a design on one side only although both-side design had been agreed, and the thickness of the doors was not uniform. The sizes of the remaining doors were out by between 2 suta and 1 inch, so that my client\u2019s customer had to engage carpenters to cut and re-fit the consignment.'),
      p('3.5  CHARGES LEVIED WITHOUT DISCLOSURE. The signed order note of 02.06.2026 records only the rates, delivery at Jhargram, 50% billing and the advance. It contains no cutting charge, no loading charge and no differential rate. Your own rate message of 01.06.2026 likewise contains none. Yet at the time of despatch you levied for the first time a \u201ccutting extra for different sizes\u201d of Rs 2 per sq. ft. on 1622.907 sq. ft., amounting to Rs 3,245.81, a loading charge of Rs 1,000, and a rate of Rs 132 per sq. ft. on 15 pcs described as \u201clamination short sizes less than 25 inches\u201d as against Rs 122 for the rest. Having charged separately for cutting, you delivered scarcely a single door in the exact ordered size.'),
      p('3.6  THE PRICE WAS INFLATED AT THE LAST MOMENT. On 08.07.2026 you sent my client a statement putting the sum payable at Rs 2,13,738. On 11.07.2026, on the very day of despatch and after Rs 1,70,000 had already been paid, you sent a fresh statement putting the figure at Rs 2,21,214.50 \u2014 an increase of Rs 7,476.50 within three days. The increase is made up of the cutting charge of Rs 3,245.81 and the loading charge of Rs 1,000 introduced for the first time, the sq. ft. having been re-measured upward, and 15 pcs being re-rated at Rs 132 in place of Rs 122. By that stage the goods were loaded and my client had no practical choice but to pay.'),
      p('3.7  FAILURE AS TO DOCUMENTS. The invoice issued on 11.07.2026 was made out in the name of a stranger, \u201cASF Plywood\u201d, and had to be corrected only after my client pointed it out. Worse, on 22.08.2026 the replacement goods were released from your factory at about 6:40 p.m. without any document at all \u2014 no delivery challan under Rule 55 of the Central Goods and Services Tax Rules, 2017, no invoice and no e-way bill \u2014 although my client had asked for the challan in writing on 21.08.2026 at 16:14 hrs and again on 22.08.2026 at 17:03 hrs. Further, when the 21 defective pieces were returned to your factory on 30.07.2026 you issued no credit note and accepted no debit note, so that goods for which my client has paid in full, and which have been lying in your possession for over three weeks, remain unaccounted for in the records of both parties.'),
      p('3.8  FREIGHT. Delivery at Jhargram was part of the bargain from the outset, as your own memorandum of 02.06.2026 records in the words \u201c+ Delivery Jhargram\u201d, and my client has already paid you Rs 13,000 on that account under the head \u201cTEMPO\u201d in your statement of 11.07.2026. The freight so paid was rendered wholly infructuous as to the 21 defective pieces. The return of those pieces on 29.07.2026 and their re-transport on 22.08.2026 were occasioned solely by your own admitted default, and the cost of both \u2014 Rs 7,000 and Rs 6,940 respectively \u2014 has been borne by my client out of his own pocket because you expressly refused to bear it. A party in breach cannot require the innocent party to pay twice over for carriage made necessary by that breach.'),
      p('3.9  ADMISSION OF DEFAULT. On 17.07.2026 you stated \u201cAur isko return karwa dijiye, change karke bhejwa denge\u201d and on 30.07.2026 \u201cWill rectify and send \u2014 Cannot confirm exact time now\u201d. The defect and your liability to rectify it are therefore admitted. What is in issue is only the delay and the cost, both of which you have declined to answer for.'),
      p('3.10  CONDUCT. When pressed for a date, you replied \u201cThen wait kariye\u201d, \u201cI don\u2019t care. Mereko mat boliye\u201d, \u201cBut koi bakwass nai sunuenga\u201d and \u201cBaad mein hum kuch nai sunega\u201d. On 31.07.2026, when my client had already set out for your factory to inspect the goods, you messaged \u201cNo need to come.\u201d On 09.07.2026 and repeatedly thereafter my client informed you in writing that he wished to be present at the time of loading; you dissuaded him, and the defects that were not caught at loading are the direct consequence.'),
      p('3.11  THE REPLACEMENT GOODS ARE THEMSELVES DEFECTIVE. On inspection at your factory on 22.08.2026 the resin and skin in several of the replacement doors were found not to have adhered properly, leaving the core paper visible. My client reserves all his rights in respect of these defects, which are being separately documented upon receipt.'),

      head('4.  LOSS AND DAMAGE SUFFERED BY MY CLIENT'),
      p('4.1  By reason of the aforesaid breaches my client has suffered the following quantified out-of-pocket loss:'),
      new Table({ columnWidths: [700, 6800, 1800], width: { size: 9300, type: WidthType.DXA }, rows: damageRows }),
      p('', { after: 120 }),
      p('4.2  In addition to the above, my client has suffered a loss of business and of reputation. The goods were required for M/s Kriti Construction, a builder and promoter engaged in construction on a substantial scale, whose further project was nearing completion and against which an order for a further 150 pcs of doors was in prospect, and who had two to three further projects due to commence. By reason of your delay and your supply of non-conforming goods that customer has abused my client, attended his office and, on 11.08.2026, humiliated him before his own staff, and my client\u2019s commercial relationship with him stands gravely prejudiced. My client reserves the right to quantify and claim this loss separately.'),
      rich([t('4.3  My client further reserves the right to claim interest at the rate of '), BLANK('[RATE]'), t('% per annum on all sums claimed herein from the respective dates of payment until realisation.')], { after: 120 }),

      head('5.  DEMAND'),
      p('In the premises aforesaid, I, on behalf of and under instructions from my client, hereby call upon you to, within FIFTEEN (15) DAYS of the receipt of this notice:'),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Deliver to my client, at your own cost and risk including all freight and loading charges, the balance of the 21 pcs of lamination doors manufactured strictly in accordance with the size list furnished on 04.06.2026, and replace any of the doors despatched on 22.08.2026 in which the resin and skin have not properly adhered;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Issue forthwith the delivery challan under Rule 55 of the Central Goods and Services Tax Rules, 2017 and, where required, the e-way bill, in respect of the goods released from your factory on 22.08.2026, no document of any kind having been issued at the time of despatch, and likewise in respect of the balance pieces;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Reimburse to my client the freight of Rs 7,000 paid on 29.07.2026 and Rs 6,940 paid on 22.08.2026, both having been made necessary solely by your own admitted default, my client having already paid you Rs 13,000 towards carriage under your statement of 11.07.2026;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Reimburse to my client Rs 1,000 towards re-measurement and Rs 7,300 towards the cutting and re-fitting of 73 doors at the customer\u2019s site;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Account for the increase of Rs 7,476.50 made between your statement of 08.07.2026 and your statement of 11.07.2026, and refund so much of it as represents the cutting charge, the loading charge and the re-rating of the short sizes;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Refund the cutting charge of Rs 3,245.81, the loading charge of Rs 1,000 and the undisclosed rate differential of Rs 2,287.90, none of which was disclosed at the time of the order, at the time of the size list or at the time of payment, together with a corresponding credit note under the Central Goods and Services Tax Act, 2017; and')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 160, line: 276 },
        children: [t('Issue forthwith a credit note in respect of the 21 pcs returned to you on 30.07.2026, no credit note having been issued and no debit note accepted at the time, so as to enable my client to correctly reflect the transaction in his statutory records.')] }),

      head('6.  CONSEQUENCE OF NON-COMPLIANCE'),
      p('TAKE NOTICE that in the event of your failure or neglect to comply with the aforesaid demands within the said period of fifteen (15) days from the receipt hereof, my client shall be constrained to initiate appropriate civil proceedings against you before the competent court for recovery of the aforesaid sums together with damages, interest and costs, and to pursue such further and other remedies as may be available to him in law, entirely at your risk as to costs and consequences.'),
      p('TAKE FURTHER NOTICE that my client is in possession of the complete WhatsApp correspondence exchanged between the parties, audio recordings of the telephonic conversations between the parties, bank statements evidencing all payments, transport documents, and photographic and video records of the goods, all of which shall be produced and relied upon in any such proceedings.'),
      p('A copy of this notice is retained in my office for record and for further necessary action.'),

      p('', { after: 200 }),
      rich([t('Yours faithfully,')], { after: 320 }),
      rich([BLANK('[NAME OF ADVOCATE]')], { after: 60 }),
      rich([t('Advocate')], { after: 60 }),
      rich([t('Enrolment No.: '), BLANK('__________')], { after: 60 }),
      rich([t('Address: '), BLANK('__________________________')], { after: 60 }),
      rich([t('Mobile: '), BLANK('__________'), t('   E-mail: '), BLANK('__________')], { after: 240 }),

      rule(),

      head('ANNEXURES TO BE ENCLOSED'),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Photograph of the Noticee\u2019s signed order note dated 2/06/26 recording the rates, \u201c+ Delivery Jhargram\u201d, \u201c+ 50% Bill\u201d and the advance.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('WhatsApp message dated 19.07.2025 recording \u201c40% advance / 12-13 days delivery\u201d.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('WhatsApp message dated 25.07.2025 recording \u201chumlog 3000-3500 monthly banate hai\u201d.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('WhatsApp rate message dated 01.06.2026 (Membrane 105 / Lamination 112 / Side Pine +15).')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('WhatsApp messages dated 04.06.2026 enclosing the size list of 25 pcs membrane and 69 pcs lamination.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Bank statements and transaction receipts for all five payments aggregating Rs 2,21,200, together with the WhatsApp messages by which the Noticee nominated each account.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('The Noticee\u2019s statement of account dated 08.07.2026 showing the final figure of Rs 2,13,738.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Tax invoice no. 18/WWP/26-27 dated 11.07.2026 in the name of ASF Plywood, and the revised invoice of the same number dated 13.07.2026 in my client\u2019s name.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('E-way bill no. 8517 1317 0306 dated 13.07.2026.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('WhatsApp messages dated 16.07.2026 recording the defects, and the five pages of comparative measurements sent on 20.07.2026.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('WhatsApp messages dated 17.07.2026 and 30.07.2026 containing the Noticee\u2019s admissions.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Transport receipt dated 29.07.2026 evidencing payment of Rs 7,000 towards return freight.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Receipts for the carpentry charges of Rs 1,000 and Rs 7,000.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Transport operator\u2019s trip record for 22.08.2026 (Trip CRN113083391386, vehicle WB-01-BB-8397) showing detention of 4 hours 06 minutes, 156 minutes of billed waiting time and a fare of Rs 6,940.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Photographs of the non-conforming doors and of the defective resin and skin adhesion in the replacement doors.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Video recordings made at the Noticee\u2019s factory on 22.08.2026 evidencing the refusal to issue any despatch document.')] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Index of the audio recordings relied upon, with a certificate under Section 63 of the Bharatiya Sakshya Adhiniyam, 2023.')] }),
      p('', { after: 200 }),
      rule(),
      head('NOTES FOR THE ADVOCATE (to be deleted before issue)'),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('THE MEMORANDUM OF 02.06.2026 \u2014 how it must be pleaded. The body of the memorandum is in the NOTICEE\u2019S handwriting on a plain unheaded sheet; the signature at its foot is the CLIENT\u2019S, taken by the Noticee as an acknowledgement of the Rs 20,000 cash. The absence of any letterhead makes it the more important to plead the handwriting, the transmission and the eleven weeks of silence, rather than the document itself. The Noticee kept the original and allowed only a photograph. It should therefore be pleaded not as a document executed by the Noticee but (a) as a writing in the Noticee\u2019s own hand recording the terms he himself dictated, and (b) as a statement communicated to him on WhatsApp the same day at 17:19 hrs and never denied by him over the following eleven weeks, and so admitted by his silence and by his subsequent conduct. Please also consider a notice to produce the original, which remains in his custody; its non-production or destruction will support an adverse inference.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('The client has read receipts disabled on WhatsApp, so his outgoing messages do not display a read indicator. Receipt of the material messages is nevertheless established by the Noticee\u2019s own replies to them, and by delivery, which suffices for the purposes of notice under the contract. Please have the device examined and the messages certified accordingly.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Please verify the constitution and correct legal names of M/s Wood & Wood Products and M/s BG Timber from the GST portal before issue, and satisfy yourself that both are commonly controlled before proceeding against both. The tax invoice is issued by Wood & Wood Products (GSTIN 19AADFW7733P1ZB) but the cash advance of 02.06.2026 was taken in the name of BG Timber.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('MATTERS ADVERSE TO THE CLIENT, on which he must be advised before this notice is issued. First, the Noticee\u2019s signed order note of 02.06.2026 itself records \u201c+ 50% Bill\u201d, so the partial-billing arrangement is documented and was known to both sides; the client\u2019s customer was also aware of it. Second, on 11.02.2026 the client himself asked the Noticee \u201cAgr koi Kachhe me Maal lene khojega to kese bhejenge ??\u201d. Third, on 21.08.2026 at 16:26 hrs the client wrote \u201cBill nahi dene se bhi chalega\u2026 i will manage\u2026\u201d, which weakens the complaint at paragraph 3.7 so far as the despatch of 22.08.2026 is concerned; the demand for the challan was however repeated on 21.08.2026 and again on 22.08.2026 and is preserved. Please consider whether paragraph 3.7 should be confined to the ASF Plywood invoice and to the challan.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('Of the consideration of Rs 2,21,200, only Rs 80,000 reached the Noticee firm\u2019s own current account. Rs 20,000 was paid in cash and Rs 1,21,200 into savings accounts of three individuals nominated by the Noticee (Sri Ajay Sharma, Sri Ramkrishna Mondal and, at one stage, Sri Anuj Kumar Pandey). Proof of payment rests entirely on the WhatsApp messages by which the Noticee furnished those account particulars; those messages must be preserved and certified. Please also advise the client separately on Section 40A(3) of the Income-tax Act, 1961 in respect of the cash payment, and on his own books and input tax credit, BEFORE this notice is issued.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('The rates recorded in the signed note of 02.06.2026 (Lamination 140, Membrane Rosewood 135) do not match the rates actually billed (Lamination 122, short sizes 132, Membrane 105). Please take the client\u2019s instructions on this before settling paragraph 3.5, as the discrepancy may be explained by the design codes BGL 10 and BG 26.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('The aggregate claim is presently below the specified value of Rs 3,00,000 under the Commercial Courts Act, 2015 unless loss of business is quantified and included. Please advise the client on the appropriate forum and, if the Commercial Courts Act is attracted, on pre-institution mediation under Section 12A.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('The client is a trader who purchased the goods for onward resale; the availability of a remedy under the Consumer Protection Act, 2019 in view of Section 2(7) may require consideration.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('The client holds audio recordings of conversations to which he was himself a party, and video recordings made at the Noticee\u2019s factory on 22.08.2026. Please preserve the originals on the original device and prepare the certificate under Section 63 of the Bharatiya Sakshya Adhiniyam, 2023.', { i: true })] }),
      new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t('The demands relating to credit notes and challans are included solely to protect the client\u2019s own statutory records. They are not to be framed, and must not be understood, as any threat of a complaint to the tax authorities.', { i: true })] }),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/nationalplywood/legal/Legal-Notice-Draft-Wood-and-Wood-Products.docx', b);
  console.log('written');
});
