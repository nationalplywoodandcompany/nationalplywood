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
  ['02.06.2026', 'Meeting at the Noticee’s office. Advance of Rs. 20,000/- paid in cash. Noticee represented that delivery would be completed within 10 to 15 days.'],
  ['04.06.2026', 'Complete door size list (94 pcs) sent to the Noticee on WhatsApp.'],
  ['09.06.2026', 'Rs. 80,000/- transferred to the Noticee’s current account on his demand.'],
  ['[DATE]', 'Balance of Rs. [AMOUNT]/- transferred to the savings account nominated by the Noticee.'],
  ['24.06.2026', 'Outer limit of the promised 15-day delivery period expired. No delivery made.'],
  ['25.06.2026 to 10.07.2026', 'Delivery repeatedly postponed by the Noticee on varying pretexts (rain, delayed payment, vehicle problems). Assurances of 30.06, 02.07, 09.07, 10.07 and 11.07 given and broken in turn.'],
  ['09.07.2026 to 11.07.2026', 'The Complainant travelled to and stayed at Kolkata for three days specifically to be present at the time of loading, having informed the Noticee in advance in writing. The Noticee assured him that his attendance was unnecessary as the goods were correct.'],
  ['11.07.2026', 'Goods dispatched. No invoice, no challan and no transporter contact was shared at the time of dispatch, despite repeated requests. The transporter’s number was furnished only past midnight after repeated calls.'],
  ['12.07.2026', 'Goods received at the client’s site at Jhargram. On inspection, 21 pcs found wholly wrong — 10 pcs manufactured at 36" against an order of 32", and 11 pcs manufactured in a design never ordered. The remaining pieces were not manufactured to the exact sizes furnished on 04.06.2026.'],
  ['13.07.2026 to 29.07.2026', 'The Complainant engaged his own carpenter for two days to re-measure and segregate the entire consignment (Rs. 1,000/-), and the client’s carpenters had to cut and re-fit 73 doors at Rs. 100/- per door (Rs. 7,000/-).'],
  ['30.07.2026', 'The 21 defective doors returned to the Noticee’s factory at the Complainant’s own cost of Rs. 7,000/- towards freight, the Noticee having expressly refused to bear the same.'],
  ['17.08.2026', 'The Noticee stated that manufacture requires only one day and that the goods would be ready on the first clear day.'],
  ['18.08.2026 to 21.08.2026', 'Ready-date shifted daily. On 21.08.2026 the Noticee asked that a vehicle be sent that evening, stating the goods were ready and being packed.'],
  ['22.08.2026', 'Loading time shifted from 10:30 a.m. to 2:00 p.m. The Complainant attended the factory with a vehicle at the appointed time and found that only 18 of the 21 doors had been made, and those still under manufacture. He was made to wait several hours.'],
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
  ['1.', 'Freight paid for returning 21 defective doors to the Noticee’s factory on 30.07.2026', '7,000'],
  ['2.', 'Freight now payable for re-transporting the replaced doors to the client’s site', '[AMOUNT]'],
  ['3.', 'Charges paid to the Complainant’s own carpenter for two days’ re-measurement and segregation of the consignment', '1,000'],
  ['4.', 'Cutting and re-fitting charges for 73 doors at Rs. 100/- per door, paid to the client’s carpenters', '7,000'],
  ['5.', 'Cutting charge of Rs. 2/- per sq. ft. levied in the invoice without prior disclosure', '[AMOUNT]'],
  ['6.', 'Loading charge levied in the invoice without prior disclosure', '1,000'],
  ['7.', 'Travel, lodging and incidental expenses incurred at Kolkata on 09.07.2026 to 11.07.2026 and on 22.08.2026', '[AMOUNT]'],
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
  dmgRow(['', 'Total quantified out-of-pocket loss (excluding items to be filled in)', '15,000'], { b: true }),
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
      rich([t('1.  Sri Akshat ', { b: true }), BLANK('[VERIFY FULL NAME AND SPELLING]')], { after: 60 }),
      rich([t('     Proprietor / Partner / Authorised Signatory')], { after: 60 }),
      rich([t('     M/s Wood and Wood Products', { b: true })], { after: 60 }),
      rich([t('     Factory and Registered Office: '), BLANK('[FULL POSTAL ADDRESS]')], { after: 60 }),
      rich([t('     GSTIN: '), BLANK('[GSTIN FROM INVOICE]')], { after: 160 }),

      rich([t('2.  M/s BG Timber', { b: true }), t(' (sister concern / associated firm of the above)')], { after: 60 }),
      rich([t('     Office: Neem Tala, Kolkata, West Bengal — '), BLANK('[FULL ADDRESS AND PIN]')], { after: 60 }),
      rich([t('     GSTIN: '), BLANK('[IF AVAILABLE]')], { after: 60 }),
      rich([t('     (Hereinafter jointly and severally referred to as “the Noticee”)', { i: true })], { after: 200 }),

      head('FROM:'),
      rich([t('Sri '), BLANK('[FULL NAME OF COMPLAINANT]')], { after: 60 }),
      rich([t('Proprietor / Partner, '), BLANK('[FULL NAME OF FIRM]')], { after: 60 }),
      rich([t('Address: '), BLANK('[FULL POSTAL ADDRESS]')], { after: 60 }),
      rich([t('GSTIN: '), BLANK('[YOUR GSTIN]'), t('    Mobile: '), BLANK('[NUMBER]')], { after: 60 }),
      rich([t('(Hereinafter referred to as “my client” or “the Complainant”)', { i: true })], { after: 200 }),

      rule(),

      rich([
        t('SUBJECT: ', { b: true }),
        t('Notice demanding delivery of conforming goods, refund of undisclosed charges and compensation for loss and damage caused by inordinate delay and supply of defective and non-conforming goods against order dated 04.06.2026 for 94 pcs flush doors.', { b: true }),
      ], { after: 200 }),

      p('Sir,'),
      p('Under instructions from and on behalf of my client above named, I do hereby serve upon you this notice as follows:'),

      head('1.  PARTIES AND THE CONTRACT'),
      p('1.1  My client carries on business as a dealer in plywood, doors and allied building materials. The Noticee carries on the business of manufacture and sale of flush doors and allied wood products.'),
      p('1.2  On 02.06.2026 my client met you at your office and placed an order for 94 pcs of flush doors, to be manufactured strictly in accordance with the sizes to be furnished by my client. On the same date my client paid you a sum of Rs. 20,000/- in cash by way of advance at your office.'),
      p('1.3  At the time of placing the said order you expressly represented and assured my client that the goods would be manufactured and delivered within 10 to 15 days. You further represented, in writing on WhatsApp, that your factory has a manufacturing capacity of approximately 3,000 doors per month. My client placed the order and made payment relying upon these representations.'),
      p('1.4  On 04.06.2026 my client furnished to you, in writing on WhatsApp, the complete size list of all 94 doors. That list is the contractual specification and is on record.'),
      rich([t('1.5  On your demand my client transferred a further sum of Rs. 80,000/- to your current account on 09.06.2026, and a further sum of Rs. '), BLANK('[AMOUNT]'), t('/- on '), BLANK('[DATE]'), t(' to a savings account nominated by you. The entire consideration stood duly paid by my client.')], { after: 120 }),

      head('2.  CHRONOLOGY OF EVENTS'),
      p('2.1  The material events are set out below. Each of them is supported by contemporaneous WhatsApp messages, call recordings, bank statements and transport documents in my client’s possession:'),
      new Table({ columnWidths: [2100, 7200], width: { size: 9300, type: WidthType.DXA }, rows: chronRows }),
      p('', { after: 160 }),

      head('3.  BREACH OF CONTRACT BY THE NOTICEE'),
      p('3.1  DELAY IN DELIVERY. Against a firm assurance of 10 to 15 days, delivery was effected only on 11.07.2026, being approximately 38 days from the date of the order dated 04.06.2026 and approximately 32 days from the date of the payment of 09.06.2026. Time was of the essence, as you were fully aware that the goods were required for onward supply to my client’s customer at a construction site.'),
      p('3.2  THE PLEA OF RAIN IS FALSE. You have throughout attributed the delay to the monsoon. That plea is untenable. My client has personally observed the manufacturing process at your factory on 22.08.2026: the resin application, skin pressing and levelling of a flush door takes approximately twenty minutes to set, and within half an hour three doors were completed in his presence. On that basis the entire order of 94 pcs was capable of being manufactured within a matter of days. In any event, there was at no point a continuous spell of rain of such duration as could account for a delay of this magnitude.'),
      p('3.3  SUPPLY OF NON-CONFORMING GOODS. Of the 94 pcs supplied, 21 pcs were wholly non-conforming: 10 pcs were manufactured at 36 inches against an ordered size of 32 inches, and 11 pcs were manufactured in a design that was never ordered by my client at all. The remaining pieces were not manufactured to the exact sizes furnished in the size list dated 04.06.2026, and my client’s customer was compelled to engage carpenters to cut and re-fit 73 doors at site.'),
      p('3.4  UNDISCLOSED CHARGES. At the time of final dispatch you levied for the first time a cutting charge of Rs. 2/- per sq. ft. and a loading charge of Rs. 1,000/-. Neither of these charges was disclosed to my client on 02.06.2026 when the order was placed, or on 04.06.2026 when the size list was furnished, or on 09.06.2026 when payment was made. Such unilateral imposition of charges after the contract was concluded and paid for is wholly impermissible. It is further material that, having charged separately for cutting, you failed to deliver even a single door in the exact ordered size.'),
      p('3.5  REFUSAL TO BEAR THE COST OF YOUR OWN DEFECT. Upon the defect being pointed out, you admitted in terms (and the said admission stands recorded) that the goods were wrongly manufactured by you and that you would rectify the same. You nevertheless refused to bear the freight for the return of the defective goods, requiring my client to bear Rs. 7,000/- out of his own pocket, and have since refused to commit to bearing the freight for their re-delivery.'),
      p('3.6  FAILURE OF DUTIES AT DISPATCH. At the time of dispatch on 11.07.2026 you failed to furnish the invoice, the challan or the transporter’s contact details, despite my client requesting the same in writing from the early evening onwards. As a consequence my client was unable to make arrangements for unloading labour at the destination and was compelled to telephone you past midnight, whereupon you responded with objection rather than with the information sought.'),
      p('3.7  CONTINUING BREACH IN RESPECT OF THE REPLACEMENT. The 21 defective pieces were returned to your factory on 30.07.2026. On 17.08.2026 you stated that manufacture required only one day. Thereafter, from 18.08.2026 onwards, you shifted the ready-date daily. On 21.08.2026 you directed my client to send a vehicle that very evening on the representation that the goods were ready and were being packed; had my client done so, his vehicle would have stood idle overnight at his cost. On 22.08.2026 you first fixed loading for 10:30 a.m., then deferred it to 2:00 p.m., and when my client attended with a vehicle at the appointed hour only 18 of the 21 pieces had been made and those were still under manufacture, my client being made to wait for several hours at your premises.'),

      head('4.  LOSS AND DAMAGE SUFFERED BY MY CLIENT'),
      p('4.1  By reason of the aforesaid breaches my client has suffered the following quantified out-of-pocket loss:'),
      new Table({ columnWidths: [700, 6800, 1800], width: { size: 9300, type: WidthType.DXA }, rows: damageRows }),
      p('', { after: 120 }),
      p('4.2  In addition to the above, my client has suffered a loss of business and of reputation. The goods were required for a customer of my client who is a builder and promoter engaged in construction on a substantial scale and who was a prospective source of continuing business worth several lakhs of rupees to my client. By reason of your delay of approximately 38 days and your supply of non-conforming goods, my client has been put to serious embarrassment before the said customer and his commercial relationship with him stands gravely prejudiced. My client reserves the right to quantify and claim this loss separately.'),
      rich([t('4.3  My client further reserves the right to claim interest at the rate of '), BLANK('[RATE]'), t('% per annum on all sums claimed herein from the respective dates of payment until realisation.')], { after: 120 }),

      head('5.  DEMAND'),
      p('In the premises aforesaid, I, on behalf of and under instructions from my client, hereby call upon you to, within FIFTEEN (15) DAYS of the receipt of this notice:'),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Deliver to my client, at your own cost and risk including all freight and loading charges, the 21 pcs of flush doors manufactured strictly in accordance with the size list furnished on 04.06.2026, together with a proper tax invoice, delivery challan and e-way bill;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Reimburse to my client the sum of Rs. 7,000/- paid by him towards freight for the return of the defective goods on 30.07.2026;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Reimburse to my client the sums of Rs. 1,000/- and Rs. 7,000/- paid by him towards re-measurement and towards the cutting and re-fitting of 73 doors respectively;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Refund to my client the cutting charge of Rs. 2/- per sq. ft. and the loading charge of Rs. 1,000/- levied without prior disclosure, together with a corresponding credit note under the Central Goods and Services Tax Act, 2017;')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 100, line: 276 },
        children: [t('Issue to my client a proper credit note and/or delivery challan in respect of the 21 pcs returned to you on 30.07.2026, so as to enable my client to correctly reflect the transaction in his statutory returns; and')] }),
      new Paragraph({ numbering: { reference: 'demands', level: 0 }, spacing: { after: 160, line: 276 },
        children: [t('Confirm in writing that the full value of the goods supplied has been duly reported by you in your outward supply returns under the Central Goods and Services Tax Act, 2017, so that my client may avail of the input tax credit to which he is lawfully entitled.')] }),

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
      ...[
        'Copy of the WhatsApp message dated 04.06.2026 enclosing the size list of 94 doors.',
        'Copies of the WhatsApp messages recording the assurance of delivery within 10 to 15 days and the representation as to manufacturing capacity.',
        'Bank statement / transaction receipts evidencing the payments dated 09.06.2026 and thereafter.',
        'Copy of the tax invoice dated 11.07.2026 showing the cutting and loading charges.',
        'Copy of the e-way bill(s) generated in respect of the consignment dispatched on 11.07.2026.',
        'Transport receipt / bilty dated 30.07.2026 evidencing payment of Rs. 7,000/- towards return freight.',
        'Receipts or acknowledgements in respect of the carpentry charges of Rs. 1,000/- and Rs. 7,000/-.',
        'Photographs and measurement records of the 21 non-conforming doors.',
        'Index of the audio recordings relied upon, together with a certificate under Section 63 of the Bharatiya Sakshya Adhiniyam, 2023.',
      ].map(s => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t(s)] })),

      p('', { after: 200 }),
      rule(),
      head('NOTES FOR THE ADVOCATE (to be deleted before issue)'),
      ...[
        'Please verify the correct legal name, constitution and address of both firms from the tax invoice and from the GST portal before issue. The notice is addressed to both concerns on the footing that they are commonly controlled; please confirm this before proceeding against both.',
        'All highlighted fields require figures or dates from the client’s records.',
        'If the aggregate claim exceeds the specified value of Rs. 3,00,000/-, the suit will lie before the Commercial Court and pre-institution mediation under Section 12A of the Commercial Courts Act, 2015 will be mandatory save where urgent interim relief is sought. Please advise the client accordingly.',
        'The client is a trader who purchased the goods for onward resale; the availability of a remedy under the Consumer Protection Act, 2019 in view of Section 2(7) may require consideration.',
        'The client holds audio recordings of the conversations to which he was himself a party. Please preserve the original recordings on the original device and prepare the certificate under Section 63 of the Bharatiya Sakshya Adhiniyam, 2023.',
        'The demands relating to GST reporting and credit notes are included solely to protect the client’s own input tax credit position and his statutory records. They are not to be framed, and must not be understood, as any threat of a complaint to the tax authorities.',
      ].map(s => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { after: 80, line: 276 }, children: [t(s, { i: true })] })),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('/home/user/nationalplywood/legal/Legal-Notice-Draft-Wood-and-Wood-Products.docx', b);
  console.log('written');
});
