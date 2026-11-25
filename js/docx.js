/* Word (.docx) export. Builds an editable Word document with the same numbering as the
   official form (Part-1 to Part-5), entirely in the browser: WordprocessingML XML packed
   into a ZIP (stored, no compression). No external library, nothing leaves the device. */
(function () {
  const S = window.Schema, F = window.FormPrint.F;

  // ---------- tiny ZIP writer (store only) ----------
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  const crc32 = (u8) => { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const enc = new TextEncoder();

  function zip(files) { // files: [{ name, data: string | Uint8Array }]
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const parts = [], central = [];
    let offset = 0;
    for (const f of files) {
      const name = enc.encode(f.name);
      const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      const crc = crc32(data);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
      local.setUint16(8, 0, true); local.setUint16(10, dosTime, true); local.setUint16(12, dosDate, true);
      local.setUint32(14, crc, true); local.setUint32(18, data.length, true); local.setUint32(22, data.length, true);
      local.setUint16(26, name.length, true); local.setUint16(28, 0, true);
      parts.push(new Uint8Array(local.buffer), name, data);
      const cen = new DataView(new ArrayBuffer(46));
      cen.setUint32(0, 0x02014b50, true); cen.setUint16(4, 20, true); cen.setUint16(6, 20, true); cen.setUint16(8, 0x0800, true);
      cen.setUint16(10, 0, true); cen.setUint16(12, dosTime, true); cen.setUint16(14, dosDate, true);
      cen.setUint32(16, crc, true); cen.setUint32(20, data.length, true); cen.setUint32(24, data.length, true);
      cen.setUint16(28, name.length, true); cen.setUint32(42, offset, true);
      central.push(new Uint8Array(cen.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cenSize = central.reduce((n, p) => n + p.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cenSize, true); end.setUint32(16, offset, true);
    return new Blob([...parts, ...central, new Uint8Array(end.buffer)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  }

  // ---------- WordprocessingML helpers ----------
  const x = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const TEXT_W = 9906; // A4 width 11906 twips − 2 × 1000 margins

  function run(text, o = {}) {
    const rPr = `<w:rPr>${o.b ? '<w:b/><w:bCs/>' : ''}${o.color ? `<w:color w:val="${o.color}"/>` : ''}${o.sz ? `<w:sz w:val="${o.sz}"/><w:szCs w:val="${o.sz}"/>` : ''}</w:rPr>`;
    return String(text ?? '').split('\n').map((line, i) => `<w:r>${rPr}${i ? '<w:br/>' : ''}<w:t xml:space="preserve">${x(line)}</w:t></w:r>`).join('');
  }
  function para(content, o = {}) {
    const pPr = `<w:pPr>${o.keepNext ? '<w:keepNext/>' : ''}<w:spacing w:before="${o.before || 0}" w:after="${o.after ?? 0}"/>${o.align ? `<w:jc w:val="${o.align}"/>` : ''}</w:pPr>`;
    return `<w:p>${pPr}${content}</w:p>`;
  }
  const P = (text, o = {}) => para(run(text, o), o);

  function cell(content, w, o = {}) {
    const tcPr = `<w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${o.span ? `<w:gridSpan w:val="${o.span}"/>` : ''}${o.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${o.fill}"/>` : ''}${o.noBorder ? '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>' : ''}<w:vAlign w:val="${o.vAlign || 'top'}"/></w:tcPr>`;
    return `<w:tc>${tcPr}${content || P('')}</w:tc>`;
  }
  function table(widths, rows, o = {}) {
    const border = o.noBorder ? 'nil' : 'single';
    const b = (side) => `<w:${side} w:val="${border}" w:sz="4" w:space="0" w:color="7A8A82"/>`;
    return `<w:tbl><w:tblPr><w:tblW w:w="${widths.reduce((a, c) => a + c, 0)}" w:type="dxa"/><w:tblLayout w:type="fixed"/>
      <w:tblBorders>${b('top')}${b('left')}${b('bottom')}${b('right')}${b('insideH')}${b('insideV')}</w:tblBorders>
      <w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr>
      <w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>
      ${rows.map((r) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${r}</w:tr>`).join('')}</w:tbl>`;
  }

  // numbered question rows: [ক্রম, প্রশ্ন, উত্তর]
  const QW = [760, 3500, TEXT_W - 760 - 3500];
  const qRows = (rows) => table(QW, rows.map(([num, label, value]) =>
    cell(P(num, { b: true, align: 'center' }), QW[0], { fill: 'F2F5F3' }) + cell(P(label, { b: true }), QW[1], { fill: 'F2F5F3' }) + cell(P(value), QW[2])));
  const heading = (text) => P(text, { b: true, sz: 26, color: '0F3D2E', before: 280, after: 120, keepNext: true });

  // ---------- content ----------
  const v = (d, id) => String(S.getVal(d, id) || '').trim();
  const bn = (s) => S.toBn(s);
  const amt = (s) => { const n = S.parseNum(s); return n == null ? String(s || '') : S.money(n); };
  const triple = (d, id) => {
    const parts = [['m', 'মাসিক'], ['q', 'ত্রৈমাসিক'], ['y', 'বার্ষিক']].filter(([k]) => v(d, id + '_' + k)).map(([k, t]) => `${t}: ${amt(v(d, id + '_' + k))}`);
    return parts.join('\n');
  };
  const married = (d) => d.marital && d.marital !== 'অবিবাহিত';

  function addressTable(d, p) {
    const heads = ['বাড়ি/মহল্লার নং ও নাম', 'রোড নং/ এলাকা/ গ্রামের নাম', 'ডাকঘর ও পোস্ট কোড', 'থানা/উপজেলা', 'জেলা', 'বিভাগ'];
    const keys = ['house', 'road', 'post', 'thana', 'district', 'division'];
    const w = [1750, 1750, 1650, 1600, 1550, TEXT_W - 1750 - 1750 - 1650 - 1600 - 1550];
    return table(w, [
      heads.map((h, i) => cell(P(h, { b: true, sz: 19, align: 'center' }), w[i], { fill: 'F2F5F3' })).join(''),
      keys.map((k, i) => cell(P(k === 'post' ? bn(v(d, p + '_' + k)) : v(d, p + '_' + k), { align: 'center' }), w[i])).join(''),
    ]);
  }

  function expenseTable(d) {
    const w = [700, 3000, 2406, 1900, 1900];
    const head = ['ক্রম', 'ব্যয়ের খাত', 'বিবরণ', 'মোট (টাকা)', 'অর্থের উৎস'];
    const rows = [head.map((h, i) => cell(P(h, { b: true, align: 'center' }), w[i], { fill: 'F2F5F3' })).join('')];
    S.EXPENSE_HEADS.forEach((h, i) => {
      const n = i + 1;
      rows.push([
        cell(P(bn(String(n).padStart(2, '0')), { align: 'center' }), w[0]), cell(P(h), w[1]), cell(P(v(d, `e${n}_desc`)), w[2]),
        cell(P(amt(v(d, `e${n}_amt`)), { align: 'center' }), w[3]), cell(P(v(d, `e${n}_src`), { align: 'center' }), w[4]),
      ].join(''));
    });
    const t = S.expensesTotal(d);
    rows.push(cell(P('সর্বমোট', { b: true, align: 'right' }), w[0] + w[1] + w[2], { span: 3, fill: 'F2F5F3' }) +
      cell(P(t == null ? '' : S.money(t), { b: true, align: 'center' }), w[3]) + cell(P(''), w[4]));
    return table(w, rows);
  }

  function body(d, photoRel) {
    const out = [];
    // title + photo box
    const photo = photoRel
      ? para(`<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1005840" cy="1173480"/><wp:docPr id="1" name="Photo"/>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="photo"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="${photoRel}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1005840" cy="1173480"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
          </a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`, { align: 'center' })
      : P('পাসপোর্ট সাইজ ছবি', { align: 'center', sz: 18, color: '7A8A82' });
    const tw = [TEXT_W - 1900, 1900];
    out.push(table(tw, [
      cell(P('উদ্যোগ (UDYOG)', { b: true, sz: 36, color: '0F3D2E' }) + P('তরুণ উদ্যোক্তা অর্থায়ন কর্মসূচি — আবেদন ফরম', { b: true, sz: 24, before: 60 }) +
        P(`আবেদনকারী: ${v(d, 'name_bn') || '—'}`, { before: 120, color: '5D6B64' }), tw[0], { noBorder: true, vAlign: 'center' }) +
      cell(photo, tw[1], { vAlign: 'center' }),
    ], { noBorder: false }));

    out.push(heading('পার্ট-১: আবেদনকারীর তথ্য'));
    out.push(qRows([
      ['০১', 'আবেদনকারীর নাম (বাংলায়)', v(d, 'name_bn')],
      ['০১', 'আবেদনকারীর নাম (In English, Block Letters)', v(d, 'name_en').toUpperCase()],
      ['০২', 'পিতার নাম', v(d, 'father')],
      ['০৩', 'মাতার নাম', v(d, 'mother')],
      ['০৪', 'জাতীয় পরিচয়পত্র নম্বর', bn(v(d, 'nid'))],
      ['০৫', 'জন্ম তারিখ ও আবেদনের শেষ তারিখে বয়স', F.dob(d)],
      ['০৬', 'বৈবাহিক অবস্থা', v(d, 'marital')],
      ['০৭', 'স্বামী/স্ত্রীর নাম', married(d) ? v(d, 'spouse_name') : ''],
      ['০৭', 'স্বামী/স্ত্রীর জাতীয় পরিচয়পত্র নম্বর', married(d) ? bn(v(d, 'spouse_nid')) : ''],
      ['০৮', 'শিক্ষাগত যোগ্যতা (সর্বশেষ পাসকৃত পরীক্ষা, বোর্ড/বিশ্ববিদ্যালয় ও সাল)', F.edu(d)],
      ['০৯', 'ফোন নম্বর / বিকল্প ফোন নম্বর', F.phone(d)],
      ['১০', 'ই-মেইল আইডি', v(d, 'email')],
      ['১১', 'দক্ষতা উন্নয়নমূলক কোনরূপ প্রশিক্ষণ থাকলে তার বিবরণ', v(d, 'training')],
      ['১২', 'ব্যবসা/চাকরির অভিজ্ঞতা থাকলে তার বিবরণ', v(d, 'experience')],
      ['১৩', 'ব্যাংক হিসাব (থাকলে)', F.bank(d)],
      ['১৪', 'কোনো ব্যাংক/আর্থিক প্রতিষ্ঠানে বিদ্যমান ঋণ আছে কিনা; থাকলে বিবরণ', F.yesDetail(d, 'loan_has', 'loan_details')],
    ]));
    out.push(P('১৫. বর্তমান ঠিকানা (জাতীয় পরিচয়পত্র মোতাবেক)', { b: true, before: 200, after: 80, keepNext: true }));
    out.push(addressTable(d, 'pre'));
    out.push(P('১৬. স্থায়ী ঠিকানা', { b: true, before: 200, after: 80, keepNext: true }));
    out.push(addressTable(d, 'per'));

    if (d.biz_mode !== 'existing') {
      const g = (fn) => (d.biz_mode === 'new' ? fn() : '');
      out.push(heading('পার্ট-২: ব্যবসায়িক পরিকল্পনা সংক্রান্ত তথ্যাদি (নতুন উদ্যোগের ক্ষেত্রে)'));
      out.push(qRows([
        ['০১', 'ব্যবসায়িক ধারণা/উদ্যোগের সংক্ষিপ্ত বিবরণ (২০০ শব্দের মধ্যে)', g(() => v(d, 'p2_idea'))],
        ['০২', 'সম্ভাব্য ক্রেতা', g(() => v(d, 'p2_buyers'))],
        ['০৩', 'কত সময়ের মধ্যে ব্যবসায় শুরু করার পরিকল্পনা রয়েছে', g(() => v(d, 'p2_start'))],
        ['০৪', 'ব্যবসার ধরন (উৎপাদন/সেবা/ব্যবসা)', g(() => F.type(d, 'p2_type', 'p2_type_desc'))],
        ['০৫', 'ব্যবসার সম্ভাব্য অবস্থান', g(() => v(d, 'p2_location'))],
        ['০৬', 'পণ্যের ধরন (আমদানির বিকল্প বা রপ্তানিযোগ্য পণ্য কিনা)', g(() => v(d, 'p2_product'))],
        ['০৭', 'কাঁচামালের উৎস (প্রযোজ্য ক্ষেত্রে)', g(() => v(d, 'p2_raw'))],
        ['০৮', 'পণ্যের বাজার (স্থানীয়/আন্তর্জাতিক)', g(() => v(d, 'p2_market'))],
        ['০৯', 'সম্ভাব্য বিক্রয়ের পরিমাণ (টাকায়)', g(() => triple(d, 'p2_sales'))],
        ['১০', 'সম্ভাব্য লাভ/মুনাফা (টাকায়)', g(() => triple(d, 'p2_profit'))],
        ['১১', 'ব্যবসা পরিচালনার ক্ষেত্রে সম্ভাব্য সমস্যাসমূহ, ঝুঁকিসমূহ এবং ঝুঁকিসমূহ মোকাবেলায় পরিকল্পনা', g(() => v(d, 'p2_risk'))],
        ['১২', 'প্রথম বছরে সম্ভাব্য কর্মসংস্থান', g(() => F.people(d, 'p2_emp_total', 'p2_emp_women'))],
        ['১৩', 'যে পণ্য/সেবা বিক্রয় করা হবে, বর্তমানে বাজারে তার অবস্থান কেমন (কমপক্ষে ৩টি প্রতিযোগী প্রতিষ্ঠানের তথ্যসহ)', g(() => v(d, 'p2_compete'))],
      ]));
    }
    if (d.biz_mode !== 'new') {
      const g = (fn) => (d.biz_mode === 'existing' ? fn() : '');
      out.push(heading('পার্ট-৩: ব্যবসা সম্পর্কিত তথ্যাবলি (চলমান ব্যবসার ক্ষেত্রে)'));
      out.push(qRows([
        ['০১', 'ব্যবসা প্রতিষ্ঠানের নাম', g(() => v(d, 'p3_name'))],
        ['০২', 'ব্যবসার ধরন (উৎপাদন/সেবা/ব্যবসা)', g(() => F.type(d, 'p3_type', 'p3_type_desc'))],
        ['০৩', 'ব্যবসা শুরুর তারিখ', g(() => S.fmtDate(d.p3_start))],
        ['০৪', 'ট্রেড লাইসেন্স আছে কিনা; থাকলে ইস্যুর তারিখ ও নম্বর', g(() => F.license(d))],
        ['০৫', 'পণ্য/সেবার গড় বিক্রয়মূল্য (একক প্রতি)', g(() => v(d, 'p3_price'))],
        ['০৬', 'বিক্রয়ের পরিমাণ (টাকায়)', g(() => triple(d, 'p3_sales'))],
        ['০৭', 'লাভ/মুনাফার পরিমাণ (টাকায়)', g(() => triple(d, 'p3_profit'))],
        ['০৮', 'কাঁচামাল ও ব্যবসা পরিচালনা বাবদ বিভিন্ন খাতে মাসিক সম্ভাব্য খরচ (টাকা)', g(() => v(d, 'p3_cost'))],
        ['০৯', 'উদ্যোক্তা কীভাবে আলোচ্য ব্যবসায়ের সঙ্গে জড়িত হয়েছেন', g(() => v(d, 'p3_how'))],
        ['১০', 'ব্যবসার সঙ্গে সম্পর্কিত মোট বিনিয়োগের পরিমাণ (টাকা)', g(() => amt(v(d, 'p3_invest')))],
        ['১১', 'কাঁচামালের উৎস (প্রযোজ্য ক্ষেত্রে)', g(() => v(d, 'p3_raw'))],
        ['১২', 'পণ্য/সেবা বিপণন এলাকা', g(() => v(d, 'p3_area'))],
        ['১৩', 'যে পণ্য/সেবা বিক্রয় করা হয়, বর্তমানে বাজারে তার অবস্থান (কমপক্ষে ৩টি প্রতিযোগী প্রতিষ্ঠানের তথ্যসহ)', g(() => v(d, 'p3_compete'))],
        ['১৪', 'বিদ্যমান ব্যবসায় পরিচালনায় সমস্যাসমূহ, ঝুঁকিসমূহ এবং ঝুঁকিসমূহ মোকাবেলায় পরিকল্পনা', g(() => v(d, 'p3_risk'))],
        ['১৫', 'কর্মীর সংখ্যা', g(() => F.people(d, 'p3_staff_total', 'p3_staff_women'))],
        ['১৬', 'চলমান ব্যবসায় সম্প্রসারণসহ কর্মী বাড়ানোর কোনো পরিকল্পনা রয়েছে কিনা; থাকলে বিবরণ', g(() => v(d, 'p3_expand'))],
      ]));
    }

    out.push(heading('পার্ট-৪: অর্থায়ন পরিকল্পনা'));
    out.push(P('ক. অর্থায়ন প্যাকেজ', { b: true, after: 80, keepNext: true }));
    out.push(qRows([
      ['০১', 'প্রয়োজনীয় মোট মূলধনের পরিমাণ (টাকা)', amt(v(d, 'f_total'))],
      ['০২', 'নিজস্ব বিনিয়োগ (যদি থাকে)', amt(v(d, 'f_own'))],
      ['০৩', 'আবেদনকৃত অর্থায়নের পরিমাণ (টাকা)', amt(v(d, 'f_req'))],
    ]));
    out.push(P('খ. খাতভিত্তিক ব্যয় বিভাজন', { b: true, before: 200, after: 80, keepNext: true }));
    out.push(expenseTable(d));

    out.push(heading('পার্ট-৫: অন্যান্য তথ্য'));
    out.push(qRows([
      ['০১', 'শারীরিক প্রতিবন্ধকতা আছে কিনা', F.yesDetail(d, 'disab', 'disab_type')],
      ['০২', 'ক্ষুদ্র নৃগোষ্ঠীভুক্ত কিনা', v(d, 'ethnic')],
      ['০৩', 'অন্যান্য', v(d, 'other')],
    ]));

    out.push(heading('ঘোষণা'));
    [
      'আমি এই মর্মে ঘোষণা করছি যে, উপরে প্রদত্ত সমস্ত তথ্য সত্য ও সঠিক। কোনো তথ্য অসত্য প্রমাণিত হলে আমার আবেদন/অনুমোদন যেকোনো পর্যায়ে বাতিলযোগ্য হবে।',
      'আমি আরও ঘোষণা করছি যে, বাংলাদেশ ব্যাংকের এসএমইএসপিডি সার্কুলার নং-০৫, তারিখ: ১৩ সেপ্টেম্বর ২০২৬-এ উল্লিখিত সকল শর্তাবলি সম্পর্কে আমি অবগত আছি।',
      'আমি অবগত আছি যে অনুমোদিত হলে, প্রাপ্ত ঋণ ও অনুদান নির্ধারিত খাতে ব্যবহৃত হলে এবং ঋণ অংশ সম্পূর্ণ ও যথাসময়ে পরিশোধ করলেই কেবল অনুদানের প্রাপ্ত অর্থ অনুদান হিসেবে বিবেচিত হবে এবং ঋণের সদ্ব্যবহার না করা হলে অনুদান অংশ ঋণে পরিণত হবে এবং সম্পূর্ণ অর্থ ঋণ হিসেবে গণ্য হয়ে আমি খেলাপি হিসেবে চিহ্নিত হব। আমি অঙ্গীকার করছি যে প্রদানকৃত অর্থ ব্যক্তিগত ভোগ, জমি বা যানবাহন ক্রয়ে ব্যবহার করব না; প্রতিটি ব্যয়ের ভাউচার সংরক্ষণ করব এবং চাহিদামতো প্রতিবেদন দাখিল করব।',
    ].forEach((t) => out.push(P(t, { after: 100, align: 'both' })));
    const sw = [TEXT_W / 2, TEXT_W / 2];
    out.push(P('', { before: 600 }));
    out.push(table(sw, [
      cell(P('আবেদনকারীর স্বাক্ষর: ............................................'), sw[0], { noBorder: true }) +
      cell(P('তারিখ: ............................................', { align: 'right' }), sw[1], { noBorder: true }),
    ], { noBorder: true }));
    return out.join('');
  }

  // ---------- package parts ----------
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const FONT = 'Nirmala UI';
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:eastAsia="${FONT}" w:cs="${FONT}"/><w:sz w:val="21"/><w:szCs w:val="21"/><w:lang w:val="en-US" w:bidi="bn-BD"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`;

  // Word reads JPEG and PNG; anything else (e.g. an imported WebP) is re-encoded to JPEG.
  async function photoBytes(dataUrl) {
    if (!dataUrl) return null;
    let url = dataUrl;
    if (!/^data:image\/(jpeg|png);/.test(url)) {
      url = await new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
          const c = document.createElement('canvas'); c.width = img.width; c.height = img.height;
          c.getContext('2d').drawImage(img, 0, 0); resolve(c.toDataURL('image/jpeg', 0.9));
        };
        img.onerror = () => resolve(null);
        img.src = dataUrl;
      });
      if (!url) return null;
    }
    const ext = url.startsWith('data:image/png') ? 'png' : 'jpeg';
    const bin = atob(url.slice(url.indexOf(',') + 1));
    const u8 = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return { ext, u8 };
  }

  async function build(d) {
    d = d || {};
    const photo = await photoBytes(d.photo);
    const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="900" w:right="1000" w:bottom="900" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/></w:sectPr>';
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${body(d, photo ? 'rIdPhoto' : null)}${sect}</w:body></w:document>`;
    const files = [
      { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>` },
      { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>` },
      { name: 'docProps/core.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${x('UDYOG আবেদন ফরম — ' + (v(d, 'name_bn') || v(d, 'name_en')))}</dc:title>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</dcterms:created>
</cp:coreProperties>` },
      { name: 'word/_rels/document.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  ${photo ? `<Relationship Id="rIdPhoto" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/photo.${photo.ext}"/>` : ''}
</Relationships>` },
      { name: 'word/styles.xml', data: styles },
      { name: 'word/document.xml', data: doc },
    ];
    if (photo) files.push({ name: `word/media/photo.${photo.ext}`, data: photo.u8 });
    return zip(files);
  }

  async function download(d, baseName) {
    const blob = await build(d);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = baseName + '.docx';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1500);
  }

  window.FormDocx = { build, download };
})();
