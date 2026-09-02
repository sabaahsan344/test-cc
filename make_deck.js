const pptxgen = require("pptxgenjs");

// ---------- palette ----------
const GREEN_DK = "1F4D2E"; // dominant
const GREEN_MD = "2F7A47";
const SAGE = "9BBF8A";
const TINT = "EDF3EA"; // card background
const GOLD = "E0A526"; // accent
const INK = "1A2620";
const MUTED = "5F6F66";
const WHITE = "FFFFFF";

const HEAD = "Cambria";
const BODY = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.author = "University of the Punjab - Semester Rules";
pres.title = "Attendance and Evaluation Rules";

const M = 0.55; // side margin
const W = 10 - M * 2; // usable width = 8.9

// ---------- helpers ----------
function slideTitle(s, text, sub) {
  s.addText(text, {
    x: M, y: 0.34, w: W, h: 0.62,
    fontFace: HEAD, fontSize: 32, bold: true, color: GREEN_DK,
    isTextBox: true, margin: 0, valign: "middle",
  });
  if (sub) {
    s.addText(sub, {
      x: M, y: 0.98, w: W, h: 0.34,
      fontFace: BODY, fontSize: 14, italic: true, color: MUTED,
      isTextBox: true, margin: 0, valign: "middle",
    });
  }
}

function card(s, opts) {
  s.addShape(pres.ShapeType.roundRect, {
    x: opts.x, y: opts.y, w: opts.w, h: opts.h,
    fill: { color: opts.fill || TINT },
    line: { color: opts.fill || TINT, width: 0 },
    rectRadius: 0.09,
    shadow: { type: "outer", color: "1F4D2E", opacity: 0.13, blur: 8, offset: 2, angle: 90 },
  });
}

// the repeating motif: a filled circle with a short label inside
function circleBadge(s, x, y, d, label, opts) {
  opts = opts || {};
  s.addShape(pres.ShapeType.ellipse, {
    x: x, y: y, w: d, h: d,
    fill: { color: opts.fill || GREEN_DK },
    line: { color: opts.fill || GREEN_DK, width: 0 },
  });
  s.addText(label, {
    x: x, y: y, w: d, h: d,
    fontFace: HEAD, fontSize: opts.size || 16, bold: true,
    color: opts.color || WHITE, align: "center", valign: "middle",
    isTextBox: true, margin: 0,
  });
}

function footer(s, n) {
  s.addText(String(n), {
    x: 10 - M - 0.6, y: 5.06, w: 0.6, h: 0.28,
    fontFace: BODY, fontSize: 10, color: MUTED, align: "right",
    isTextBox: true, margin: 0,
  });
}

function sectionSlide(title, kicker, points) {
  const s = pres.addSlide();
  s.background = { color: GREEN_DK };
  s.addText(kicker, {
    x: M, y: 1.55, w: W, h: 0.34,
    fontFace: BODY, fontSize: 15, bold: true, color: GOLD, charSpacing: 2,
    isTextBox: true, margin: 0,
  });
  s.addText(title, {
    x: M, y: 1.95, w: 6.3, h: 1.0,
    fontFace: HEAD, fontSize: 40, bold: true, color: WHITE,
    isTextBox: true, margin: 0, valign: "top",
  });
  s.addText(points, {
    x: M, y: 3.1, w: 6.3, h: 1.0,
    fontFace: BODY, fontSize: 15, color: SAGE,
    isTextBox: true, margin: 0,
  });
  // motif: three stacked circles on the right
  circleBadge(s, 7.35, 1.85, 1.05, "", { fill: GREEN_MD });
  circleBadge(s, 8.05, 2.75, 0.75, "", { fill: SAGE });
  circleBadge(s, 7.05, 3.15, 0.5, "", { fill: GOLD });
  return s;
}

let pageNo = 0;
function contentSlide(title, sub) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  slideTitle(s, title, sub);
  pageNo += 1;
  footer(s, pageNo);
  return s;
}

// =====================================================================
// 1. TITLE
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: GREEN_DK };

  s.addText("UNIVERSITY OF THE PUNJAB", {
    x: M, y: 1.05, w: W, h: 0.3,
    fontFace: BODY, fontSize: 14, bold: true, color: GOLD, charSpacing: 3,
    isTextBox: true, margin: 0,
  });
  s.addText("Attendance & Evaluation Rules", {
    x: M, y: 1.5, w: 6.6, h: 1.5,
    fontFace: HEAD, fontSize: 44, bold: true, color: WHITE,
    isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.0,
  });
  s.addText(
    "Semester Rules and Regulations for Undergraduate Studies\nApproved 27-07-2023  •  Applicable from Fall 2023",
    {
      x: M, y: 3.1, w: 6.1, h: 0.8,
      fontFace: BODY, fontSize: 14, color: SAGE,
      isTextBox: true, margin: 0, lineSpacingMultiple: 1.3,
    }
  );

  // motif cluster
  circleBadge(s, 7.55, 1.55, 1.4, "75%", { fill: GOLD, color: GREEN_DK, size: 25 });
  s.addText("minimum attendance", {
    x: 6.9, y: 3.02, w: 2.7, h: 0.28,
    fontFace: BODY, fontSize: 11, color: SAGE, align: "center",
    isTextBox: true, margin: 0,
  });
  circleBadge(s, 7.8, 3.45, 0.9, "50%", { fill: GREEN_MD, color: WHITE, size: 16 });
  s.addText("marks needed to pass", {
    x: 6.9, y: 4.42, w: 2.7, h: 0.28,
    fontFace: BODY, fontSize: 11, color: SAGE, align: "center",
    isTextBox: true, margin: 0,
  });

  s.addNotes(
    "This presentation explains two things every student must know: how much attendance is required, and how students are marked and graded. All rules are taken from the University of the Punjab Semester Rules and Regulations for Undergraduate Studies, effective Fall 2023."
  );
}

// =====================================================================
// 2. AGENDA
// =====================================================================
{
  const s = contentSlide("What this presentation covers", "Two topics, explained in simple steps");

  const items = [
    ["1", "Attendance rules", "How many classes you must attend, and what happens if you miss too many."],
    ["2", "Absence & penalties", "The 25% allowance, the FW grade, fines and special cases."],
    ["3", "Marks division", "Sessional work, mid term and final term weightage."],
    ["4", "Grades & promotion", "Grade table, GPA, probation and passing your degree."],
  ];

  const cw = (W - 0.4) / 2; // 4.25
  const ch = 1.55;
  items.forEach((it, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * (cw + 0.4);
    const y = 1.5 + row * (ch + 0.32);
    card(s, { x, y, w: cw, h: ch });
    circleBadge(s, x + 0.28, y + 0.28, 0.52, it[0], { size: 17 });
    s.addText(it[1], {
      x: x + 0.95, y: y + 0.28, w: cw - 1.25, h: 0.36,
      fontFace: HEAD, fontSize: 17, bold: true, color: GREEN_DK,
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.95, y: y + 0.7, w: cw - 1.25, h: 0.68,
      fontFace: BODY, fontSize: 13, color: MUTED,
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
    });
  });

  s.addNotes("Give students a quick map of the talk before going into detail.");
}

// =====================================================================
// 3. SECTION: ATTENDANCE
// =====================================================================
{
  const s = sectionSlide("Attendance Rules", "PART ONE", "Section 9 of the Semester Regulations");
  s.addNotes("Section 9 of the regulations deals with class attendance.");
}

// =====================================================================
// 4. THE 75% RULE
// =====================================================================
{
  const s = contentSlide("The 75% rule", "The single most important attendance rule");

  card(s, { x: M, y: 1.5, w: 3.4, h: 2.75, fill: GREEN_DK });
  s.addText("75%", {
    x: M, y: 1.85, w: 3.4, h: 1.0,
    fontFace: HEAD, fontSize: 66, bold: true, color: GOLD, align: "center",
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("of the classes held in a course\nmust be attended", {
    x: M + 0.3, y: 2.9, w: 2.8, h: 0.9,
    fontFace: BODY, fontSize: 14, color: WHITE, align: "center",
    isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
  });

  const bullets = [
    "Attendance is counted separately for every course, not for the whole semester.",
    "Without 75% attendance, you are not allowed to sit in the mid term and final term exams.",
    "If your attendance is short at mid term, you may be allowed to sit provisionally, but you must complete 75% before the final exam.",
  ];
  const bx = M + 3.4 + 0.4;
  const bw = W - 3.4 - 0.4;
  bullets.forEach((t, i) => {
    const y = 1.55 + i * 0.92;
    circleBadge(s, bx, y + 0.04, 0.3, "", { fill: SAGE });
    s.addText(t, {
      x: bx + 0.46, y: y, w: bw - 0.46, h: 0.82,
      fontFace: BODY, fontSize: 14, color: INK,
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.18,
    });
  });

  s.addNotes("Attendance is per course. A student may have 90% in one subject and 60% in another - the 60% subject is still a problem.");
}

// =====================================================================
// 5. 25% ABSENCE ALLOWANCE + FW
// =====================================================================
{
  const s = contentSlide("How much absence is allowed?", "Maximum 25% absence - and not all at once");

  const items = [
    ["25%", "Maximum absence", "You may miss up to 25% of classes, but not in one long stretch."],
    ["FW", "Forced Withdrawal", "Use up the full allowance and your low-attendance course is marked FW."],
    ["0", "Effect on GPA", "FW is not counted in GPA or CGPA, but it appears on your transcript."],
  ];
  const cw = (W - 0.6) / 3; // 2.5667
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.3);
    card(s, { x, y: 1.5, w: cw, h: 2.35 });
    s.addText(it[0], {
      x: x, y: 1.68, w: cw, h: 0.65,
      fontFace: HEAD, fontSize: 38, bold: true, color: GREEN_DK, align: "center",
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(it[1], {
      x: x + 0.18, y: 2.38, w: cw - 0.36, h: 0.3,
      fontFace: BODY, fontSize: 13, bold: true, color: GREEN_MD, align: "center",
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(it[2], {
      x: x + 0.22, y: 2.72, w: cw - 0.44, h: 1.0,
      fontFace: BODY, fontSize: 12.5, color: MUTED, align: "center",
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
    });
  });

  card(s, { x: M, y: 4.08, w: W, h: 0.72, fill: GREEN_DK });
  s.addText(
    "A course with an FW grade must be repeated with a junior session, and you pay the course fee and exam fee again.",
    {
      x: M + 0.3, y: 4.08, w: W - 0.6, h: 0.72,
      fontFace: BODY, fontSize: 13.5, color: WHITE,
      isTextBox: true, margin: 0, valign: "middle",
    }
  );

  s.addNotes("From the 2nd semester onwards, availing the full 25% absence in a course leads to an FW grade in that low-attendance course.");
}

// =====================================================================
// 6. 5% RELAXATION AND FINE
// =====================================================================
{
  const s = contentSlide("Extra 5% relaxation - with a fine", "Only the Head of Department can give it");

  card(s, { x: M, y: 1.5, w: 4.3, h: 1.4 });
  circleBadge(s, M + 0.28, 1.78, 0.55, "+5%", { size: 12 });
  s.addText("Relaxation on request", {
    x: M + 1.0, y: 1.72, w: 3.1, h: 0.3,
    fontFace: HEAD, fontSize: 16, bold: true, color: GREEN_DK,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("The HoD may relax attendance by 5% if the student applies for it.", {
    x: M + 1.0, y: 2.06, w: 3.1, h: 0.7,
    fontFace: BODY, fontSize: 13, color: MUTED,
    isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
  });

  card(s, { x: M + 4.6, y: 1.5, w: 4.3, h: 1.4, fill: GREEN_DK });
  s.addText("Rs. 2,000", {
    x: M + 4.85, y: 1.68, w: 3.8, h: 0.5,
    fontFace: HEAD, fontSize: 30, bold: true, color: GOLD,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("fine per short-attendance course\n(maximum Rs. 5,000 in one semester)", {
    x: M + 4.85, y: 2.2, w: 3.8, h: 0.6,
    fontFace: BODY, fontSize: 13, color: WHITE,
    isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
  });

  card(s, { x: M, y: 3.1, w: W, h: 1.32, fill: TINT });
  s.addText("So what is the real limit?", {
    x: M + 0.35, y: 3.28, w: W - 0.7, h: 0.32,
    fontFace: HEAD, fontSize: 17, bold: true, color: GREEN_DK,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "25% allowed absence  +  5% relaxation by HoD  =  30% total", options: { bullet: true, breakLine: true } },
      { text: "This 30% includes both absents (without permission) and leave (with permission).", options: { bullet: true, breakLine: false } },
    ],
    {
      x: M + 0.35, y: 3.66, w: W - 0.7, h: 0.9,
      fontFace: BODY, fontSize: 13.5, color: INK,
      isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 6,
    }
  );

  s.addNotes("Remember: the 5% is not automatic. The student must request it, the HoD must approve it, and a fine is payable.");
}

// =====================================================================
// 7. SPECIAL CASES
// =====================================================================
{
  const s = contentSlide("Special cases", "The rules are slightly different for these students");

  const rows = [
    ["H", "Hajj or maternity", "The full 25% absence may be taken in one stretch, with prior approval of the HoD."],
    ["S", "National / international sportspersons", "Attendance is counted after removing the classes held on the days spent in the games. The department arranges make-up lectures before exams."],
    ["L", "Late admission", "If admitted after 5 weeks but before the mid term, a special mid term exam is held. Admission after the mid term is moved to the next session."],
  ];
  rows.forEach((r, i) => {
    const y = 1.5 + i * 1.18;
    card(s, { x: M, y: y, w: W, h: 1.02 });
    circleBadge(s, M + 0.26, y + 0.25, 0.52, r[0], { size: 17, fill: i === 2 ? GOLD : GREEN_DK, color: i === 2 ? GREEN_DK : WHITE });
    s.addText(r[1], {
      x: M + 0.95, y: y + 0.14, w: W - 1.25, h: 0.3,
      fontFace: HEAD, fontSize: 15.5, bold: true, color: GREEN_DK,
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(r[2], {
      x: M + 0.95, y: y + 0.46, w: W - 1.25, h: 0.48,
      fontFace: BODY, fontSize: 12.5, color: MUTED,
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.12,
    });
  });

  s.addNotes("Sportspersons need verification from the Director of Sports and prior approval of the Chairman, Director or Principal. They must still sit the mid and final exams with the class.");
}

// =====================================================================
// 8. TWO WEEKS ABSENCE WARNING
// =====================================================================
{
  const s = contentSlide("Absent for two weeks in a row", "Missing all classes of all courses for two weeks is very serious");

  card(s, { x: M, y: 1.5, w: (W - 0.4) / 2, h: 2.2, fill: GREEN_DK });
  s.addText("Semester 1", {
    x: M + 0.3, y: 1.7, w: 3.85, h: 0.32,
    fontFace: BODY, fontSize: 13, bold: true, color: GOLD, charSpacing: 1,
    isTextBox: true, margin: 0,
  });
  s.addText("Admission cancelled", {
    x: M + 0.3, y: 2.06, w: 3.85, h: 0.45,
    fontFace: HEAD, fontSize: 22, bold: true, color: WHITE,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("You may apply for admission again next year, if you meet the conditions.", {
    x: M + 0.3, y: 2.6, w: 3.85, h: 0.85,
    fontFace: BODY, fontSize: 13, color: SAGE,
    isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
  });

  const x2 = M + (W - 0.4) / 2 + 0.4;
  card(s, { x: x2, y: 1.5, w: (W - 0.4) / 2, h: 2.2 });
  s.addText("Semester 2 onwards", {
    x: x2 + 0.3, y: 1.7, w: 3.85, h: 0.32,
    fontFace: BODY, fontSize: 13, bold: true, color: GREEN_MD, charSpacing: 1,
    isTextBox: true, margin: 0,
  });
  s.addText("Semester freezed", {
    x: x2 + 0.3, y: 2.06, w: 3.85, h: 0.45,
    fontFace: HEAD, fontSize: 22, bold: true, color: GREEN_DK,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("You cannot continue the session and must repeat the same semester with your junior batch.", {
    x: x2 + 0.3, y: 2.6, w: 3.85, h: 0.85,
    fontFace: BODY, fontSize: 13, color: MUTED,
    isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
  });

  s.addText(
    "Also note: admission is cancelled if a student does not attend a single lecture during the first four weeks of the semester.",
    {
      x: M, y: 3.95, w: W, h: 0.7,
      fontFace: BODY, fontSize: 13.5, italic: true, color: GREEN_DK,
      isTextBox: true, margin: 0, valign: "top",
    }
  );

  s.addNotes("This applies when the student is absent from ALL classes of ALL courses for two consecutive weeks without approved leave.");
}

// =====================================================================
// 9. ATTENDANCE RECAP
// =====================================================================
{
  const s = contentSlide("Attendance: numbers to remember", "Keep these five figures in mind all semester");

  const nums = [
    ["75%", "minimum attendance in every course"],
    ["25%", "maximum absence allowed"],
    ["5%", "extra relaxation by the HoD"],
    ["2,000", "rupees fine per short course"],
    ["5,000", "rupees maximum fine per semester"],
  ];
  const cw = (W - 4 * 0.22) / 5; // 1.604
  nums.forEach((n, i) => {
    const x = M + i * (cw + 0.22);
    card(s, { x, y: 1.6, w: cw, h: 2.1, fill: i === 0 ? GREEN_DK : TINT });
    s.addText(n[0], {
      x: x, y: 1.8, w: cw, h: 0.6,
      fontFace: HEAD, fontSize: n[0].length > 3 ? 24 : 30, bold: true,
      color: i === 0 ? GOLD : GREEN_DK, align: "center",
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(n[1], {
      x: x + 0.14, y: 2.48, w: cw - 0.28, h: 1.05,
      fontFace: BODY, fontSize: 11.5, color: i === 0 ? WHITE : MUTED, align: "center",
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
    });
  });

  s.addText("Your monthly attendance record is displayed on the department notice board - check it regularly.", {
    x: M, y: 4.1, w: W, h: 0.5,
    fontFace: BODY, fontSize: 13.5, italic: true, color: GREEN_DK,
    isTextBox: true, margin: 0, valign: "middle",
  });

  s.addNotes("Teachers send the attendance statement to the HoD at the end of every month, and the cumulative record is displayed on the notice board.");
}

// =====================================================================
// 10. SECTION: EVALUATION
// =====================================================================
{
  const s = sectionSlide("Evaluation & Exams", "PART TWO", "Sections 10 to 13 of the Semester Regulations");
  s.addNotes("Now we move to how students are tested, marked and graded.");
}

// =====================================================================
// 11. MARKS DIVISION (chart)
// =====================================================================
{
  const s = contentSlide("How your marks are divided", "Every course is marked out of 100");

  s.addChart(
    pres.ChartType.doughnut,
    [{
      name: "Weightage",
      labels: ["Sessional activities", "Mid term exam", "Final term exam"],
      values: [25, 35, 40],
    }],
    {
      x: 0.35, y: 1.35, w: 4.1, h: 3.5,
      chartColors: [SAGE, GREEN_MD, GREEN_DK],
      holeSize: 52,
      showLegend: false,
      showValue: false,
      showTitle: false,
      dataBorder: { pt: 3, color: WHITE },
    }
  );

  // label inside the doughnut hole
  s.addText("100", {
    x: 1.75, y: 2.72, w: 1.3, h: 0.5,
    fontFace: HEAD, fontSize: 26, bold: true, color: GREEN_DK, align: "center",
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("marks", {
    x: 1.75, y: 3.18, w: 1.3, h: 0.26,
    fontFace: BODY, fontSize: 11, color: MUTED, align: "center",
    isTextBox: true, margin: 0, valign: "middle",
  });

  const legend = [
    [SAGE, "25%", "Sessional activities", "Home assignments, quizzes and class work. At least one home assignment is required."],
    [GREEN_MD, "35%", "Mid term examination", "Held after eight weeks of the semester. Paper duration: one and a half hours."],
    [GREEN_DK, "40%", "Final term examination", "Held at the end of the 17th week. Paper duration: two hours."],
  ];
  const lx = 4.75;
  legend.forEach((l, i) => {
    const y = 1.42 + i * 1.15;
    s.addShape(pres.ShapeType.roundRect, {
      x: lx, y: y, w: 0.72, h: 0.4,
      fill: { color: l[0] }, line: { color: l[0], width: 0 }, rectRadius: 0.06,
    });
    s.addText(l[1], {
      x: lx, y: y, w: 0.72, h: 0.4,
      fontFace: HEAD, fontSize: 14, bold: true, color: i === 0 ? GREEN_DK : WHITE,
      align: "center", valign: "middle", isTextBox: true, margin: 0,
    });
    s.addText(l[2], {
      x: lx + 0.88, y: y - 0.02, w: 4.0, h: 0.32,
      fontFace: HEAD, fontSize: 15, bold: true, color: GREEN_DK,
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(l[3], {
      x: lx + 0.88, y: y + 0.33, w: 4.0, h: 0.62,
      fontFace: BODY, fontSize: 12.5, color: MUTED,
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.12,
    });
  });

  s.addNotes("It is compulsory to appear in both the mid term and the final term, and to complete the sessional activities.");
}

// =====================================================================
// 12. PASSING A COURSE
// =====================================================================
{
  const s = contentSlide("Passing a course", "What you need to clear a subject");

  card(s, { x: M, y: 1.5, w: 3.15, h: 2.6, fill: GREEN_DK });
  s.addText("50%", {
    x: M, y: 1.95, w: 3.15, h: 0.85,
    fontFace: HEAD, fontSize: 58, bold: true, color: GOLD, align: "center",
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText("minimum marks\n(grade 'D')", {
    x: M, y: 2.85, w: 3.15, h: 0.7,
    fontFace: BODY, fontSize: 14, color: WHITE, align: "center",
    isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.2,
  });

  const pts = [
    "The 50% is counted on the total of sessional work, mid term and final term together.",
    "You must appear in both exams. Missing one without a valid reason means zero marks in it.",
    "Below 50%, or absent, means grade 'F' - the course must be repeated.",
    "There is no choice in questions in the mid term and final term papers.",
  ];
  const bx = M + 3.15 + 0.4;
  const bw = W - 3.15 - 0.4;
  pts.forEach((t, i) => {
    const y = 1.5 + i * 0.72;
    circleBadge(s, bx, y + 0.06, 0.28, "", { fill: i < 2 ? GREEN_MD : SAGE });
    s.addText(t, {
      x: bx + 0.44, y: y, w: bw - 0.44, h: 0.66,
      fontFace: BODY, fontSize: 13.5, color: INK,
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
    });
  });

  s.addNotes("A 'D' grade is a pass, but it may still need to be repeated if the CGPA falls below 2.00.");
}

// =====================================================================
// 13. EXAM SCHEDULE
// =====================================================================
{
  const s = contentSlide("When are the exams held?", "A semester is 18 weeks long");

  const steps = [
    ["Weeks 1-16", "Teaching", "Sixteen weeks of lectures. The teacher gives the course outline within the first 7 days."],
    ["After week 8", "Mid term exam", "1.5 hours. Worth 35% of the course marks."],
    ["Week 18", "Final term exam", "2 hours, held at the end of the 17th week. Worth 40%."],
  ];
  const cw = (W - 0.6) / 3;
  steps.forEach((st, i) => {
    const x = M + i * (cw + 0.3);
    card(s, { x, y: 1.75, w: cw, h: 2.35, fill: i === 2 ? GREEN_DK : TINT });
    circleBadge(s, x + cw / 2 - 0.26, 1.5, 0.52, String(i + 1), {
      size: 17, fill: i === 2 ? GOLD : GREEN_DK, color: i === 2 ? GREEN_DK : WHITE,
    });
    s.addText(st[0], {
      x: x + 0.2, y: 2.15, w: cw - 0.4, h: 0.28,
      fontFace: BODY, fontSize: 12, bold: true, color: i === 2 ? GOLD : GREEN_MD,
      align: "center", charSpacing: 1, isTextBox: true, margin: 0,
    });
    s.addText(st[1], {
      x: x + 0.2, y: 2.46, w: cw - 0.4, h: 0.38,
      fontFace: HEAD, fontSize: 18, bold: true, color: i === 2 ? WHITE : GREEN_DK,
      align: "center", isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(st[2], {
      x: x + 0.24, y: 2.92, w: cw - 0.48, h: 1.05,
      fontFace: BODY, fontSize: 12.5, color: i === 2 ? SAGE : MUTED,
      align: "center", isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.15,
    });
  });

  s.addText("The date sheet for both exams is notified by the Chairman, Director or Principal.", {
    x: M, y: 4.35, w: W, h: 0.4,
    fontFace: BODY, fontSize: 13, italic: true, color: MUTED,
    isTextBox: true, margin: 0, valign: "middle",
  });

  s.addNotes("Exams are held on consecutive days excluding holidays.");
}

// =====================================================================
// 14. AFTER THE EXAM
// =====================================================================
{
  const s = contentSlide("After the exam: your rights", "You are allowed to see your paper and question your marks");

  const rows = [
    ["1", "Paper showing", "Teachers must show you the scripts of your sessional work, mid term and final term. The scripts are collected back straight after showing."],
    ["2", "Still not satisfied?", "Apply in writing to the Chairman, Director or Principal within one week of the paper showing date."],
    ["3", "Rechecking fee", "Rs. 1,000 per course. The case goes to the Departmental Examination Committee, and its decision is final."],
  ];
  rows.forEach((r, i) => {
    const y = 1.42 + i * 1.12;
    card(s, { x: M, y: y, w: W, h: 0.96, fill: i === 2 ? GREEN_DK : TINT });
    circleBadge(s, M + 0.26, y + 0.22, 0.5, r[0], {
      size: 16, fill: i === 2 ? GOLD : GREEN_DK, color: i === 2 ? GREEN_DK : WHITE,
    });
    s.addText(r[1], {
      x: M + 0.93, y: y + 0.11, w: W - 1.23, h: 0.3,
      fontFace: HEAD, fontSize: 15.5, bold: true, color: i === 2 ? WHITE : GREEN_DK,
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(r[2], {
      x: M + 0.93, y: y + 0.43, w: W - 1.23, h: 0.48,
      fontFace: BODY, fontSize: 12.5, color: i === 2 ? SAGE : MUTED,
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.12,
    });
  });

  s.addText("Results are also displayed on the department notice board.", {
    x: M, y: 4.72, w: 7.0, h: 0.32,
    fontFace: BODY, fontSize: 12, italic: true, color: MUTED,
    isTextBox: true, margin: 0, valign: "middle",
  });

  s.addNotes("Scripts are kept for three months after the result is declared.");
}

// =====================================================================
// 15. MISSED AN EXAM
// =====================================================================
{
  const s = contentSlide("If you miss an exam", "A special examination may be arranged - but only on three conditions");

  const conds = [
    ["1", "Tell the department", "Within one week of the start of the examination."],
    ["2", "Get approval", "The Head of Department must approve your application."],
    ["3", "Pay the fee", "Pay the special exam fee notified by the treasurer."],
  ];
  const cw = (W - 0.6) / 3;
  conds.forEach((c, i) => {
    const x = M + i * (cw + 0.3);
    card(s, { x, y: 1.45, w: cw, h: 1.85 });
    circleBadge(s, x + 0.24, 1.68, 0.48, c[0], { size: 16 });
    s.addText(c[1], {
      x: x + 0.22, y: 2.26, w: cw - 0.44, h: 0.32,
      fontFace: HEAD, fontSize: 15, bold: true, color: GREEN_DK,
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(c[2], {
      x: x + 0.22, y: 2.6, w: cw - 0.44, h: 0.62,
      fontFace: BODY, fontSize: 12, color: MUTED,
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.12,
    });
  });

  card(s, { x: M, y: 3.62, w: W, h: 1.28, fill: GREEN_DK });
  s.addText("Important limits", {
    x: M + 0.35, y: 3.76, w: W - 0.7, h: 0.3,
    fontFace: HEAD, fontSize: 15, bold: true, color: GOLD,
    isTextBox: true, margin: 0, valign: "middle",
  });
  s.addText(
    [
      { text: "No valid reason, or fee not paid within two weeks, means zero marks in that exam.", options: { bullet: true, breakLine: true } },
      { text: "The special exam is not available to improve a D or F grade, or if you already sat the exam.", options: { bullet: true, breakLine: false } },
    ],
    {
      x: M + 0.35, y: 4.1, w: W - 0.7, h: 0.72,
      fontFace: BODY, fontSize: 12.5, color: WHITE,
      isTextBox: true, margin: 0, valign: "top", paraSpaceAfter: 5,
    }
  );

  s.addNotes("You must also have the required attendance to be allowed a special examination. The special exam is held within two weeks of the regular exam.");
}

// =====================================================================
// 16. GRADING TABLE
// =====================================================================
{
  const s = contentSlide("The grading table", "Your percentage becomes a letter grade and a grade point");

  const head = ["Marks (%)", "Grade", "Points", "Marks (%)", "Grade", "Points"];
  const left = [
    ["85 & above", "A", "4.00"],
    ["80 - 84", "A-", "3.70"],
    ["75 - 79", "B+", "3.30"],
    ["70 - 74", "B", "3.00"],
    ["65 - 69", "B-", "2.70"],
  ];
  const right = [
    ["61 - 64", "C+", "2.30"],
    ["58 - 60", "C", "2.00"],
    ["55 - 57", "C-", "1.70"],
    ["50 - 54", "D", "1.00"],
    ["Below 50 / Absent", "F", "0.00"],
  ];

  const rows = [head.map((h) => ({
    text: h,
    options: { bold: true, color: WHITE, fill: { color: GREEN_DK }, fontSize: 12, align: "center" },
  }))];
  for (let i = 0; i < 5; i++) {
    const bg = i % 2 === 0 ? TINT : WHITE;
    const cells = left[i].concat(right[i]).map((c, j) => ({
      text: c,
      options: {
        fill: { color: bg },
        color: j === 1 || j === 4 ? GREEN_DK : INK,
        bold: j === 1 || j === 4,
        fontSize: 12,
        align: "center",
      },
    }));
    rows.push(cells);
  }

  s.addTable(rows, {
    x: M, y: 1.45, w: W, colW: [1.75, 1.0, 1.0, 1.9, 1.0, 1.0],
    rowH: 0.3, fontFace: BODY, valign: "middle",
    border: { type: "solid", color: "DCE5DA", pt: 1 },
  });

  const notes = [
    ["W", "Withdrawal - not counted in GPA and not shown on the transcript."],
    ["FW", "Forced Withdrawal - not counted in GPA, but shown on the transcript."],
    ["I", "Incomplete - must be cleared by a special exam within one month, or 0 marks."],
  ];
  notes.forEach((n, i) => {
    const y = 3.5 + i * 0.44;
    s.addShape(pres.ShapeType.roundRect, {
      x: M, y: y, w: 0.6, h: 0.34,
      fill: { color: SAGE }, line: { color: SAGE, width: 0 }, rectRadius: 0.05,
    });
    s.addText(n[0], {
      x: M, y: y, w: 0.6, h: 0.34,
      fontFace: HEAD, fontSize: 13, bold: true, color: GREEN_DK,
      align: "center", valign: "middle", isTextBox: true, margin: 0,
    });
    s.addText(n[1], {
      x: M + 0.75, y: y, w: W - 0.75, h: 0.34,
      fontFace: BODY, fontSize: 12.5, color: INK,
      isTextBox: true, margin: 0, valign: "middle",
    });
  });

  s.addNotes("A fraction of a mark is rounded up to the next whole mark: 64.1 or 64.9 both become 65.");
}

// =====================================================================
// 17. GPA, PROMOTION, PROBATION
// =====================================================================
{
  const s = contentSlide("Promotion and probation", "Your CGPA decides whether you move to the next semester");

  const head = ["", "Promoted", "On probation", "Dropped"];
  const body = [
    ["1st semester", "GPA above 2.00", "1.50 to below 2.00", "GPA below 1.50"],
    ["2nd semester onwards", "CGPA above 2.00", "1.70 to below 2.00", "CGPA below 1.70"],
  ];
  const rows = [head.map((h, j) => ({
    text: h,
    options: {
      bold: true, color: WHITE, fill: { color: GREEN_DK }, fontSize: 12.5,
      align: j === 0 ? "left" : "center",
    },
  }))];
  body.forEach((r, i) => {
    rows.push(r.map((c, j) => ({
      text: c,
      options: {
        fill: { color: i % 2 === 0 ? TINT : WHITE },
        color: j === 0 ? GREEN_DK : INK,
        bold: j === 0,
        fontSize: 12.5,
        align: j === 0 ? "left" : "center",
      },
    })));
  });

  s.addTable(rows, {
    x: M, y: 1.45, w: W, colW: [2.3, 2.2, 2.2, 2.2],
    rowH: 0.42, fontFace: BODY, valign: "middle",
    border: { type: "solid", color: "DCE5DA", pt: 1 },
    margin: [0, 0.12, 0, 0.12],
  });

  const facts = [
    ["2.00", "The minimum CGPA needed for your degree."],
    ["2", "Probation is allowed only twice in the whole programme."],
    ["Repeat", "'F' courses must be repeated, and 'D' courses too if your CGPA is low. The better grade counts."],
  ];
  const cw = (W - 0.6) / 3;
  facts.forEach((f, i) => {
    const x = M + i * (cw + 0.3);
    card(s, { x, y: 3.2, w: cw, h: 1.55, fill: i === 0 ? GREEN_DK : TINT });
    s.addText(f[0], {
      x: x + 0.15, y: 3.35, w: cw - 0.3, h: 0.45,
      fontFace: HEAD, fontSize: f[0].length > 4 ? 20 : 28, bold: true,
      color: i === 0 ? GOLD : GREEN_DK, align: "center",
      isTextBox: true, margin: 0, valign: "middle",
    });
    s.addText(f[1], {
      x: x + 0.2, y: 3.84, w: cw - 0.4, h: 0.82,
      fontFace: BODY, fontSize: 11.5, color: i === 0 ? WHITE : MUTED, align: "center",
      isTextBox: true, margin: 0, valign: "top", lineSpacingMultiple: 1.12,
    });
  });

  s.addNotes("After two probations, a student who still cannot reach a 2.00 CGPA is dropped from the rolls.");
}

// =====================================================================
// 18. CLOSING
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: GREEN_DK };

  s.addText("REMEMBER THIS", {
    x: M, y: 0.6, w: W, h: 0.3,
    fontFace: BODY, fontSize: 13, bold: true, color: GOLD, charSpacing: 3,
    isTextBox: true, margin: 0,
  });
  s.addText("Five things to take away", {
    x: M, y: 0.95, w: W, h: 0.6,
    fontFace: HEAD, fontSize: 34, bold: true, color: WHITE,
    isTextBox: true, margin: 0, valign: "middle",
  });

  const take = [
    ["1", "Attend at least 75% of every course - attendance is checked subject by subject."],
    ["2", "Too many absents means an FW grade and repeating the course with juniors."],
    ["3", "Marks are 25% sessional, 35% mid term and 40% final term."],
    ["4", "You need 50% (grade D) to pass a course, and 2.00 CGPA for your degree."],
    ["5", "Check the notice board every month for your attendance and results."],
  ];
  take.forEach((t, i) => {
    const y = 1.78 + i * 0.62;
    circleBadge(s, M, y, 0.42, t[0], { fill: GOLD, color: GREEN_DK, size: 15 });
    s.addText(t[1], {
      x: M + 0.62, y: y - 0.01, w: W - 0.62, h: 0.45,
      fontFace: BODY, fontSize: 14.5, color: WHITE,
      isTextBox: true, margin: 0, valign: "middle",
    });
  });

  s.addText(
    "Source: University of the Punjab, Semester Rules and Regulations for Undergraduate Studies (Fall 2023)",
    {
      x: M, y: 5.0, w: W, h: 0.3,
      fontFace: BODY, fontSize: 10, color: SAGE,
      isTextBox: true, margin: 0, valign: "middle",
    }
  );

  s.addNotes("Close by reminding students that these rules protect them as much as they bind them - knowing them prevents most problems.");
}

pres.writeFile({ fileName: process.argv[2] || "Attendance_and_Evaluation_Rules.pptx" }).then((f) =>
  console.log("wrote " + f)
);
