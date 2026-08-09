import { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle } from "docx";
import { writeFileSync } from "fs";
const thinBorder = { style: BorderStyle.SINGLE, size: 6, color: "222222" };
const sectionBorder = { bottom: thinBorder };
const h = (text) =>
  new Paragraph({
    spacing: { before: 280, after: 120 },
    border: sectionBorder,
    children: [
      new TextRun({
        text,
        bold: true,
        size: 22,
        font: "Calibri",
        color: "111111",
      }),
    ],
  });
const p = (text, opts = {}) =>
  new Paragraph({
    spacing: { after: opts.after ?? 80 },
    children: [
      new TextRun({
        text,
        size: 20,
        font: "Calibri",
        color: "222222",
        bold: opts.bold || false,
      }),
    ],
  });
const mixed = (parts, spacingAfter = 80) =>
  new Paragraph({
    spacing: { after: spacingAfter },
    children: parts.map(
      (part) =>
        new TextRun({
          text: part.text,
          size: 20,
          font: "Calibri",
          color: "222222",
          bold: !!part.bold,
        })
    ),
  });
const doc = new Document({
  sections: [
    {
      properties: {
        page: {
          margin: { top: 720, right: 720, bottom: 720, left: 720 },
        },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [
            new TextRun({
              text: "RANCE ADRIEL M. PASCUA",
              bold: true,
              size: 32,
              font: "Calibri",
              color: "111111",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 40 },
          children: [
            new TextRun({
              text: "3rd Year BSN, Universidad de Dagupan",
              size: 20,
              font: "Calibri",
              color: "333333",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 40 },
          children: [
            new TextRun({
              text: "Poblacion, Aguilar, Pangasinan | Born July 31, 2005",
              size: 18,
              font: "Calibri",
              color: "444444",
            }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          border: sectionBorder,
          children: [
            new TextRun({
              text: "adrielrancepascua@gmail.com  |  09454903281  |  github.com/adrielrancepascua-dev",
              size: 18,
              font: "Calibri",
              color: "444444",
            }),
          ],
        }),
        h("ABOUT"),
        p(
          "I am a 21 year old Bachelor of Science in Nursing student at Universidad de Dagupan. I build TypeScript and React apps for local businesses and nursing education. I already ship work for a paying client, and I keep shipping demos for hospitality and cafe operations."
        ),
        h("EDUCATION"),
        mixed([{ text: "Bachelor of Science in Nursing", bold: true }]),
        p("Universidad de Dagupan, 3rd Year"),
        p("2024 to 2028 (expected)"),
        p("Presidential Scholar, First Year", { after: 40 }),
        h("EXPERIENCE"),
        mixed([{ text: "Freelance Developer", bold: true }, { text: ", Papers & Petals (paying client)" }]),
        p(
          "Built and maintains cafe-flowershop, the back office for Papers & Petals across Dagupan, San Carlos, and Urdaneta. Covers staff login, multi branch orders, inventory transfers, expenses, and branch sales reports. Live demo: flower-backoffice-demo.vercel.app"
        ),
        mixed([{ text: "Shop Attendant", bold: true }, { text: ", Buks' Computershop, 2014 to 2021" }]),
        p(
          "Watched over the family internet cafe. Learned customer service, basic PC upkeep, and how to keep a small tech business running day to day."
        ),
        h("SELECTED PROJECTS"),
        mixed([{ text: "NursePath / block9nurseapp", bold: true }]),
        p(
          "Clinical reference PWA for student nurses. Vital signs reference ranges, OTC med lookup, and common nursing calculators. Works offline after first load. Currently under faculty review for adoption at Universidad de Dagupan. Live: block9nurseapp.vercel.app"
        ),
        mixed([{ text: "NursePath Faculty Dashboard", bold: true }]),
        p(
          "Usage analytics for faculty. Sessions, feature use, per student rollups, and CSV export. Live: nursepath-dashboard.vercel.app"
        ),
        mixed([{ text: "Hotel System Demo", bold: true }]),
        p(
          "Front desk ops suite for a small Philippine hotel. Check in, folio charges, checkout, room status, and owner reports in PHP. Live: hotel-system-demo.vercel.app"
        ),
        mixed([{ text: "Other shipped work", bold: true }]),
        p(
          "Cafe and shop demos (cafe-brewsco, chef-s-cafe, OUR-cafe, kanto-cafe, flowershop-demo), booking UI (Stay-Awhile), sales tools (sales), personal site (portfolio), and automation experiments (n8n-rance). Full list: github.com/adrielrancepascua-dev"
        ),
        h("SKILLS"),
        p(
          "TypeScript, JavaScript, React, HTML, CSS, Vite, Supabase, Vercel, Git, Progressive Web Apps, basic operations tooling for shops and clinics"
        ),
        h("CONTACT"),
        p("Open to freelance builds, campus tech projects, and roles where nursing context meets software."),
        p("Email: adrielrancepascua@gmail.com"),
        p("Phone: 09454903281"),
        p("GitHub: https://github.com/adrielrancepascua-dev", { after: 0 }),
      ],
    },
  ],
});
const buffer = await Packer.toBuffer(doc);
const out = "Rance-Adriel-Pascua-Resume.docx";
try {
  writeFileSync(out, buffer);
  console.log("Wrote " + out);
} catch (err) {
  const alt = "Rance-Adriel-Pascua-Resume-updated.docx";
  writeFileSync(alt, buffer);
  console.log("Locked " + out + "; wrote " + alt + " instead");
}
