const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, LevelFormat, TabStopType,
  TabStopPosition, ImageRun
} = require('docx');
const fs = require('fs');

const TEAL = "0F6E56";
const TEAL_LIGHT = "E1F5EE";
const TEAL_MID = "1D9E75";
const DARK = "1A1A1A";
const MID = "444441";
const LIGHT_BORDER = "CCCCCC";
const WHITE = "FFFFFF";

const border = { style: BorderStyle.SINGLE, size: 1, color: LIGHT_BORDER };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const tableWidth = 9360;

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: DARK })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: TEAL })]
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 60 },
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: MID })]
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 100 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: DARK, ...opts })]
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: DARK })]
  });
}

function spacer(lines = 1) {
  const arr = [];
  for (let i = 0; i < lines; i++) {
    arr.push(new Paragraph({ children: [new TextRun({ text: "", size: 22 })] }));
  }
  return arr;
}

function sectionDivider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL_MID, space: 1 } },
    children: [new TextRun({ text: "" })]
  });
}

function metricRow(metric, baseline, month3, month6, shaded = false) {
  const fill = shaded ? "F2FAF7" : WHITE;
  const cell = (txt, bold = false) => new TableCell({
    borders,
    width: { size: 2340, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, size: 20, font: "Arial", bold, color: bold ? TEAL : DARK })] })]
  });
  return new TableRow({ children: [cell(metric), cell(baseline), cell(month3, true), cell(month6, true)] });
}

function budgetRow(category, amount, notes, shaded = false) {
  const fill = shaded ? "F2FAF7" : WHITE;
  const cell = (txt, w, bold = false) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text: txt, size: 20, font: "Arial", bold, color: bold ? TEAL : DARK })] })]
  });
  return new TableRow({ children: [cell(category, 3120), cell(amount, 2120, true), cell(notes, 4120)] });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }, {
          level: 1, format: LevelFormat.BULLET, text: "\u25E6",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1080, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: DARK },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: TEAL },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: MID },
        paragraph: { spacing: { before: 180, after: 60 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            spacing: { after: 80 },
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL_MID, space: 4 } },
            children: [
              new TextRun({ text: "Koko Marketplace  |  Marketing Strategy & Proposal", size: 18, font: "Arial", color: MID }),
              new TextRun({ text: "  |  Meluora Consulting", size: 18, font: "Arial", color: TEAL }),
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
            spacing: { before: 80 },
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: LIGHT_BORDER, space: 4 } },
            children: [
              new TextRun({ text: "Confidential  |  June 1, 2026", size: 18, font: "Arial", color: MID }),
              new TextRun({ text: "\tPage ", size: 18, font: "Arial", color: MID }),
              new PageNumber({ size: 18, font: "Arial", color: MID }),
            ]
          })
        ]
      })
    },
    children: [
      // COVER
      ...spacer(3),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        children: [new TextRun({ text: "KOKO MARKETPLACE", size: 56, bold: true, font: "Arial", color: TEAL })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 240 },
        children: [new TextRun({ text: "Marketing Strategy & Proposal", size: 36, font: "Arial", color: MID })]
      }),
      new Table({
        width: { size: 6000, type: WidthType.DXA },
        columnWidths: [6000],
        alignment: AlignmentType.CENTER,
        rows: [new TableRow({ children: [new TableCell({
          borders: noBorders,
          shading: { fill: TEAL_LIGHT, type: ShadingType.CLEAR },
          width: { size: 6000, type: WidthType.DXA },
          margins: { top: 200, bottom: 200, left: 300, right: 300 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Prepared for: King Kosi & Koko Team", size: 22, font: "Arial", color: TEAL })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Prepared by: Meluora Consulting", size: 22, font: "Arial", color: TEAL })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Date: June 1, 2026  |  Delivery Target: June 6, 2026", size: 22, font: "Arial", color: TEAL })] }),
          ]
        })]})]}),
      ...spacer(2),

      // EXECUTIVE SUMMARY
      heading1("Executive Summary"),
      sectionDivider(),
      body("Koko is a peer-to-peer marketplace platform with a unique value proposition centered on negotiation-driven commerce — a feature deeply rooted in African cultural trading traditions. With over 2,000 downloads across multiple countries (Nigeria, Indonesia, Pakistan, Nepal, Bangladesh, Sri Lanka, Ghana, and India), the platform has achieved organic international traction."),
      body("This proposal outlines a comprehensive 6-month marketing strategy designed to:"),
      bullet("Establish strong brand identity and positioning"),
      bullet("Build direct audience engagement beyond influencer dependency"),
      bullet("Drive daily active user engagement and retention"),
      bullet("Scale user acquisition across Nigeria and international markets"),
      bullet("Create sustainable, measurable growth with clear KPIs"),
      ...spacer(1),

      // CURRENT STATE
      heading1("Current State Assessment"),
      sectionDivider(),
      heading2("Strengths"),
      bullet("Organic international adoption — 2,000+ downloads pre-marketing"),
      bullet("Unique value proposition: negotiation-based commerce"),
      bullet("Multi-country user base demonstrating global appeal"),
      bullet("Existing social media presence (Facebook, TikTok, Instagram, Snapchat)"),
      bullet("Strong local market control with Nigeria business registration"),

      heading2("Gaps & Challenges"),
      bullet("Weak brand positioning across all channels"),
      bullet("Difficult to locate official social media pages"),
      bullet("Over-reliance on influencers — not building a direct audience"),
      bullet("Inconsistent content strategy with sporadic posting patterns"),
      bullet("No clear daily engagement incentive for users"),
      bullet("Unprofessional communication channels (Gmail instead of branded domain)"),
      ...spacer(1),

      // STRATEGIC PILLARS
      heading1("Strategic Pillars"),
      sectionDivider(),

      heading2("1. Brand Positioning & Identity"),
      body("Establish Koko as the 'Power of Negotiation' marketplace — a platform that celebrates African trading culture while modernizing peer-to-peer commerce."),
      heading3("Key Messaging"),
      bullet("Tagline: \"Where Negotiation Meets Opportunity\""),
      bullet("Brand Promise: Empower buyers and sellers through transparent, conversation-driven transactions"),
      bullet("Cultural Anchor: Connect to traditional African market dynamics (haggling, relationship-building, trust)"),
      bullet("Modern Twist: Digital-first, secure, and accessible to everyone"),
      heading3("Deliverables"),
      bullet("Brand style guide (colors, typography, imagery, tone of voice)"),
      bullet("Professional branded email infrastructure (info@koko.com, support@koko.com, etc.)"),
      bullet("Unified visual identity across all platforms"),
      bullet("Brand messaging framework for all communications"),
      bullet("Logo refinement and application guidelines"),
      body("Timeline: Weeks 1–2 | Complete brand audit and style guide development", { italics: true, color: MID }),
      ...spacer(1),

      heading2("2. Social Media Strategy & Content Calendar"),
      body("Build a consistent, engaging presence on primary platforms with strategic content that drives daily app engagement."),
      heading3("Platform Strategy"),
      body("Instagram (Primary Visual Platform)", { bold: true }),
      bullet("Content Mix: 60% user success stories, 20% educational, 20% promotional"),
      bullet("Posting Frequency: 5x per week (Monday–Friday, 9 AM & 6 PM Nigeria time)"),
      bullet("Key Pillars: success stories, negotiation tips, product spotlights, testimonials"),
      body("TikTok (Viral Growth & Youth Engagement)", { bold: true }),
      bullet("Content Mix: 70% entertaining/educational, 30% promotional"),
      bullet("Posting Frequency: 3–4x per week"),
      bullet("Key Pillars: negotiation hacks, haul videos, seller spotlights, trending audio"),
      body("Facebook (Community Building)", { bold: true }),
      bullet("Content Mix: 50% community discussion, 30% educational, 20% promotional"),
      bullet("Posting Frequency: 2x per day"),
      bullet("Key Pillars: marketplace discussions, seller stories, safety tips, live Q&A"),
      body("Snapchat (Real-time Engagement)", { bold: true }),
      bullet("Content Mix: 80% behind-the-scenes, 20% promotional"),
      bullet("Posting Frequency: Daily stories"),
      body("Timeline: Weeks 1–3 | Content calendar creation & template development", { italics: true, color: MID }),
      ...spacer(1),

      heading2("3. Community Building & Grassroots Marketing"),
      body("Establish on-the-ground presence across Nigeria through student ambassadors, local influencers, and community partnerships."),
      heading3("Student Ambassador Program"),
      bullet("Structure: 1–2 ambassadors per state (~40 ambassadors across 36 states + FCT)"),
      bullet("Responsibilities: campus promotion, demo sessions, local content creation, user feedback"),
      bullet("Incentives: Monthly stipend (N10,000–N15,000), referral commissions, merchandise, recognition"),
      heading3("Community Events & Activations"),
      bullet("University roadshows — 5–6 major universities per quarter"),
      bullet("Market activations in Lagos, Abuja, Port Harcourt, Enugu, Kano"),
      bullet("Pop-up demo stations at high-traffic locations"),
      bullet("Weekly online webinars featuring sellers and success stories"),
      body("Timeline: Weeks 1–2 | Ambassador recruitment & onboarding", { italics: true, color: MID }),
      ...spacer(1),

      heading2("4. User Engagement & Retention Strategy"),
      body("Give users compelling reasons to engage daily and build habits around the platform."),
      heading3("Daily Engagement Mechanics"),
      bullet("Daily login rewards (points, badges, exclusive access)"),
      bullet("Negotiation challenges — weekly themes: \"Best Deal of the Week\""),
      bullet("Seller spotlights — feature top sellers daily"),
      bullet("Flash deals — limited-time offers to drive urgency"),
      bullet("Referral incentives — free listings and discounts for successful referrals"),
      heading3("Gamification Elements"),
      bullet("Reputation badges: Trusted Buyer, Savvy Negotiator, Top Seller"),
      bullet("Leaderboards for top negotiators and most active sellers"),
      bullet("Achievement milestones and exclusive perks for loyal users"),
      body("Timeline: Weeks 1–4 | Strategy development; Week 5+ | Implementation", { italics: true, color: MID }),
      ...spacer(1),

      heading2("5. Paid Advertising Strategy"),
      heading3("Channel Breakdown"),
      bullet("Google Ads (App Install): N500,000/month — Nigeria-wide, ages 18–45, app store targeting"),
      bullet("Meta Ads (Facebook & Instagram): N750,000/month — lookalike audiences, retargeting, testimonials"),
      bullet("TikTok Ads: N300,000/month — in-feed native ads, ages 16–35, entertainment & shopping interests"),
      body("Total Paid Budget: N1,550,000/month (~$3,100 USD)", { bold: true }),
      body("Timeline: Week 1 | Campaign setup; Week 2+ | Launch & continuous optimization", { italics: true, color: MID }),
      ...spacer(1),

      heading2("6. Email & Direct Communication Strategy"),
      heading3("Email Infrastructure"),
      bullet("Branded domain emails: info@koko.com, support@koko.com, marketing@koko.com"),
      bullet("Email marketing platform setup (Mailchimp, Brevo, or similar)"),
      bullet("Mobile-responsive branded email templates"),
      heading3("Campaign Types"),
      bullet("Weekly Newsletter: market insights, seller spotlights, trending deals"),
      bullet("Onboarding Series: 5-email sequence for new users"),
      bullet("Re-engagement Campaigns: for inactive users"),
      bullet("Promotional Campaigns: flash sales, referral bonuses, educational content"),
      body("Timeline: Weeks 1–2 | Email infrastructure setup; Week 3+ | Campaign execution", { italics: true, color: MID }),
      ...spacer(1),

      // KPIs
      heading1("Measurement & KPIs"),
      sectionDivider(),
      heading2("Growth Targets"),
      new Table({
        width: { size: tableWidth, type: WidthType.DXA },
        columnWidths: [2340, 2340, 2340, 2340],
        rows: [
          new TableRow({
            tableHeader: true,
            children: ["Metric", "Current Baseline", "3-Month Target", "6-Month Target"].map(txt =>
              new TableCell({
                borders,
                shading: { fill: TEAL, type: ShadingType.CLEAR },
                width: { size: 2340, type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: txt, size: 20, font: "Arial", bold: true, color: WHITE })] })]
              })
            )
          }),
          metricRow("Monthly Active Users", "~2,000", "10,000", "50,000", true),
          metricRow("Daily Active Users", "Unknown", "2,000", "10,000"),
          metricRow("App Downloads", "2,000", "15,000", "75,000", true),
          metricRow("Day 7 Retention", "Unknown", "40%", "55%"),
          metricRow("Day 30 Retention", "Unknown", "20%", "35%", true),
          metricRow("Instagram Followers", "—", "5,000", "25,000"),
          metricRow("TikTok Followers", "—", "10,000", "50,000", true),
          metricRow("Facebook Followers", "—", "3,000", "15,000"),
          metricRow("Successful Transactions", "Unknown", "200/month", "1,000/month", true),
        ]
      }),
      ...spacer(1),

      heading2("Paid Advertising Targets"),
      bullet("Cost Per Install (CPI): N200–300"),
      bullet("Return on Ad Spend (ROAS): 3:1 minimum"),
      bullet("Conversion Rate (Install to Active User): 30%"),
      bullet("Lifetime Value (LTV): N2,000+"),
      ...spacer(1),

      // TIMELINE
      heading1("Implementation Timeline"),
      sectionDivider(),
      new Table({
        width: { size: tableWidth, type: WidthType.DXA },
        columnWidths: [2000, 1600, 5760],
        rows: [
          new TableRow({
            tableHeader: true,
            children: ["Phase", "Period", "Key Activities"].map((txt, i) =>
              new TableCell({
                borders,
                shading: { fill: TEAL, type: ShadingType.CLEAR },
                width: { size: [2000,1600,5760][i], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: txt, size: 20, font: "Arial", bold: true, color: WHITE })] })]
              })
            )
          }),
          ...[
            ["Phase 1: Foundation", "Weeks 1–2", "Brand positioning, email setup, content calendar, ambassador recruitment, ad account setup"],
            ["Phase 2: Launch", "Weeks 3–4", "Social content execution, paid campaigns live, ambassador onboarding, email campaigns launch"],
            ["Phase 3: Optimise", "Weeks 5–12", "Continuous content, campaign monitoring, community events, user feedback, monthly reporting"],
            ["Phase 4: Scale", "Months 4–6", "Expand ambassador program, increase ad budget, strategic partnerships, international expansion planning"],
          ].map(([phase, period, activities], idx) => new TableRow({
            children: [
              new TableCell({ borders, shading: { fill: idx % 2 === 0 ? "F2FAF7" : WHITE, type: ShadingType.CLEAR }, width: { size: 2000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: phase, size: 20, font: "Arial", bold: true, color: TEAL })] })] }),
              new TableCell({ borders, shading: { fill: idx % 2 === 0 ? "F2FAF7" : WHITE, type: ShadingType.CLEAR }, width: { size: 1600, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: period, size: 20, font: "Arial", color: DARK })] })] }),
              new TableCell({ borders, shading: { fill: idx % 2 === 0 ? "F2FAF7" : WHITE, type: ShadingType.CLEAR }, width: { size: 5760, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: activities, size: 20, font: "Arial", color: DARK })] })] }),
            ]
          }))
        ]
      }),
      ...spacer(1),

      // BUDGET
      heading1("Budget Breakdown"),
      sectionDivider(),
      new Table({
        width: { size: tableWidth, type: WidthType.DXA },
        columnWidths: [3120, 2120, 4120],
        rows: [
          new TableRow({
            tableHeader: true,
            children: ["Category", "Monthly Cost", "Notes"].map((txt, i) =>
              new TableCell({
                borders,
                shading: { fill: TEAL, type: ShadingType.CLEAR },
                width: { size: [3120, 2120, 4120][i], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: txt, size: 20, font: "Arial", bold: true, color: WHITE })] })]
              })
            )
          }),
          budgetRow("Paid Advertising", "N1,550,000", "Google, Meta, TikTok", true),
          budgetRow("Meluora Services", "N1,000,000", "Team allocation & management"),
          budgetRow("Content Creation", "N400,000", "Video production, graphics", true),
          budgetRow("Ambassador Program", "N300,000", "Stipends & incentives (40 ambassadors)"),
          budgetRow("Events & Activations", "N150,000", "Roadshows, pop-ups, webinars", true),
          budgetRow("Tools & Software", "N100,000", "Email platform, analytics, scheduling"),
          budgetRow("Contingency", "N100,000", "Buffer for opportunities", true),
          new TableRow({
            children: [
              new TableCell({ borders, shading: { fill: TEAL_LIGHT, type: ShadingType.CLEAR }, width: { size: 3120, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "TOTAL", size: 22, font: "Arial", bold: true, color: TEAL })] })] }),
              new TableCell({ borders, shading: { fill: TEAL_LIGHT, type: ShadingType.CLEAR }, width: { size: 2120, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "N3,600,000", size: 22, font: "Arial", bold: true, color: TEAL })] })] }),
              new TableCell({ borders, shading: { fill: TEAL_LIGHT, type: ShadingType.CLEAR }, width: { size: 4120, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: "~$7,200 USD per month", size: 22, font: "Arial", bold: true, color: TEAL })] })] }),
            ]
          }),
        ]
      }),
      body("6-Month Total Investment: N21,600,000 (~$43,200 USD)", { bold: true }),
      ...spacer(1),

      // RISK
      heading1("Risk Mitigation"),
      sectionDivider(),
      new Table({
        width: { size: tableWidth, type: WidthType.DXA },
        columnWidths: [2800, 1280, 1280, 3800],
        rows: [
          new TableRow({
            tableHeader: true,
            children: ["Risk", "Likelihood", "Impact", "Mitigation"].map((txt, i) =>
              new TableCell({
                borders,
                shading: { fill: TEAL, type: ShadingType.CLEAR },
                width: { size: [2800,1280,1280,3800][i], type: WidthType.DXA },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                children: [new Paragraph({ children: [new TextRun({ text: txt, size: 20, font: "Arial", bold: true, color: WHITE })] })]
              })
            )
          }),
          ...[
            ["Low engagement on social media", "Medium", "High", "A/B test content, adjust posting times, increase influencer partnerships", true],
            ["High customer acquisition cost", "Medium", "Medium", "Optimise ad targeting, improve landing pages, leverage organic referrals", false],
            ["Ambassador program dropout", "Medium", "Medium", "Competitive incentives, regular communication, recognition programs", true],
            ["Competitive pressure", "High", "High", "Emphasise unique value prop, build community loyalty, innovate features", false],
            ["Platform algorithm changes", "High", "Medium", "Diversify channels, build email list, create owned community", true],
          ].map(([risk, likelihood, impact, mitigation, shaded]) => new TableRow({
            children: [
              new TableCell({ borders, shading: { fill: shaded ? "F2FAF7" : WHITE, type: ShadingType.CLEAR }, width: { size: 2800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: risk, size: 20, font: "Arial", color: DARK })] })] }),
              new TableCell({ borders, shading: { fill: shaded ? "F2FAF7" : WHITE, type: ShadingType.CLEAR }, width: { size: 1280, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: likelihood, size: 20, font: "Arial", color: likelihood === "High" ? "993C1D" : "854F0B" })] })] }),
              new TableCell({ borders, shading: { fill: shaded ? "F2FAF7" : WHITE, type: ShadingType.CLEAR }, width: { size: 1280, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: impact, size: 20, font: "Arial", color: impact === "High" ? "993C1D" : DARK })] })] }),
              new TableCell({ borders, shading: { fill: shaded ? "F2FAF7" : WHITE, type: ShadingType.CLEAR }, width: { size: 3800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: mitigation, size: 20, font: "Arial", color: DARK })] })] }),
            ]
          }))
        ]
      }),
      ...spacer(1),

      // NEXT STEPS
      heading1("Next Steps & Approval Process"),
      sectionDivider(),
      heading2("For Koko Leadership Review"),
      bullet("Review this proposal (timeline: 3–5 business days)"),
      bullet("Provide feedback on strategy, budget, and resource allocation"),
      bullet("Approve or request modifications to specific sections"),
      bullet("Confirm team availability for collaboration"),
      bullet("Schedule kickoff meeting with full team"),
      heading2("Upon Approval"),
      bullet("Meluora begins Phase 1 implementation immediately"),
      bullet("Weekly sync meetings scheduled (Mondays, 2 PM Nigeria time)"),
      bullet("Detailed project plan with daily/weekly tasks created"),
      bullet("Ambassador recruitment begins and content calendar finalised"),
      ...spacer(1),

      // CONCLUSION
      heading1("Conclusion"),
      sectionDivider(),
      body("Koko has achieved impressive organic growth with 2,000+ downloads across multiple countries. This marketing strategy is designed to capitalise on that momentum while establishing a strong, consistent brand presence that drives daily engagement and sustainable growth."),
      body("By implementing this comprehensive strategy across brand positioning, social media, community building, user engagement, and paid advertising, Koko can realistically achieve:"),
      bullet("50,000 Monthly Active Users within 6 months"),
      bullet("Strong brand recognition across Nigeria"),
      bullet("Sustainable user acquisition at profitable unit economics"),
      bullet("An engaged community that drives organic growth through referrals"),
      body("The strategy balances grassroots community building (student ambassadors, local events) with digital-first tactics (social media, paid ads), ensuring maximum reach and engagement across Nigeria's diverse market."),
      ...spacer(1),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 200, after: 100 },
        children: [new TextRun({ text: "Kenneth Nwafor  |  Meluora Team Lead", size: 22, font: "Arial", bold: true, color: TEAL })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "kenneth@meluora.com", size: 20, font: "Arial", color: MID })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/mnt/user-data/outputs/Koko_Marketing_Proposal.docx", buf);
  console.log("Done");
});