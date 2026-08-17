from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "pdf" / "Presentation-Studio-Installation-Guide.pdf"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

GREEN = colors.HexColor("#00662C")
NAVY = colors.HexColor("#00454D")
INK = colors.HexColor("#373A36")
MUTED = colors.HexColor("#63716A")
GRAPHITE = colors.HexColor("#DBDCDB")
SOFT = colors.HexColor("#F2F5F3")
LIGHT_GREEN = colors.HexColor("#E8F2EC")
FORGE = colors.HexColor("#FF9E1B")
WHITE = colors.white

APTOS_ROOT = Path("/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts")
if (APTOS_ROOT / "Aptos.ttf").is_file() and (APTOS_ROOT / "Aptos-Bold.ttf").is_file():
    pdfmetrics.registerFont(TTFont("Aptos", APTOS_ROOT / "Aptos.ttf"))
    pdfmetrics.registerFont(TTFont("Aptos-Bold", APTOS_ROOT / "Aptos-Bold.ttf"))
    FONT_REGULAR = "Aptos"
    FONT_BOLD = "Aptos-Bold"
else:
    FONT_REGULAR = "Helvetica"
    FONT_BOLD = "Helvetica-Bold"


class GuideDoc(BaseDocTemplate):
    def __init__(self, filename):
        super().__init__(
            filename,
            pagesize=letter,
            leftMargin=0.72 * inch,
            rightMargin=0.72 * inch,
            topMargin=0.72 * inch,
            bottomMargin=0.62 * inch,
            title="Presentation Studio Source Installation Guide",
            author="Presentation Studio",
            subject="Source setup, MCP configuration, and project security",
        )
        frame = Frame(self.leftMargin, self.bottomMargin, self.width, self.height, id="content")
        self.addPageTemplates(PageTemplate(id="guide", frames=[frame], onPage=self.draw_page))

    def draw_page(self, canvas, doc):
        canvas.saveState()
        canvas.setFillColor(GREEN)
        canvas.rect(0, letter[1] - 0.16 * inch, letter[0], 0.16 * inch, stroke=0, fill=1)
        if doc.page > 1:
            canvas.setFont(FONT_BOLD, 7.5)
            canvas.setFillColor(NAVY)
            canvas.drawString(0.72 * inch, letter[1] - 0.42 * inch, "PRESENTATION STUDIO")
            canvas.setFont(FONT_REGULAR, 7.5)
            canvas.setFillColor(MUTED)
            canvas.drawRightString(letter[0] - 0.72 * inch, letter[1] - 0.42 * inch, "SOURCE INSTALLATION GUIDE")
        canvas.setStrokeColor(GRAPHITE)
        canvas.line(0.72 * inch, 0.42 * inch, letter[0] - 0.72 * inch, 0.42 * inch)
        canvas.setFont(FONT_REGULAR, 7.2)
        canvas.setFillColor(MUTED)
        canvas.drawString(0.72 * inch, 0.24 * inch, "Local-first source installer - no unsigned app package")
        canvas.drawRightString(letter[0] - 0.72 * inch, 0.24 * inch, f"{doc.page}")
        canvas.restoreState()


base = getSampleStyleSheet()
styles = {
    "cover_kicker": ParagraphStyle("cover_kicker", parent=base["Normal"], fontName=FONT_BOLD, fontSize=9, leading=11, textColor=GREEN, spaceAfter=10, tracking=1.2),
    "cover_title": ParagraphStyle("cover_title", parent=base["Title"], fontName=FONT_BOLD, fontSize=31, leading=33, textColor=NAVY, alignment=TA_LEFT, spaceAfter=12),
    "cover_subtitle": ParagraphStyle("cover_subtitle", parent=base["Normal"], fontName=FONT_REGULAR, fontSize=13, leading=19, textColor=INK, spaceAfter=22),
    "h1": ParagraphStyle("h1", parent=base["Heading1"], fontName=FONT_BOLD, fontSize=23, leading=27, textColor=NAVY, spaceBefore=3, spaceAfter=11),
    "h2": ParagraphStyle("h2", parent=base["Heading2"], fontName=FONT_BOLD, fontSize=13, leading=16, textColor=GREEN, spaceBefore=12, spaceAfter=6),
    "body": ParagraphStyle("body", parent=base["BodyText"], fontName=FONT_REGULAR, fontSize=9.5, leading=14.2, textColor=INK, spaceAfter=8),
    "small": ParagraphStyle("small", parent=base["BodyText"], fontName=FONT_REGULAR, fontSize=8, leading=11.5, textColor=MUTED),
    "step": ParagraphStyle("step", parent=base["BodyText"], fontName=FONT_REGULAR, fontSize=9.2, leading=13.2, textColor=INK),
    "code": ParagraphStyle("code", parent=base["Code"], fontName="Courier", fontSize=7.8, leading=11, textColor=NAVY, backColor=SOFT, borderPadding=8, borderColor=GRAPHITE, borderWidth=0.5, spaceBefore=4, spaceAfter=9),
    "code_one_line": ParagraphStyle("code_one_line", parent=base["Code"], fontName="Courier", fontSize=5.25, leading=9, textColor=NAVY, backColor=SOFT, borderPadding=7, borderColor=GRAPHITE, borderWidth=0.5, spaceBefore=4, spaceAfter=9, splitLongWords=False),
    "prompt": ParagraphStyle("prompt", parent=base["Code"], fontName="Courier", fontSize=6.6, leading=9.1, textColor=NAVY, backColor=SOFT, borderPadding=8, borderColor=GRAPHITE, borderWidth=0.5, spaceBefore=4, spaceAfter=9),
    "callout": ParagraphStyle("callout", parent=base["BodyText"], fontName=FONT_REGULAR, fontSize=9, leading=13.5, textColor=INK),
    "center": ParagraphStyle("center", parent=base["BodyText"], fontName=FONT_BOLD, fontSize=9, leading=13, textColor=NAVY, alignment=TA_CENTER),
    "step_number": ParagraphStyle("step_number", parent=base["BodyText"], fontName=FONT_BOLD, fontSize=9, leading=13, textColor=WHITE, alignment=TA_CENTER),
    "table_header": ParagraphStyle("table_header", parent=base["BodyText"], fontName=FONT_BOLD, fontSize=9, leading=13, textColor=WHITE, alignment=TA_CENTER),
}


def p(text, style="body"):
    return Paragraph(text, styles[style])


def callout(title, text, accent=GREEN):
    table = Table([[Paragraph(title, ParagraphStyle("callout-title", parent=styles["body"], fontName=FONT_BOLD, textColor=accent, spaceAfter=3)), p(text, "callout")]], colWidths=[1.15 * inch, 5.7 * inch])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT_GREEN),
        ("BOX", (0, 0), (-1, -1), 0.6, accent),
        ("LINEBEFORE", (0, 0), (0, -1), 4, accent),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
    ]))
    return table


def steps(items):
    rows = []
    for index, item in enumerate(items, 1):
        number = Table([[p(str(index), "step_number")]], colWidths=[0.29 * inch], rowHeights=[0.29 * inch])
        number.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), GREEN), ("TEXTCOLOR", (0, 0), (-1, -1), WHITE), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BOX", (0, 0), (-1, -1), 0, GREEN)]))
        rows.append([number, p(item, "step")])
    table = Table(rows, colWidths=[0.42 * inch, 6.45 * inch], hAlign="LEFT")
    table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 9), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 6)]))
    return table


story = []
story += [Spacer(1, 0.38 * inch), p("LOCAL-FIRST PRESENTATION PRODUCTION", "cover_kicker"), p("Presentation Studio", "cover_title"), p("Source Installation Guide", "cover_title"), p("Set up the Electron desktop app, connect any standard MCP client, and understand the project and export security boundaries.", "cover_subtitle")]

cover_table = Table([
    [p("INSTALLATION", "table_header"), p("PROJECTS", "table_header"), p("AI CONTROL", "table_header")],
    [p("One-line source install<br/>macOS and Windows", "small"), p("Self-contained<br/>optional encryption", "small"), p("Local MCP<br/>human review gates", "small")],
], colWidths=[2.28 * inch] * 3, rowHeights=[0.35 * inch, 0.6 * inch])
cover_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
    ("BACKGROUND", (0, 1), (-1, 1), SOFT), ("BOX", (0, 0), (-1, -1), 0.8, GRAPHITE),
    ("INNERGRID", (0, 0), (-1, -1), 0.5, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
]))
story += [cover_table, Spacer(1, 0.32 * inch), callout("Important", "This release uses a one-line source installer, not an unsigned DMG, PKG, MSI, EXE, or app-store package. It installs prerequisites and verifies the app without disabling operating-system protections.", FORGE), Spacer(1, 1.15 * inch), p("Version 0.2.1 - August 2026", "small"), PageBreak()]

story += [p("Before you begin", "h1"), p("Presentation Studio is designed for local presentation review and redesign. Imported PowerPoint files are copied into a self-contained project, audited without changing the originals, and exported only as new files through a human-controlled save or export action."), p("Requirements", "h2")]
requirements = Table([
    [p("Operating system", "small"), p("macOS or Windows", "body")],
    [p("Runtime", "small"), p("Checked and installed automatically when Node.js 22.13 or newer with npm is unavailable", "body")],
    [p("Disk space", "small"), p("Enough local space for source decks, packaged Resources, recovery files, and exported copies", "body")],
    [p("Network", "small"), p("Needed during installation; the running app blocks non-local network requests", "body")],
], colWidths=[1.38 * inch, 5.45 * inch])
requirements.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, -1), SOFT), ("BOX", (0, 0), (-1, -1), 0.6, GRAPHITE), ("INNERGRID", (0, 0), (-1, -1), 0.4, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 9), ("RIGHTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
story += [requirements, p("Install Version 0.2.1 on macOS", "h2"), p("Paste this complete line into Terminal. It does not require Git."), p("curl -fsSL https://raw.githubusercontent.com/adammalin/Presentation-Studio/codex/web-slide-design-engine/scripts/install-macos.sh | /bin/zsh", "code_one_line"), p("Install Version 0.2.1 on Windows", "h2"), p("Paste this complete line into PowerShell. It does not change execution policy."), p("irm https://raw.githubusercontent.com/adammalin/Presentation-Studio/codex/web-slide-design-engine/scripts/install-windows.ps1 | iex", "code_one_line"), callout("Run again to update", "Close Presentation Studio, then run the same one-line command. The installer verifies a staged copy before replacing the current managed app and keeps one previous managed copy."), callout("Data boundary", "Do not place client decks, manuscripts, project packages, extracted text, previews, or exports in the installation source folder. Projects and exports remain outside the managed app source."), PageBreak()]

story += [p("What the installer does", "h1"), steps([
    "Checks the operating system, processor architecture, install location, Node.js, and npm.",
    "When needed, downloads the official portable Node.js 22.13 runtime and verifies it against the official SHA-256 manifest.",
    "Downloads Presentation Studio 0.2.1 from the isolated release branch without requiring Git.",
    "Runs the locked dependency install, automated tests, repository data-safety scan, and production renderer build in a staging folder.",
    "Activates the verified app, creates a reusable launcher, and starts Presentation Studio.",
]), p("Managed locations", "h2")]
managed_locations = Table([
    [p("macOS", "small"), p("~/Applications/Presentation Studio", "body"), p("Launch Presentation Studio.command", "small")],
    [p("Windows", "small"), p("%LOCALAPPDATA%\\Presentation Studio", "body"), p("Launch Presentation Studio.cmd", "small")],
], colWidths=[1.0 * inch, 3.05 * inch, 2.8 * inch])
managed_locations.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, -1), SOFT), ("BOX", (0, 0), (-1, -1), 0.6, GRAPHITE), ("INNERGRID", (0, 0), (-1, -1), 0.4, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 9), ("RIGHTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
story += [managed_locations, Spacer(1, 0.14 * inch), callout("Microsoft PowerPoint", "PowerPoint is optional for launching the app but required for PowerPoint-native rendering and final native validation. Licensed Microsoft software is not installed automatically.", FORGE), callout("System security", "The installers do not alter Gatekeeper, quarantine settings, PowerShell execution policy, SmartScreen, or other protections. Follow your organization's approved software process if execution is blocked."), p("Developer checkout", "h2"), p("Developers who need Git history can use the manual checkout procedure in README.md. The one-line installer is the default for ordinary users."), PageBreak()]

story += [p("Connect an MCP client", "h1"), p("Presentation Studio exposes a standard STDIO MCP server. It can be used by any compatible AI client; the app does not contain a provider-specific model or API-key workflow."), steps([
    "Open Presentation Studio and keep the desktop app running.",
    "Print the standard configuration snippet with the command below.",
    "Add the <b>presentation-studio</b> entry to your MCP client's normal <b>mcpServers</b> configuration, then restart that client if required.",
    "Turn on the visible <b>AI session</b> switch in Presentation Studio only when you want the current project and deck context available.",
]), p("node scripts/configure-mcp.mjs", "code"), p("To merge the entry into an explicitly selected JSON file:", "body"), p("node scripts/configure-mcp.mjs --write /absolute/path/to/mcp-config.json", "code"), p("MCP safety model", "h2")]
mcp_table = Table([
    [p("Allowed", "table_header"), p("Human only", "table_header")],
    [p("Check app status<br/>Read authorized deck context<br/>Stage semantic design changes<br/>Build private local candidates<br/>Record qualification evidence", "small"), p("Confirm a template exception<br/>Apply or reject a legacy proposal<br/>Save a project<br/>Export a PowerPoint copy<br/>Distribute an output", "small")],
], colWidths=[3.42 * inch, 3.42 * inch])
mcp_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), GREEN), ("BACKGROUND", (1, 0), (1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE), ("BACKGROUND", (0, 1), (-1, 1), SOFT), ("BOX", (0, 0), (-1, -1), 0.7, GRAPHITE), ("INNERGRID", (0, 0), (-1, -1), 0.5, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9)]))
story += [mcp_table, Spacer(1, 0.16 * inch), callout("Current MCP boundary", "Authorized MCP models can inspect, stage semantic design changes, build private local candidates, and record qualification evidence. They cannot overwrite an original, save the project, export to a user destination, or distribute an output."), Spacer(1, 0.12 * inch), callout("Local connection", "The STDIO server reaches the active app through a per-session token on a loopback-only bridge. The token descriptor is written with private local permissions and removed when the app closes."), PageBreak()]

session_prompt = """Connect to the Presentation Studio MCP and work in the currently open project.

First, read the Presentation Studio design contract, check the app status, inventory the authorized project Resources, and inspect the installed ORNL Template Pack. Confirm that you can read the source content - not merely filenames or metadata. If anything required is inaccessible, tell me exactly what must be shared or attached. Do not invent missing content.

Create a polished, editable, 16:9 ORNL presentation from the supplied source materials.

Content direction:
- Organize the material into a clear narrative using assertion-evidence slides.
- Give each slide one primary takeaway and supporting evidence.
- You may condense source prose, but preserve technical meaning, names, numbers, units, qualifications, and attribution.
- Preserve approved or locked copy exactly. Do not introduce unsupported claims, data, diagrams, or conclusions.
- Infer routine structure and design choices. Ask only about genuine audience, technical, content-authority, or approval ambiguities.

Design direction:
- Use Aptos and the current approved ORNL Template Pack.
- For a new title slide, use an approved ORNL title layout and edit only intended placeholders. Never alter its artwork, marks, master, or layout.
- Make substantive whole-slide composition decisions using shared Studio recipes and compatible ORNL layouts. Do not merely keep the source arrangement or make text smaller.
- Establish one deck-wide system for titles, spacing, alignment, figures, captions, tables, colors, and repeated components.
- Use authorized Resource images only when they support the message. Preserve technical figures as relationship-aware groups unless a verified editable reconstruction is clearer.
- Keep tables editable and readable; preserve meaning-bearing colors.

Workflow:
1. Develop the narrative and slide plan.
2. Create the presentation in the single central Studio HTML/CSS scene.
3. Build the complete editable PowerPoint candidate.
4. Inspect the PowerPoint-native contact sheet and every full-size candidate slide.
5. Run Found issues -> Fixing -> Rechecking original intent. Correct overflow, alignment, hierarchy, spacing, tables, missing imagery, and message drift.
6. Do not call the presentation ready while any blocker or major visual issue remains.

Use your best design judgment and minimize routine questions. Leave the completed central design visible for my review. Do not save or export the final PowerPoint until I explicitly request it."""

story += [p("Start a ChatGPT Desktop design session", "h1"), p("Use this after Presentation Studio is open, the MCP connection is enabled, and the approved project material is in place. Copy the complete prompt below into a fresh ChatGPT Desktop conversation."), p(session_prompt.replace("\n", "<br/>"), "prompt"), callout("Current 0.2.1 limitation", "Presentation Studio can package document Resources, but MCP currently returns only authorized Resource metadata and bounded image previews - not extracted document text. It also cannot create a brand-new deck directly from document-only Resources. For now, attach the same cleared source documents directly to ChatGPT Desktop and add a starter PowerPoint to Presentation Studio. Use only material approved for the selected AI environment.", FORGE), PageBreak()]

story += [p("Projects, encryption, and verification", "h1"), p("Project formats", "h2")]
project_table = Table([
    [p(".pstudio", "table_header"), p(".pstudio-secure", "table_header")],
    [p("ZIP-based self-contained package<br/>Canonical project JSON<br/>Immutable-by-hash Resources<br/>Portable without original file paths", "small"), p("Encrypts the complete .pstudio payload<br/>AES-256-GCM<br/>PBKDF2-SHA-256 with 250,000 iterations<br/>Password cannot be recovered", "small")],
], colWidths=[3.42 * inch, 3.42 * inch])
project_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE), ("BACKGROUND", (0, 1), (-1, 1), SOFT), ("BOX", (0, 0), (-1, -1), 0.7, GRAPHITE), ("INNERGRID", (0, 0), (-1, -1), 0.5, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9)]))
story += [project_table, Spacer(1, 0.12 * inch), callout("Encryption boundary", "Encrypted project files protect packaged JSON and Resources. They do not encrypt the external originals or separately exported PowerPoint, PDF, SVG, or PNG files.", FORGE), p("Optional developer verification", "h2"), p("npm run quality<br/>npm run desktop:smoke", "code"), p("Expected result", "h2"), p("Tests should pass, the repository scan should report no tracked client artifacts, Vite should produce a production bundle, and the Electron smoke check should save a valid renderer capture before exiting."), p("Troubleshooting", "h2"), KeepTogether([p("App reports that setup is missing", "body"), p("Run the one-line installer again. In a developer checkout, rerun the platform setup script from the repository root.", "small")]), Spacer(1, 7), KeepTogether([p("MCP says the app is unavailable", "body"), p("Open Presentation Studio first. Restart the desktop app if the local runtime descriptor is stale, then retry the MCP tool.", "small")]), Spacer(1, 7), KeepTogether([p("A deck needs manual review", "body"), p("Macros, embedded OLE objects, external relationships, uncertain templates, and ambiguous formatting intentionally stop automated cleanup. Preserve the source and resolve the finding in the app.", "small")])]

GuideDoc(str(OUTPUT)).build(story)
print(OUTPUT)
