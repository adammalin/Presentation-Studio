from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
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
            canvas.setFont("Helvetica-Bold", 7.5)
            canvas.setFillColor(NAVY)
            canvas.drawString(0.72 * inch, letter[1] - 0.42 * inch, "PRESENTATION STUDIO")
            canvas.setFont("Helvetica", 7.5)
            canvas.setFillColor(MUTED)
            canvas.drawRightString(letter[0] - 0.72 * inch, letter[1] - 0.42 * inch, "SOURCE INSTALLATION GUIDE")
        canvas.setStrokeColor(GRAPHITE)
        canvas.line(0.72 * inch, 0.42 * inch, letter[0] - 0.72 * inch, 0.42 * inch)
        canvas.setFont("Helvetica", 7.2)
        canvas.setFillColor(MUTED)
        canvas.drawString(0.72 * inch, 0.24 * inch, "Local-first source distribution - no app installer")
        canvas.drawRightString(letter[0] - 0.72 * inch, 0.24 * inch, f"{doc.page}")
        canvas.restoreState()


base = getSampleStyleSheet()
styles = {
    "cover_kicker": ParagraphStyle("cover_kicker", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=9, leading=11, textColor=GREEN, spaceAfter=10, tracking=1.2),
    "cover_title": ParagraphStyle("cover_title", parent=base["Title"], fontName="Helvetica-Bold", fontSize=31, leading=33, textColor=NAVY, alignment=TA_LEFT, spaceAfter=12),
    "cover_subtitle": ParagraphStyle("cover_subtitle", parent=base["Normal"], fontName="Helvetica", fontSize=13, leading=19, textColor=INK, spaceAfter=22),
    "h1": ParagraphStyle("h1", parent=base["Heading1"], fontName="Helvetica-Bold", fontSize=23, leading=27, textColor=NAVY, spaceBefore=3, spaceAfter=11),
    "h2": ParagraphStyle("h2", parent=base["Heading2"], fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=GREEN, spaceBefore=12, spaceAfter=6),
    "body": ParagraphStyle("body", parent=base["BodyText"], fontName="Helvetica", fontSize=9.5, leading=14.2, textColor=INK, spaceAfter=8),
    "small": ParagraphStyle("small", parent=base["BodyText"], fontName="Helvetica", fontSize=8, leading=11.5, textColor=MUTED),
    "step": ParagraphStyle("step", parent=base["BodyText"], fontName="Helvetica", fontSize=9.2, leading=13.2, textColor=INK),
    "code": ParagraphStyle("code", parent=base["Code"], fontName="Courier", fontSize=7.8, leading=11, textColor=NAVY, backColor=SOFT, borderPadding=8, borderColor=GRAPHITE, borderWidth=0.5, spaceBefore=4, spaceAfter=9),
    "callout": ParagraphStyle("callout", parent=base["BodyText"], fontName="Helvetica", fontSize=9, leading=13.5, textColor=INK),
    "center": ParagraphStyle("center", parent=base["BodyText"], fontName="Helvetica-Bold", fontSize=9, leading=13, textColor=NAVY, alignment=TA_CENTER),
    "table_header": ParagraphStyle("table_header", parent=base["BodyText"], fontName="Helvetica-Bold", fontSize=9, leading=13, textColor=WHITE, alignment=TA_CENTER),
}


def p(text, style="body"):
    return Paragraph(text, styles[style])


def callout(title, text, accent=GREEN):
    table = Table([[Paragraph(title, ParagraphStyle("callout-title", parent=styles["body"], fontName="Helvetica-Bold", textColor=accent, spaceAfter=3)), p(text, "callout")]], colWidths=[1.15 * inch, 5.7 * inch])
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
        number = Table([[p(str(index), "center")]], colWidths=[0.29 * inch], rowHeights=[0.29 * inch])
        number.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), GREEN), ("TEXTCOLOR", (0, 0), (-1, -1), WHITE), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BOX", (0, 0), (-1, -1), 0, GREEN)]))
        rows.append([number, p(item, "step")])
    table = Table(rows, colWidths=[0.42 * inch, 6.45 * inch], hAlign="LEFT")
    table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BOTTOMPADDING", (0, 0), (-1, -1), 9), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 6)]))
    return table


story = []
story += [Spacer(1, 0.38 * inch), p("LOCAL-FIRST PRESENTATION PRODUCTION", "cover_kicker"), p("Presentation Studio", "cover_title"), p("Source Installation Guide", "cover_title"), p("Set up the Electron desktop app, connect any standard MCP client, and understand the project and export security boundaries.", "cover_subtitle")]

cover_table = Table([
    [p("INSTALLATION", "table_header"), p("PROJECTS", "table_header"), p("AI CONTROL", "table_header")],
    [p("Source setup only<br/>macOS and Windows", "small"), p("Self-contained<br/>optional encryption", "small"), p("Local MCP<br/>human review gates", "small")],
], colWidths=[2.28 * inch] * 3, rowHeights=[0.35 * inch, 0.6 * inch])
cover_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
    ("BACKGROUND", (0, 1), (-1, 1), SOFT), ("BOX", (0, 0), (-1, -1), 0.8, GRAPHITE),
    ("INNERGRID", (0, 0), (-1, -1), 0.5, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
]))
story += [cover_table, Spacer(1, 0.32 * inch), callout("Important", "There is no DMG, PKG, MSI, EXE, app-store package, or other application installer in this release. The setup scripts install and verify the source checkout without disabling operating-system protections.", FORGE), Spacer(1, 1.15 * inch), p("Version 0.1 - August 2026", "small"), PageBreak()]

story += [p("Before you begin", "h1"), p("Presentation Studio is designed for local presentation review. Imported PowerPoint files are copied into a self-contained project, audited without changing the originals, and exported only as new files after a human approves a cleanup plan."), p("Requirements", "h2")]
requirements = Table([
    [p("Operating system", "small"), p("macOS or Windows", "body")],
    [p("Runtime", "small"), p("Node.js 22.13 or newer with npm", "body")],
    [p("Disk space", "small"), p("Enough local space for source decks, packaged Resources, recovery files, and exported copies", "body")],
    [p("Network", "small"), p("Needed only to obtain source dependencies when they are not already cached; the running app blocks non-local network requests", "body")],
], colWidths=[1.38 * inch, 5.45 * inch])
requirements.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, -1), SOFT), ("BOX", (0, 0), (-1, -1), 0.6, GRAPHITE), ("INNERGRID", (0, 0), (-1, -1), 0.4, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 9), ("RIGHTPADDING", (0, 0), (-1, -1), 9), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
story += [requirements, p("What setup verifies", "h2"), p("The setup scripts run the locked dependency install, automated tests, repository data-safety scan, and production renderer build. A successful setup proves the current checkout can run; it does not approve any presentation for distribution."), callout("Data boundary", "Do not place client decks, manuscripts, project packages, extracted text, previews, or exports in the repository. Synthetic fixtures are generated locally and ignored by Git."), PageBreak()]

story += [p("Install on macOS", "h1"), p("Use the included source setup script. It checks the installed Node.js version and stops with a clear message if the requirement is not met."), steps([
    "Download or clone the Presentation Studio repository to a local folder.",
    "Double-click <b>scripts/setup-macos.command</b>, or run it from Terminal.",
    "Wait for dependency installation, tests, data-safety checks, and the production build to pass.",
    "Start the app with <b>scripts/start-macos.command</b>.",
]), p("Terminal commands", "h2"), p("./scripts/setup-macos.command<br/>./scripts/start-macos.command", "code"), callout("macOS security", "The script does not change Gatekeeper, quarantine attributes, or other operating-system protections. Follow your organization's approved software process if local policy blocks source execution.", FORGE), p("Install on Windows", "h1"), steps([
    "Download or clone the repository and open PowerShell in that folder.",
    "Run the setup script shown below. It performs the same checks as macOS.",
    "Start the app with the separate start script after setup passes.",
]), p("& .\\scripts\\setup-windows.ps1<br/>& .\\scripts\\start-windows.ps1", "code"), callout("Windows security", "The script does not change PowerShell execution policy, SmartScreen, or other protections. Use an approved policy exception or source-install process instead of disabling controls.", FORGE), PageBreak()]

story += [p("Connect an MCP client", "h1"), p("Presentation Studio exposes a standard STDIO MCP server. It can be used by any compatible AI client; the app does not contain a provider-specific model or API-key workflow."), steps([
    "Open Presentation Studio and keep the desktop app running.",
    "Print the standard configuration snippet with the command below.",
    "Add the <b>presentation-studio</b> entry to your MCP client's normal <b>mcpServers</b> configuration, then restart that client if required.",
    "Turn on the visible <b>AI session</b> switch in Presentation Studio only when you want the current audit metadata available.",
]), p("node scripts/configure-mcp.mjs", "code"), p("To merge the entry into an explicitly selected JSON file:", "body"), p("node scripts/configure-mcp.mjs --write /absolute/path/to/mcp-config.json", "code"), p("MCP safety model", "h2")]
mcp_table = Table([
    [p("Allowed", "table_header"), p("Human only", "table_header")],
    [p("Check app status<br/>List authorized deck metadata<br/>Read deterministic audit findings<br/>Stage a bounded cleanup proposal", "small"), p("Confirm a template<br/>Apply or reject a proposal<br/>Save a project<br/>Export a PowerPoint copy<br/>Distribute an output", "small")],
], colWidths=[3.42 * inch, 3.42 * inch])
mcp_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (0, 0), GREEN), ("BACKGROUND", (1, 0), (1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE), ("BACKGROUND", (0, 1), (-1, 1), SOFT), ("BOX", (0, 0), (-1, -1), 0.7, GRAPHITE), ("INNERGRID", (0, 0), (-1, -1), 0.5, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9)]))
story += [mcp_table, Spacer(1, 0.16 * inch), callout("Local connection", "The STDIO server reaches the active app through a per-session token on a loopback-only bridge. The token descriptor is written with private local permissions and removed when the app closes."), PageBreak()]

story += [p("Projects, encryption, and verification", "h1"), p("Project formats", "h2")]
project_table = Table([
    [p(".pstudio", "table_header"), p(".pstudio-secure", "table_header")],
    [p("ZIP-based self-contained package<br/>Canonical project JSON<br/>Immutable-by-hash Resources<br/>Portable without original file paths", "small"), p("Encrypts the complete .pstudio payload<br/>AES-256-GCM<br/>PBKDF2-SHA-256 with 250,000 iterations<br/>Password cannot be recovered", "small")],
], colWidths=[3.42 * inch, 3.42 * inch])
project_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), WHITE), ("BACKGROUND", (0, 1), (-1, 1), SOFT), ("BOX", (0, 0), (-1, -1), 0.7, GRAPHITE), ("INNERGRID", (0, 0), (-1, -1), 0.5, GRAPHITE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 10), ("RIGHTPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 9), ("BOTTOMPADDING", (0, 0), (-1, -1), 9)]))
story += [project_table, Spacer(1, 0.12 * inch), callout("Encryption boundary", "Encrypted project files protect packaged JSON and Resources. They do not encrypt the external originals or separately exported PowerPoint, PDF, SVG, or PNG files.", FORGE), p("Verify the checkout", "h2"), p("npm test<br/>npm run check:data-safety<br/>npm run build<br/>npm run desktop:smoke", "code"), p("Expected result", "h2"), p("Tests should pass, the repository scan should report no tracked client artifacts, Vite should produce a production bundle, and the Electron smoke check should save a valid renderer capture before exiting."), p("Troubleshooting", "h2"), KeepTogether([p("App reports that setup is missing", "body"), p("Run the platform setup script from the repository root so that node_modules and the production bundle are created.", "small")]), Spacer(1, 7), KeepTogether([p("MCP says the app is unavailable", "body"), p("Open Presentation Studio first. Restart the desktop app if the local runtime descriptor is stale, then retry the MCP tool.", "small")]), Spacer(1, 7), KeepTogether([p("A deck needs manual review", "body"), p("Macros, embedded OLE objects, external relationships, uncertain templates, and ambiguous formatting intentionally stop automated cleanup. Preserve the source and resolve the finding in the app.", "small")])]

GuideDoc(str(OUTPUT)).build(story)
print(OUTPUT)
