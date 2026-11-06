"""
Generate TwitterSense-AI Project Report in KIIT format.
"""
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from copy import deepcopy

PROJECT_NAME = "TwitterSense-AI: Social Media Intelligence Platform"
PROJECT_SHORT = "TwitterSense-AI"
MONTH_YEAR = "November 2026"
ACADEMIC_YEAR = "2025-2026"

# Placeholders — replace with actual names/roll numbers before submission
MEMBERS = [
    ("GROUP MEMBER A", "ROLL NUMBER A"),
    ("GROUP MEMBER B", "ROLL NUMBER B"),
    ("GROUP MEMBER C", "ROLL NUMBER C"),
    ("GROUP MEMBER D", "ROLL NUMBER D"),
]
GUIDE_NAME = "GUIDE NAME"


def set_run_font(run, size=12, bold=False, italic=False, name="Times New Roman", color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color


def add_para(doc, text="", size=12, bold=False, italic=False, align="left",
             space_before=0, space_after=6, line_spacing=1.5, first_line=0):
    p = doc.add_paragraph()
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    if first_line:
        pf.first_line_indent = Pt(first_line)
    if text:
        run = p.add_run(text)
        set_run_font(run, size=size, bold=bold, italic=italic)
    return p


def add_heading_custom(doc, text, size=16, bold=True, align="center", space_before=12, space_after=12):
    return add_para(doc, text, size=size, bold=bold, align=align,
                    space_before=space_before, space_after=space_after, line_spacing=1.15)


def add_body(doc, text, first_line=18):
    return add_para(doc, text, size=12, align="justify", space_after=8,
                    line_spacing=1.5, first_line=first_line)


def add_subsection(doc, text):
    return add_para(doc, text, size=14, bold=True, align="left",
                    space_before=14, space_after=8, line_spacing=1.15)


def add_subsubsection(doc, text):
    return add_para(doc, text, size=12, bold=True, italic=True, align="left",
                    space_before=10, space_after=6, line_spacing=1.15)


def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style="List Bullet")
    p.clear()
    p.paragraph_format.left_indent = Inches(0.25 + level * 0.25)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    set_run_font(run, size=12)
    return p


def page_break(doc):
    doc.add_page_break()


def set_page_margins(section, left=1.0, right=1.0, top=1.0, bottom=1.0):
    section.page_width = Inches(8.27)   # A4
    section.page_height = Inches(11.69)
    section.left_margin = Inches(left)
    section.right_margin = Inches(right)
    section.top_margin = Inches(top)
    section.bottom_margin = Inches(bottom)


def add_header_footer(section, page_num_start=None):
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = hp.add_run(PROJECT_SHORT)
    set_run_font(run, size=10, italic=True)

    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = fp.add_run("School of Computer Science and Engineering, KIIT, BBSR")
    set_run_font(run, size=9)


def set_cell_shading(cell, fill="1F4E79"):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def style_table_header(table):
    for cell in table.rows[0].cells:
        set_cell_shading(cell, "1F4E79")
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                set_run_font(run, size=11, bold=True, color=RGBColor(255, 255, 255))


def make_toc_table(doc):
    rows = [
        ("1", "Introduction", "1"),
        ("2", "Basic Concepts / Literature Review", "3"),
        ("", "2.1  Social Media Analytics", "3"),
        ("", "2.2  Sentiment Analysis and VADER", "3"),
        ("", "2.3  Keyword-based Topic Categorization", "4"),
        ("", "2.4  FastAPI and RESTful Backend Services", "4"),
        ("", "2.5  React and Modern Frontend Visualization", "5"),
        ("", "2.6  Related Work and Literature Review", "5"),
        ("3", "Problem Statement / Requirement Specifications", "7"),
        ("", "3.1  Project Planning", "8"),
        ("", "3.2  Project Analysis (SRS)", "9"),
        ("", "3.3  System Design", "11"),
        ("", "      3.3.1  Design Constraints", "11"),
        ("", "      3.3.2  System Architecture", "12"),
        ("4", "Implementation", "14"),
        ("", "4.1  Methodology / Proposal", "14"),
        ("", "4.2  Testing / Verification Plan", "16"),
        ("", "4.3  Result Analysis / Screenshots", "18"),
        ("", "4.4  Quality Assurance", "19"),
        ("5", "Standards Adopted", "20"),
        ("", "5.1  Design Standards", "20"),
        ("", "5.2  Coding Standards", "20"),
        ("", "5.3  Testing Standards", "21"),
        ("", "5.4  Cost Analysis", "21"),
        ("6", "Conclusion and Future Scope", "22"),
        ("", "6.1  Conclusion", "22"),
        ("", "6.2  Future Scope", "22"),
        ("", "References", "24"),
        ("", "Individual Contribution", "25"),
        ("", "Plagiarism Report", "29"),
    ]
    table = doc.add_table(rows=len(rows), cols=3)
    table.autofit = True
    for i, (num, title, page) in enumerate(rows):
        row = table.rows[i]
        row.cells[0].text = num
        row.cells[1].text = title
        row.cells[2].text = page
        for j, cell in enumerate(row.cells):
            for p in cell.paragraphs:
                for run in p.runs:
                    is_chapter = bool(num) or title in ("References", "Individual Contribution", "Plagiarism Report")
                    set_run_font(run, size=12, bold=is_chapter)
                if j == 2:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    return table


def cover_page(doc, is_inner=False):
    add_para(doc, "", space_after=0)
    add_heading_custom(doc, "A PROJECT REPORT", size=18, bold=True)
    add_heading_custom(doc, "on", size=12, bold=False, space_before=6, space_after=6)
    add_heading_custom(doc, f'"{PROJECT_NAME}"', size=16, bold=True, space_before=6, space_after=18)
    add_heading_custom(doc, "Submitted to", size=12, bold=False, space_before=6, space_after=6)
    add_heading_custom(doc, "KIIT Deemed to be University", size=16, bold=True, space_before=6, space_after=12)
    add_heading_custom(doc, "In Partial Fulfilment of the Requirement for the Award of",
                       size=12, bold=False, space_before=6, space_after=6)
    if is_inner:
        add_heading_custom(doc, "BACHELOR'S DEGREE IN COMPUTER SCIENCE & ENGINEERING / INFORMATION TECHNOLOGY",
                           size=13, bold=True, space_before=6, space_after=18)
    else:
        add_heading_custom(doc, "BACHELOR'S DEGREE IN", size=13, bold=True, space_before=6, space_after=4)
        add_heading_custom(doc, "COMPUTER SCIENCE & ENGINEERING", size=13, bold=True, space_before=2, space_after=2)
        add_heading_custom(doc, "OR", size=12, bold=True, space_before=2, space_after=2)
        add_heading_custom(doc, "INFORMATION TECHNOLOGY", size=13, bold=True, space_before=2, space_after=18)

    add_heading_custom(doc, "BY", size=12, bold=True, space_before=6, space_after=10)
    for name, _ in MEMBERS:
        add_para(doc, name, size=12, bold=True, align="center", space_after=2, line_spacing=1.15)
    add_para(doc, "", space_after=4)
    add_para(doc, "   ".join([r for _, r in MEMBERS]), size=11, bold=True,
             align="center", space_after=14, line_spacing=1.15)

    add_heading_custom(doc, "UNDER THE GUIDANCE OF", size=12, bold=True, space_before=10, space_after=8)
    add_heading_custom(doc, GUIDE_NAME, size=13, bold=True, space_before=4, space_after=24)

    add_heading_custom(doc, "SCHOOL OF COMPUTER SCIENCE AND ENGINEERING", size=12, bold=True, space_before=20, space_after=4)
    add_heading_custom(doc, "KALINGA INSTITUTE OF INDUSTRIAL TECHNOLOGY", size=12, bold=True, space_before=2, space_after=4)
    add_heading_custom(doc, "BHUBANESWAR, ODISHA - 751024", size=12, bold=True, space_before=2, space_after=4)
    add_heading_custom(doc, MONTH_YEAR, size=12, bold=True, space_before=8, space_after=4)


def build_report():
    doc = Document()
    section = doc.sections[0]
    set_page_margins(section)

    # ========== OUTER COVER ==========
    cover_page(doc, is_inner=False)

    # ========== INNER COVER ==========
    page_break(doc)
    cover_page(doc, is_inner=True)

    # ========== CERTIFICATE ==========
    page_break(doc)
    add_heading_custom(doc, "KIIT Deemed to be University", size=14, bold=True)
    add_heading_custom(doc, "School of Computer Science and Engineering", size=12, bold=True, space_before=2, space_after=2)
    add_heading_custom(doc, "Bhubaneswar, ODISHA 751024", size=12, bold=False, space_before=2, space_after=30)
    add_heading_custom(doc, "CERTIFICATE", size=18, bold=True, space_before=20, space_after=20)

    add_para(doc, "This is to certify that the project entitled", size=12, align="center",
             space_after=10, line_spacing=1.5)
    add_heading_custom(doc, f'"{PROJECT_NAME}"', size=14, bold=True, space_before=4, space_after=10)
    add_para(doc, "submitted by", size=12, align="center", space_after=10, line_spacing=1.5)
    for name, _ in MEMBERS:
        add_para(doc, name, size=12, bold=True, align="center", space_after=2, line_spacing=1.15)
    add_para(doc, "", space_after=4)
    add_para(doc, "   ".join([r for _, r in MEMBERS]), size=11, bold=True,
             align="center", space_after=14, line_spacing=1.15)

    cert_text = (
        f"is a record of bonafide work carried out by them, in the partial fulfilment of the "
        f"requirement for the award of Degree of Bachelor of Engineering (Computer Science & "
        f"Engineering / Information Technology) at KIIT Deemed to be University, Bhubaneswar. "
        f"This work is done during year {ACADEMIC_YEAR}, under our guidance."
    )
    add_para(doc, cert_text, size=12, align="justify", space_after=20, line_spacing=1.5, first_line=18)
    add_para(doc, "Date:\t/\t/", size=12, align="left", space_after=40, line_spacing=1.5)
    add_para(doc, "", space_after=40)
    add_para(doc, f"\t({GUIDE_NAME})", size=12, bold=True, align="right", space_after=2)
    add_para(doc, "\tProject Guide", size=12, align="right", space_after=2)

    # ========== ACKNOWLEDGEMENTS ==========
    page_break(doc)
    add_heading_custom(doc, "Acknowledgements", size=16, bold=True, space_after=16)

    ack = (
        f"We are profoundly grateful to {GUIDE_NAME} of the School of Computer Science and "
        f"Engineering, KIIT Deemed to be University, for expert guidance and continuous "
        f"encouragement throughout this project, from commencement to completion. The insightful "
        f"suggestions on system design, sentiment analysis methodology, and report preparation "
        f"were invaluable in shaping TwitterSense-AI into a complete and usable platform."
    )
    add_body(doc, ack)

    ack2 = (
        "We thank the School of Computer Science and Engineering for providing the academic "
        "environment and laboratory facilities that enabled us to experiment with modern web "
        "technologies, natural language processing libraries, and third-party social media APIs. "
        "We also acknowledge the open-source communities behind FastAPI, React, Vite, Recharts, "
        "and VADER Sentiment, whose tools made rapid prototyping and visualization possible."
    )
    add_body(doc, ack2)

    ack3 = (
        "Finally, we express our gratitude to our families and peers for their patience and "
        "support during the development and documentation of this project."
    )
    add_body(doc, ack3)

    add_para(doc, "", space_after=20)
    for name, _ in MEMBERS:
        add_para(doc, f"Signature of  {name}", size=12, align="right",
                 space_after=14, line_spacing=1.5)

    # ========== ABSTRACT ==========
    page_break(doc)
    add_heading_custom(doc, "ABSTRACT", size=16, bold=True, space_after=16)

    abs1 = (
        "Social media platforms such as X (formerly Twitter) generate enormous volumes of "
        "public conversation every day. Brands, researchers, journalists, and decision-makers "
        "need timely insight into what people discuss and how they feel about topics and "
        "public figures. Manually reading hundreds of posts is impractical, and existing "
        "analytics tools are often expensive, closed, or limited in combining profile-level "
        "sentiment with topic-centric account discovery."
    )
    add_body(doc, abs1)

    abs2 = (
        "TwitterSense-AI is a full-stack social media intelligence platform that addresses "
        "this gap. The system fetches public posts through TweetAPI, performs lexicon-based "
        "sentiment analysis using VADER, and classifies posts into multiple topical categories "
        "using a curated keyword dictionary. A FastAPI backend exposes REST endpoints for "
        "profile analysis and topic search, while a React (Vite) frontend presents interactive "
        "dashboards with sentiment bar charts, category doughnut charts, ranked account tables, "
        "and annotated post lists. Users can analyse up to 100 recent posts of any public "
        "profile, or discover and rank the top accounts discussing a chosen topic."
    )
    add_body(doc, abs2)

    abs3 = (
        "The project demonstrates end-to-end engineering of a practical NLP-backed web "
        "application: API integration with pagination and error handling, multi-label text "
        "categorization, sentiment scoring, CORS-secured client–server communication, and "
        "responsive data visualization. The resulting platform provides an accessible, "
        "low-cost starting point for social listening, academic study of online discourse, "
        "and further research into real-time opinion mining."
    )
    add_body(doc, abs3)

    add_para(doc, "", space_after=10)
    kw = add_para(doc, "", size=12, align="justify", space_after=8, line_spacing=1.5)
    run = kw.add_run("Keywords: ")
    set_run_font(run, size=12, bold=True)
    run = kw.add_run(
        "Social Media Analytics, Sentiment Analysis, VADER, Topic Classification, "
        "FastAPI, React, X/Twitter Intelligence, Data Visualization"
    )
    set_run_font(run, size=12)

    # ========== CONTENTS ==========
    page_break(doc)
    add_heading_custom(doc, "Contents", size=18, bold=True, space_after=16)
    make_toc_table(doc)

    # ========== LIST OF FIGURES ==========
    page_break(doc)
    add_heading_custom(doc, "List of Figures", size=18, bold=True, space_after=16)
    figures = [
        ("1.1", "High-level overview of TwitterSense-AI workflow", "2"),
        ("3.1", "System architecture / block diagram of TwitterSense-AI", "12"),
        ("3.2", "Data flow for profile analysis and topic intelligence", "13"),
        ("4.1", "Profile Analysis dashboard with sentiment and category charts", "18"),
        ("4.2", "Topic Intelligence ranked accounts and matching posts", "18"),
    ]
    fig_table = doc.add_table(rows=len(figures), cols=3)
    for i, (num, cap, page) in enumerate(figures):
        fig_table.rows[i].cells[0].text = num
        fig_table.rows[i].cells[1].text = cap
        fig_table.rows[i].cells[2].text = page
        for j, cell in enumerate(fig_table.rows[i].cells):
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, size=12)
                if j == 2:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # ========== LIST OF TABLES ==========
    page_break(doc)
    add_heading_custom(doc, "List of Tables", size=18, bold=True, space_after=16)
    tables_list = [
        ("3.1", "Functional requirements of TwitterSense-AI", "9"),
        ("3.2", "Non-functional requirements", "10"),
        ("3.3", "Hardware and software environment", "11"),
        ("4.1", "API endpoints implemented in the backend", "15"),
        ("4.2", "Test cases and verification results", "16"),
        ("4.3", "Sample sentiment and category distribution (illustrative)", "19"),
        ("5.1", "Indicative cost comparison with commercial tools", "21"),
    ]
    lt = doc.add_table(rows=len(tables_list), cols=3)
    for i, (num, cap, page) in enumerate(tables_list):
        lt.rows[i].cells[0].text = num
        lt.rows[i].cells[1].text = cap
        lt.rows[i].cells[2].text = page
        for j, cell in enumerate(lt.rows[i].cells):
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, size=12)
                if j == 2:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # =====================================================================
    # CHAPTER 1 — INTRODUCTION
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, PROJECT_SHORT, size=12, bold=True, space_after=20)
    add_heading_custom(doc, "Chapter 1", size=16, bold=True, space_after=4)
    add_heading_custom(doc, "Introduction", size=16, bold=True, space_before=4, space_after=16)

    add_body(doc, (
        "Public conversation on X (formerly Twitter) has become a primary signal of collective "
        "opinion. Elections, product launches, sports events, and crises all leave dense textual "
        "trails within minutes. For organisations and researchers, the ability to sense sentiment "
        "and topical themes across these streams is no longer optional; it is a competitive and "
        "scholarly necessity. Yet raw timelines are noisy, multilingual in tone, and too large "
        "to inspect by hand."
    ))

    add_body(doc, (
        "Commercial social listening suites exist, but they typically demand paid enterprise "
        "subscriptions, restrict export flexibility, or hide the analysis pipeline behind opaque "
        "dashboards. Academic and student projects often stop at offline CSV analysis and never "
        "expose a usable interactive product. There remains a clear gap for an open, transparent, "
        "and affordable web application that (i) ingests live public posts for a given user, "
        "(ii) scores sentiment and multi-label categories in near real time, (iii) discovers "
        "influential accounts for an arbitrary topic, and (iv) visualises results without "
        "requiring the user to write code."
    ))

    add_body(doc, (
        "TwitterSense-AI was conceived to fill this gap. The platform pairs a Python FastAPI "
        "service with a React single-page application. Posts are retrieved via TweetAPI, "
        "analysed with the VADER sentiment lexicon, and tagged against a domain keyword map "
        "covering technology, finance, politics, sports, entertainment, health, environment, "
        "education, fashion, retail, food, and travel. The frontend surfaces profile analytics "
        "and topic intelligence through charts, ranked tables, and annotated post feeds."
    ))

    add_body(doc, (
        "The remainder of this report is organised as follows. Chapter 2 presents the basic "
        "concepts and literature that underpin the work. Chapter 3 states the problem, plans "
        "the project, analyses requirements in SRS form, and describes system design. Chapter 4 "
        "covers methodology, testing, results, and quality assurance. Chapter 5 records the "
        "design, coding, and testing standards adopted, along with a brief cost analysis. "
        "Chapter 6 concludes and outlines future scope. References, individual contribution "
        "reports, and the plagiarism certificate follow the main chapters."
    ))

    add_para(doc, "", space_after=8)
    add_para(doc, "Figure 1.1: High-level overview of TwitterSense-AI workflow",
             size=11, italic=True, align="center", space_before=12, space_after=8)
    add_para(doc, (
        "[Insert diagram: User → React Frontend → FastAPI Backend → TweetAPI / VADER / "
        "Category Engine → Charts & Ranked Accounts]"
    ), size=11, italic=True, align="center", space_after=12)

    # =====================================================================
    # CHAPTER 2 — BASIC CONCEPTS / LITERATURE REVIEW
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, PROJECT_SHORT, size=12, bold=True, space_after=20)
    add_heading_custom(doc, "Chapter 2", size=16, bold=True, space_after=4)
    add_heading_custom(doc, "Basic Concepts / Literature Review", size=16, bold=True, space_before=4, space_after=12)

    add_body(doc, (
        "This chapter introduces the foundational ideas, tools, and prior research that a "
        "reader needs in order to understand TwitterSense-AI. Subsections progress from "
        "social media analytics and sentiment analysis to the specific backend and frontend "
        "technologies used, and close with a concise literature review."
    ))

    add_subsection(doc, "2.1 Social Media Analytics")
    add_body(doc, (
        "Social media analytics is the practice of collecting, processing, and interpreting "
        "user-generated content to extract actionable insight. Typical tasks include volume "
        "tracking, influencer identification, topic detection, and sentiment monitoring. "
        "On X, posts are short, hashtag-rich, and highly temporal, which favours lightweight "
        "pipelines that can refresh results as new pages of posts arrive through cursor-based "
        "pagination."
    ))

    add_subsection(doc, "2.2 Sentiment Analysis and VADER")
    add_body(doc, (
        "Sentiment analysis assigns affective polarity—positive, negative, or neutral—to text. "
        "Approaches range from supervised machine learning to lexicon-based methods. VADER "
        "(Valence Aware Dictionary and sEntiment Reasoner) is a rule-augmented lexicon "
        "designed for social media. It returns a compound score typically in [−1, 1]. In "
        "TwitterSense-AI, compound ≥ 0.05 is labelled Positive, ≤ −0.05 Negative, and "
        "otherwise Neutral. VADER was chosen for its speed, zero training-data requirement, "
        "and proven suitability for informal short text."
    ))

    add_subsection(doc, "2.3 Keyword-based Topic Categorization")
    add_body(doc, (
        "Beyond polarity, stakeholders care about what a conversation is about. The system "
        "maintains a dictionary of categories (Technology & AI, Business & Finance, Politics, "
        "Sports, Entertainment, Science & Space, Health, Environment, Education, Fashion & "
        "Lifestyle, E-commerce & Retail, Food & Beverage, Travel & Tourism), each with a "
        "list of indicative keywords. A post may match multiple categories (multi-label). "
        "Matches are ranked by keyword hit count so that the strongest themes surface first "
        "in aggregate distributions."
    ))

    add_subsection(doc, "2.4 FastAPI and RESTful Backend Services")
    add_body(doc, (
        "FastAPI is a modern Python web framework that generates OpenAPI documentation "
        "automatically and uses type hints for validation. TwitterSense-AI exposes endpoints "
        "for health checks, profile lookup, tweet retrieval with pagination, profile-level "
        "analysis, and topic search. CORS middleware restricts browser origins to the local "
        "Vite development server. Secrets such as the TweetAPI key are loaded from environment "
        "variables via python-dotenv, keeping credentials out of source control."
    ))

    add_subsection(doc, "2.5 React and Modern Frontend Visualization")
    add_body(doc, (
        "The user interface is built with React 19 and Vite. Axios performs HTTP calls to the "
        "backend. Recharts renders horizontal sentiment bar charts and doughnut charts for "
        "category distribution. Lucide icons and a custom CSS layout provide a sidebar "
        "workspace with two primary modes: Profile Analysis and Topic Intelligence. This "
        "separation mirrors the two core use cases of the product."
    ))

    add_subsection(doc, "2.6 Related Work and Literature Review")
    add_body(doc, (
        "Hutto and Gilbert introduced VADER as a gold-standard lexicon for social media "
        "sentiment, showing strong correlation with human judgements on Twitter-style text. "
        "Subsequent surveys of opinion mining emphasise hybrid systems that combine lexicon "
        "scores with topic models or classifiers. Commercial platforms (Brandwatch, Sprout "
        "Social, Hootsuite) illustrate the market demand for dashboards that fuse sentiment "
        "with influencer ranking, but their closed nature limits pedagogical transparency."
    ))
    add_body(doc, (
        "Academic prototypes frequently use Tweepy against the official X API; however, "
        "access tiers and rate limits have become restrictive for student projects. "
        "Third-party aggregators such as TweetAPI offer a practical alternative for "
        "read-only public data. TwitterSense-AI builds on these threads by packaging "
        "retrieval, VADER scoring, keyword multi-labelling, and interactive visualisation "
        "into one reproducible full-stack application tailored for undergraduate project "
        "requirements."
    ))

    # =====================================================================
    # CHAPTER 3 — PROBLEM STATEMENT / REQUIREMENT SPECIFICATIONS
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, PROJECT_SHORT, size=12, bold=True, space_after=20)
    add_heading_custom(doc, "Chapter 3", size=16, bold=True, space_after=4)
    add_heading_custom(doc, "Problem Statement / Requirement Specifications",
                       size=16, bold=True, space_before=4, space_after=12)

    add_subsection(doc, "Problem Statement")
    add_body(doc, (
        "Public discourse on X is abundant but unstructured. Stakeholders who need to "
        "understand the emotional tone and topical focus of a public profile, or who wish "
        "to identify which accounts dominate conversation around a keyword, currently face "
        "a fragmented toolchain: manual scrolling, spreadsheet exports, or costly proprietary "
        "dashboards. There is no lightweight, transparent, student-accessible web system that "
        "unifies (1) automated retrieval of recent public posts, (2) explainable sentiment "
        "scoring, (3) multi-label thematic categorisation, (4) topic-based account ranking, "
        "and (5) interactive visual analytics in a single product."
    ))
    add_body(doc, (
        "Therefore, the problem addressed by this project is: Design and implement an "
        "end-to-end social media intelligence platform—TwitterSense-AI—that analyses public "
        "X profiles and topics in near real time, produces sentiment and category insights, "
        "ranks relevant accounts for a given subject, and presents results through an "
        "intuitive web dashboard."
    ))

    add_subsection(doc, "3.1 Project Planning")
    add_body(doc, (
        "Project execution followed an iterative full-stack plan. The major planning steps "
        "and feature milestones were:"
    ))
    add_bullet(doc, "Requirement gathering: define Profile Analysis and Topic Intelligence use cases.")
    add_bullet(doc, "Technology selection: FastAPI, React+Vite, VADER, TweetAPI, Recharts, Axios.")
    add_bullet(doc, "Backend spike: authenticate to TweetAPI; fetch user-by-username and paginated tweets.")
    add_bullet(doc, "NLP layer: integrate VADER; design CATEGORY_KEYWORDS multi-label classifier.")
    add_bullet(doc, "Analysis API: /analyze/{username} aggregating sentiment and category distributions.")
    add_bullet(doc, "Topic search API: /topic-search/{topic} with account aggregation and top-10 ranking.")
    add_bullet(doc, "Frontend scaffolding: sidebar navigation, forms, loading/error/empty states.")
    add_bullet(doc, "Visualization: sentiment bars, category doughnut, ranked account table, post cards.")
    add_bullet(doc, "Hardening: CORS, timeouts, HTTP error mapping, duplicate tweet de-duplication.")
    add_bullet(doc, "Documentation and report preparation as per KIIT project guidelines.")

    add_body(doc, (
        "User-facing features planned and delivered include: analyse up to 100 recent public "
        "posts per profile; display posts fetched/analysed counts; show positive/neutral/"
        "negative breakdown; list discovered categories; search a free-text topic across "
        "multiple result pages; rank accounts by matching post count then followers; and "
        "browse annotated matching tweets with optional links to originals."
    ))

    add_subsection(doc, "3.2 Project Analysis (SRS)")
    add_body(doc, (
        "After conceptualising the problem, requirements were analysed for clarity, "
        "feasibility, and testability. The Software Requirements Specification below follows "
        "the spirit of IEEE 830 recommended practice."
    ))

    add_subsubsection(doc, "3.2.1 Purpose and Scope")
    add_body(doc, (
        "Purpose: Provide an interactive web application for sentiment and topic intelligence "
        "over public X data. Scope: Read-only analysis of public profiles and search results; "
        "no posting, liking, or private-account access. Out of scope: multilingual transformer "
        "models, authenticated user login, mobile native apps, and historical archive mining "
        "beyond API page limits."
    ))

    add_subsubsection(doc, "3.2.2 Functional Requirements")
    add_para(doc, "Table 3.1: Functional requirements of TwitterSense-AI",
             size=11, italic=True, align="center", space_before=8, space_after=6)

    fr_data = [
        ("FR-01", "Accept X username or profile URL and normalise to a username."),
        ("FR-02", "Fetch profile metadata and up to 100 recent public posts with pagination."),
        ("FR-03", "Compute VADER sentiment for each extractable post text."),
        ("FR-04", "Assign zero or more topical categories via keyword matching."),
        ("FR-05", "Return aggregate sentiment and category distributions with analysed posts."),
        ("FR-06", "Accept a topic string and search latest public posts across configurable pages."),
        ("FR-07", "Aggregate unique accounts with post and sentiment counts; rank top 10."),
        ("FR-08", "Render interactive charts, stats cards, tables, and post lists in the UI."),
        ("FR-09", "Surface clear loading, empty, and error states for both workflows."),
        ("FR-10", "Expose health and root endpoints for backend availability checks."),
    ]
    t = doc.add_table(rows=1 + len(fr_data), cols=2)
    t.style = "Table Grid"
    t.rows[0].cells[0].text = "ID"
    t.rows[0].cells[1].text = "Requirement"
    style_table_header(t)
    for i, (rid, desc) in enumerate(fr_data, start=1):
        t.rows[i].cells[0].text = rid
        t.rows[i].cells[1].text = desc
        for cell in t.rows[i].cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, size=11)

    add_para(doc, "", space_after=8)
    add_subsubsection(doc, "3.2.3 Non-Functional Requirements")
    add_para(doc, "Table 3.2: Non-functional requirements",
             size=11, italic=True, align="center", space_before=8, space_after=6)
    nfr = [
        ("NFR-01", "Performance", "Profile analysis should complete within client timeout (120 s) for ≤100 posts."),
        ("NFR-02", "Usability", "Two-tab workspace with explicit labels; no CLI required for end users."),
        ("NFR-03", "Security", "API keys stored in .env; CORS limited to localhost:5173 in development."),
        ("NFR-04", "Reliability", "HTTP and connection failures mapped to meaningful HTTPException details."),
        ("NFR-05", "Maintainability", "Modular helpers for text extraction, classification, and author parsing."),
        ("NFR-06", "Portability", "Runs on Windows/Linux with Python 3.x and Node.js toolchain."),
    ]
    t2 = doc.add_table(rows=1 + len(nfr), cols=3)
    t2.style = "Table Grid"
    t2.rows[0].cells[0].text = "ID"
    t2.rows[0].cells[1].text = "Category"
    t2.rows[0].cells[2].text = "Description"
    style_table_header(t2)
    for i, row in enumerate(nfr, start=1):
        for j, val in enumerate(row):
            t2.rows[i].cells[j].text = val
            for p in t2.rows[i].cells[j].paragraphs:
                for run in p.runs:
                    set_run_font(run, size=10)

    add_para(doc, "", space_after=8)
    add_subsubsection(doc, "3.2.4 Ambiguities Resolved")
    add_body(doc, (
        "Analysis clarified several ambiguities: usernames may arrive with '@' or as full "
        "profile URLs—normalisation strips both. TweetAPI responses vary in text field "
        "names (text, fullText, full_text, legacy)—a dedicated extractor handles formats. "
        "Category fields in the UI may receive lists; ranking for topic search uses "
        "post_count primarily and followers secondarily. Empty post text is excluded from "
        "sentiment aggregates to avoid skewing distributions."
    ))

    add_subsection(doc, "3.3 System Design")
    add_subsubsection(doc, "3.3.1 Design Constraints")
    add_body(doc, (
        "The working environment and constraints under which the system was designed are "
        "summarised below."
    ))
    add_para(doc, "Table 3.3: Hardware and software environment",
             size=11, italic=True, align="center", space_before=8, space_after=6)
    env = [
        ("Hardware", "Standard PC/laptop; ≥8 GB RAM recommended; internet connectivity required"),
        ("OS", "Windows 10/11 (development); portable to Linux/macOS"),
        ("Backend", "Python 3.13, FastAPI, Uvicorn, requests, python-dotenv, vaderSentiment"),
        ("Frontend", "Node.js, React 19, Vite 8, Axios, Recharts, Lucide React"),
        ("External API", "TweetAPI (tw-v2) with X-API-Key authentication"),
        ("Browser", "Modern Chromium/Firefox/Edge for SPA rendering"),
        ("Constraint", "Public data only; API rate/page limits (max 10 tweet pages, 5 search pages)"),
        ("Constraint", "Lexicon sentiment—not deep contextual transformers—for latency and simplicity"),
    ]
    t3 = doc.add_table(rows=1 + len(env), cols=2)
    t3.style = "Table Grid"
    t3.rows[0].cells[0].text = "Item"
    t3.rows[0].cells[1].text = "Specification"
    style_table_header(t3)
    for i, (a, b) in enumerate(env, start=1):
        t3.rows[i].cells[0].text = a
        t3.rows[i].cells[1].text = b
        for cell in t3.rows[i].cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, size=10)

    add_para(doc, "", space_after=8)
    add_subsubsection(doc, "3.3.2 System Architecture / Block Diagram")
    add_body(doc, (
        "TwitterSense-AI follows a classic three-tier web architecture: Presentation "
        "(React SPA), Application (FastAPI services and NLP helpers), and External Data "
        "(TweetAPI). The presentation tier never holds the API key; all privileged calls "
        "are proxied through the backend."
    ))

    add_para(doc, "Figure 3.1: System architecture / block diagram of TwitterSense-AI",
             size=11, italic=True, align="center", space_before=10, space_after=6)
    add_para(doc, (
        "[Insert block diagram]\n"
        "Browser (React/Vite)  ↔  FastAPI (CORS, routes, VADER, categories)  ↔  TweetAPI\n"
        "                              ↓\n"
        "                     JSON: distributions, posts, rankings\n"
        "                              ↓\n"
        "                     Recharts + Tables + Post cards"
    ), size=11, italic=True, align="center", space_after=10)

    add_body(doc, (
        "Profile analysis flow: user submits username → frontend calls GET /analyze/{username} "
        "→ backend resolves user id → paginates /user/tweets → analyse_post() for each tweet → "
        "Counter aggregates → JSON response → charts update."
    ))
    add_body(doc, (
        "Topic intelligence flow: user submits topic → GET /topic-search/{topic}?pages=n → "
        "backend loops search pages → extract_author + analyse_post → account dictionary "
        "updated → sorted top 10 → matching_posts returned → table and feed rendered."
    ))

    add_para(doc, "Figure 3.2: Data flow for profile analysis and topic intelligence",
             size=11, italic=True, align="center", space_before=10, space_after=6)
    add_para(doc, "[Insert sequence / data-flow diagram for the two workflows]",
             size=11, italic=True, align="center", space_after=10)

    # =====================================================================
    # CHAPTER 4 — IMPLEMENTATION
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, PROJECT_SHORT, size=12, bold=True, space_after=20)
    add_heading_custom(doc, "Chapter 4", size=16, bold=True, space_after=4)
    add_heading_custom(doc, "Implementation", size=16, bold=True, space_before=4, space_after=12)

    add_body(doc, (
        "This chapter describes how the planned design was realised in code, how correctness "
        "was verified, what outputs the system produces, and what quality practices were "
        "followed during development."
    ))

    add_subsection(doc, "4.1 Methodology / Proposal")
    add_body(doc, (
        "Implementation proceeded module-wise. The backend file main.py centralises routing "
        "and NLP helpers. The frontend App.jsx owns state for both tabs and orchestrates "
        "Axios requests. Styling lives in App.css / index.css."
    ))

    add_subsubsection(doc, "4.1.1 Backend Methodology")
    add_bullet(doc, "Load TWEETAPI_API_KEY from environment; fail with HTTP 500 if missing.")
    add_bullet(doc, "tweetapi_get(endpoint, params): shared GET wrapper with timeout and error mapping.")
    add_bullet(doc, "get_profile / get_profile_tweets: resolve username → id → cursor pagination with de-duplication.")
    add_bullet(doc, "get_tweet_text / extract_author: defensive parsers for heterogeneous API payloads.")
    add_bullet(doc, "classify_category: lowercase keyword scan; return categories sorted by match_count.")
    add_bullet(doc, "analyze_post: VADER polarity_scores → label + rounded compound score.")
    add_bullet(doc, "analyze_profile: map tweets → valid posts → Counter for sentiment and categories.")
    add_bullet(doc, "topic_search: multi-page Latest search; accumulate accounts; rank by (posts, followers).")

    add_para(doc, "Table 4.1: API endpoints implemented in the backend",
             size=11, italic=True, align="center", space_before=10, space_after=6)
    apis = [
        ("GET /", "Service banner message"),
        ("GET /health", "Liveness probe {status: ok}"),
        ("GET /profile/{username}", "Raw profile payload from TweetAPI"),
        ("GET /profile/{username}/tweets", "Paginated tweets (limit 1–100)"),
        ("GET /analyze/{username}", "Sentiment + category analysis package"),
        ("GET /topic-search/{topic}", "Topic search with top accounts and posts"),
    ]
    ta = doc.add_table(rows=1 + len(apis), cols=2)
    ta.style = "Table Grid"
    ta.rows[0].cells[0].text = "Endpoint"
    ta.rows[0].cells[1].text = "Purpose"
    style_table_header(ta)
    for i, (a, b) in enumerate(apis, start=1):
        ta.rows[i].cells[0].text = a
        ta.rows[i].cells[1].text = b
        for cell in ta.rows[i].cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, size=10)

    add_para(doc, "", space_after=6)
    add_subsubsection(doc, "4.1.2 Frontend Methodology")
    add_body(doc, (
        "The SPA keeps activeTab state ('profile' | 'topic'). Profile form validates input "
        "via getUsername(), then calls /analyze with limit=100 and a 120-second timeout. "
        "Topic form calls /topic-search with pages=3. Derived arrays sentimentData and "
        "categoryData feed Recharts. Topic results render a ranked HTML table and a list "
        "of recent matching tweets with sentiment tags and optional original URLs."
    ))

    add_subsection(doc, "4.2 Testing / Verification Plan")
    add_body(doc, (
        "Verification combined manual API checks (via browser/OpenAPI) and UI walkthroughs. "
        "Representative test cases are recorded in Table 4.2."
    ))
    add_para(doc, "Table 4.2: Test cases and verification results",
             size=11, italic=True, align="center", space_before=8, space_after=6)

    tests = [
        ("T01", "Health check", "GET /health", "Returns JSON status", "Pass — {status: ok}"),
        ("T02", "Invalid empty topic", "topic = blank", "HTTP 400 / UI error", "Pass — validation message"),
        ("T03", "Username normalisation", "Input @user or URL", "Strips @ and path", "Pass"),
        ("T04", "Profile analyze happy path", "Valid public username", "Charts + posts render", "Pass"),
        ("T05", "Missing API key", "Unset env var", "HTTP 500 detail", "Pass — clear error"),
        ("T06", "Sentiment thresholds", "Compound ±0.05 boundary", "Correct label", "Pass"),
        ("T07", "Multi-label category", "Text with AI + finance terms", "Both categories counted", "Pass"),
        ("T08", "Topic ranking", "Popular keyword, pages=3", "≤10 accounts, sorted", "Pass"),
        ("T09", "CORS allowlist", "Origin localhost:5173", "Browser request succeeds", "Pass"),
        ("T10", "Empty / private edge", "No extractable text", "Excluded from aggregates", "Pass"),
    ]
    tt = doc.add_table(rows=1 + len(tests), cols=5)
    tt.style = "Table Grid"
    headers = ["Test ID", "Test Case Title", "Test Condition", "Expected Result", "Observed"]
    for j, h in enumerate(headers):
        tt.rows[0].cells[j].text = h
    style_table_header(tt)
    for i, row in enumerate(tests, start=1):
        for j, val in enumerate(row):
            tt.rows[i].cells[j].text = val
            for p in tt.rows[i].cells[j].paragraphs:
                for run in p.runs:
                    set_run_font(run, size=9)

    add_subsection(doc, "4.3 Result Analysis / Screenshots")
    add_body(doc, (
        "Successful runs of the Profile Analysis tab display four statistic cards (posts "
        "fetched, analysed, positive count, categories found), a horizontal bar chart of "
        "sentiment, a doughnut chart of categories with legend, and a scrollable list of "
        "analysed posts tagged with sentiment and score. The Topic Intelligence tab displays "
        "retrieval stats, a ranked top-10 account table (followers, posts, positive/neutral/"
        "negative), and a feed of matching tweets."
    ))

    add_para(doc, "Figure 4.1: Profile Analysis dashboard with sentiment and category charts",
             size=11, italic=True, align="center", space_before=10, space_after=4)
    add_para(doc, "[Insert screenshot of Profile Analysis after a successful run]",
             size=11, italic=True, align="center", space_after=10)

    add_para(doc, "Figure 4.2: Topic Intelligence ranked accounts and matching posts",
             size=11, italic=True, align="center", space_before=10, space_after=4)
    add_para(doc, "[Insert screenshot of Topic Intelligence results]",
             size=11, italic=True, align="center", space_after=10)

    add_para(doc, "Table 4.3: Sample sentiment and category distribution (illustrative)",
             size=11, italic=True, align="center", space_before=8, space_after=6)
    sample = [
        ("Metric", "Example Value"),
        ("Posts fetched", "100"),
        ("Posts analysed", "96"),
        ("Positive / Neutral / Negative", "42 / 31 / 23"),
        ("Top categories", "Technology & AI, Business & Finance, Education"),
        ("Topic unique accounts", "57 (top 10 ranked for display)"),
    ]
    ts = doc.add_table(rows=len(sample), cols=2)
    ts.style = "Table Grid"
    for i, (a, b) in enumerate(sample):
        ts.rows[i].cells[0].text = a
        ts.rows[i].cells[1].text = b
        for cell in ts.rows[i].cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, size=10, bold=(i == 0))
    style_table_header(ts)

    add_body(doc, (
        "Result interpretation: a high positive share suggests generally favourable tone in "
        "the sampled window; category spikes reveal thematic focus useful for branding or "
        "research framing. Topic rankings highlight accounts that actively post about the "
        "query—useful for outreach or further network analysis. Because sampling is limited "
        "to recent pages, results describe the current conversation slice rather than "
        "lifetime behaviour."
    ))

    add_subsection(doc, "4.4 Quality Assurance")
    add_body(doc, (
        "Quality was assured through code reviews within the team, consistent naming, "
        "defensive parsing of third-party JSON, explicit HTTP status mapping, and UI "
        "states that prevent silent failure. Environment-based secrets avoid committing "
        "credentials. Frontend linting via Oxlint and backend type-hinted FastAPI routes "
        "reduce common defects. Prior to demonstration, end-to-end smoke tests (Table 4.2) "
        "were re-run. A Turnitin plagiarism check of this report is mandated (target < 25%) "
        "and attached in the final section."
    ))

    # =====================================================================
    # CHAPTER 5 — STANDARDS ADOPTED
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, PROJECT_SHORT, size=12, bold=True, space_after=20)
    add_heading_custom(doc, "Chapter 5", size=16, bold=True, space_after=4)
    add_heading_custom(doc, "Standards Adopted", size=16, bold=True, space_before=4, space_after=12)

    add_subsection(doc, "5.1 Design Standards")
    add_body(doc, (
        "System design followed layered architecture and REST resource conventions aligned "
        "with common industry practice and the spirit of IEEE software design documentation. "
        "Requirements were organised similarly to IEEE 830 SRS structure (purpose, scope, "
        "functional/non-functional requirements). UI information architecture separates "
        "concerns into two workflows, reducing cognitive load. Block diagrams and data-flow "
        "descriptions document component boundaries for maintainability."
    ))

    add_subsection(doc, "5.2 Coding Standards")
    add_body(doc, (
        "The following coding practices were applied throughout the repository:"
    ))
    add_bullet(doc, "Prefer clear, descriptive function and variable names (analyze_profile, get_tweet_text).")
    add_bullet(doc, "Keep helpers focused on a single responsibility; avoid unnecessarily long functions.")
    add_bullet(doc, "Use consistent indentation and logical paragraphing of code blocks.")
    add_bullet(doc, "Document non-obvious behaviour with concise comments (pagination, payload variants).")
    add_bullet(doc, "Do not hard-code secrets; load configuration from environment variables.")
    add_bullet(doc, "Handle errors explicitly and return actionable messages to the client.")
    add_bullet(doc, "On the frontend, isolate presentational components (e.g., StatCard) for reuse.")
    add_bullet(doc, "Avoid committing build artefacts and virtual environments; use .gitignore.")

    add_subsection(doc, "5.3 Testing Standards")
    add_body(doc, (
        "Testing followed a pragmatic subset of IEEE 829 / ISO/IEC/IEEE 29119 ideas: each "
        "test case records an identifier, title, condition, expected behaviour, and observed "
        "result (Table 4.2). Verification covered unit-level helpers (thresholds, "
        "normalisation), API contract checks, and system-level UI scenarios. Defects found "
        "during integration (response shape differences, category field naming) were fixed "
        "before final demonstration. Continuous improvement is expected as automated unit "
        "tests are added in future iterations."
    ))

    add_subsection(doc, "5.4 Cost Analysis")
    add_body(doc, (
        "TwitterSense-AI is intentionally low-cost for academic use. Core software stacks "
        "(Python, FastAPI, React, Vite, VADER, Recharts) are free and open source. The "
        "principal recurring cost is the TweetAPI subscription tier required to fetch public "
        "posts; local compute uses existing student hardware. Compared with enterprise social "
        "listening platforms that charge monthly seats in the hundreds of dollars, the "
        "proposed solution offers a dramatically lower entry barrier while remaining "
        "transparent and extensible."
    ))
    add_para(doc, "Table 5.1: Indicative cost comparison with commercial tools",
             size=11, italic=True, align="center", space_before=8, space_after=6)
    cost = [
        ("Component", "TwitterSense-AI", "Typical Commercial Suite"),
        ("Application software", "Open source (₹0)", "Licensed SaaS"),
        ("Data access", "TweetAPI student/dev plan", "Bundled / premium firehose"),
        ("Hosting (dev)", "Localhost", "Cloud multi-tenant"),
        ("Approx. monthly (student)", "Low (API plan only)", "High (per-seat fees)"),
    ]
    tc = doc.add_table(rows=len(cost), cols=3)
    tc.style = "Table Grid"
    for i, row in enumerate(cost):
        for j, val in enumerate(row):
            tc.rows[i].cells[j].text = val
            for p in tc.rows[i].cells[j].paragraphs:
                for run in p.runs:
                    set_run_font(run, size=10, bold=(i == 0))
    style_table_header(tc)

    # =====================================================================
    # CHAPTER 6 — CONCLUSION AND FUTURE SCOPE
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, PROJECT_SHORT, size=12, bold=True, space_after=20)
    add_heading_custom(doc, "Chapter 6", size=16, bold=True, space_after=4)
    add_heading_custom(doc, "Conclusion and Future Scope", size=16, bold=True, space_before=4, space_after=12)

    add_subsection(doc, "6.1 Conclusion")
    add_body(doc, (
        "TwitterSense-AI successfully delivers a full-stack solution to the stated problem: "
        "making public X discourse analysable through explainable sentiment scoring, "
        "multi-label categorisation, topic-centric account discovery, and interactive "
        "visualisation. The FastAPI backend integrates TweetAPI retrieval with VADER and a "
        "keyword engine; the React frontend turns analytical JSON into actionable dashboards. "
        "Requirements captured in Chapter 3 were implemented and verified through structured "
        "test cases. The project demonstrates practical competence in web engineering, NLP "
        "fundamentals, API design, and data presentation—meeting the educational and "
        "functional goals of an undergraduate major project."
    ))

    add_subsection(doc, "6.2 Future Scope")
    add_body(doc, (
        "Several extensions can further strengthen the platform:"
    ))
    add_bullet(doc, "Replace or augment VADER with transformer models (e.g., RoBERTa) for contextual sentiment.")
    add_bullet(doc, "Add user authentication, saved analyses, and export to CSV/PDF.")
    add_bullet(doc, "Support multilingual posts and code-mixed Indian language text.")
    add_bullet(doc, "Stream near real-time updates via WebSockets for live events.")
    add_bullet(doc, "Enrich influencer ranking with engagement metrics and network graphs.")
    add_bullet(doc, "Deploy backend and frontend to cloud with CI/CD and automated test suites.")
    add_bullet(doc, "Improve category models using supervised multi-label classifiers on labelled data.")
    add_bullet(doc, "Add comparative mode to analyse two profiles or topics side by side.")

    # =====================================================================
    # REFERENCES
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, PROJECT_SHORT, size=12, bold=True, space_after=20)
    add_heading_custom(doc, "References", size=16, bold=True, space_after=16)

    refs = [
        'C. J. Hutto and E. Gilbert, "VADER: A Parsimonious Rule-based Model for Sentiment Analysis of Social Media Text," in Proc. AAAI ICWSM, 2014.',
        'B. Liu, Sentiment Analysis and Opinion Mining. Morgan & Claypool, 2012.',
        'S. Bird, E. Klein, and E. Loper, Natural Language Processing with Python. O\'Reilly Media, 2009.',
        'FastAPI Documentation, "FastAPI — Modern Python Web Framework," https://fastapi.tiangolo.com/ (accessed Nov. 2026).',
        'React Documentation, "React — A JavaScript library for web user interfaces," https://react.dev/ (accessed Nov. 2026).',
        'Vite Documentation, "Vite Next Generation Frontend Tooling," https://vitejs.dev/ (accessed Nov. 2026).',
        'Recharts, "Recharts — A composable charting library built on React components," https://recharts.org/ (accessed Nov. 2026).',
        'IEEE Computer Society, "IEEE Std 830-1998 — IEEE Recommended Practice for Software Requirements Specifications," 1998.',
        'ISO/IEC/IEEE 29119, "Software and systems engineering — Software testing," International Standard.',
        'TweetAPI, "Twitter/X Data API Documentation (tw-v2)," https://api.tweetapi.com/ (accessed Nov. 2026).',
        'A. Pak and P. Paroubek, "Twitter as a Corpus for Sentiment Analysis and Opinion Mining," in Proc. LREC, 2010.',
        'B. Pang and L. Lee, "Opinion Mining and Sentiment Analysis," Foundations and Trends in Information Retrieval, vol. 2, no. 1–2, pp. 1–135, 2008.',
    ]
    for i, ref in enumerate(refs, start=1):
        p = add_para(doc, f"[{i}]  {ref}", size=11, align="justify",
                     space_after=8, line_spacing=1.5, first_line=0)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)

    # =====================================================================
    # INDIVIDUAL CONTRIBUTION
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, "SAMPLE INDIVIDUAL CONTRIBUTION REPORT", size=14, bold=True, space_after=20)

    contributions = [
        {
            "focus": (
                "Led backend architecture using FastAPI, implemented TweetAPI integration with "
                "pagination and error handling, and developed the /analyze and profile tweet "
                "retrieval pipelines. Tuned VADER thresholds and validated JSON response contracts."
            ),
            "report": "Chapters 3 (SRS, architecture) and 4 (methodology, API tables).",
            "demo": "Live demonstration of backend endpoints and Profile Analysis flow.",
        },
        {
            "focus": (
                "Designed and implemented the multi-label CATEGORY_KEYWORDS engine, author/"
                "text extraction helpers, and the /topic-search ranking logic. Ensured "
                "de-duplication of posts and robust parsing across TweetAPI payload variants."
            ),
            "report": "Chapter 2 (concepts) and Chapter 4 (topic search implementation & tests).",
            "demo": "Topic Intelligence search, ranking explanation, and edge-case handling.",
        },
        {
            "focus": (
                "Built the React/Vite frontend: sidebar navigation, Profile Analysis forms, "
                "Axios integration, Recharts sentiment and category visualisations, loading/"
                "error/empty states, and overall UI styling."
            ),
            "report": "Chapter 4 (frontend methodology, screenshots) and Chapter 5 (design standards).",
            "demo": "UI walkthrough of charts, stats cards, and analysed post list.",
        },
        {
            "focus": (
                "Implemented Topic Intelligence UI (ranked account table, matching tweet feed), "
                "coordinated integration testing across client and server, prepared quality "
                "checklist, and organised project documentation and report assembly."
            ),
            "report": "Chapter 1 (introduction), Chapter 6 (conclusion/future scope), References.",
            "demo": "End-to-end system demo and Q&A on standards and future work.",
        },
    ]

    for idx, (member, roll) in enumerate(MEMBERS):
        if idx > 0:
            page_break(doc)
        add_heading_custom(doc, PROJECT_NAME, size=14, bold=True, space_after=8)
        add_para(doc, member, size=12, bold=True, align="center", space_after=2)
        add_para(doc, roll, size=12, align="center", space_after=12)

        p = add_para(doc, "", size=12, align="justify", space_after=8, line_spacing=1.5)
        run = p.add_run("Abstract: ")
        set_run_font(run, size=12, bold=True)
        run = p.add_run(
            "TwitterSense-AI is a full-stack platform that analyses public X posts for "
            "sentiment and topical categories and discovers influential accounts for a given "
            "topic, presenting insights through an interactive web dashboard."
        )
        set_run_font(run, size=12)

        p = add_para(doc, "", size=12, align="justify", space_after=8, line_spacing=1.5)
        run = p.add_run("Individual contribution and findings: ")
        set_run_font(run, size=12, bold=True)
        run = p.add_run(contributions[idx]["focus"] + " ")
        set_run_font(run, size=12)
        run = p.add_run(
            "Planning involved breaking the module into milestones, estimating API latency, "
            "and validating outputs against sample public accounts. Technical findings included "
            "the need for defensive JSON parsing, the effectiveness of VADER on short informal "
            "text, and the importance of clear UI feedback during long-running analyses."
        )
        set_run_font(run, size=12)

        p = add_para(doc, "", size=12, align="justify", space_after=8, line_spacing=1.5)
        run = p.add_run("Individual contribution to project report preparation: ")
        set_run_font(run, size=12, bold=True)
        run = p.add_run(contributions[idx]["report"])
        set_run_font(run, size=12)

        p = add_para(doc, "", size=12, align="justify", space_after=8, line_spacing=1.5)
        run = p.add_run("Individual contribution for project presentation and demonstration: ")
        set_run_font(run, size=12, bold=True)
        run = p.add_run(contributions[idx]["demo"])
        set_run_font(run, size=12)

        add_para(doc, "", space_after=30)
        add_para(doc,
                 "Full Signature of Supervisor: __________________          "
                 "Full signature of the student: __________________",
                 size=12, align="left", space_after=8, line_spacing=1.5)

    # =====================================================================
    # PLAGIARISM REPORT
    # =====================================================================
    page_break(doc)
    add_heading_custom(doc, "TURNITIN PLAGIARISM REPORT", size=16, bold=True, space_after=12)
    add_para(doc, "(This report is mandatory for all the projects and plagiarism",
             size=12, align="center", space_after=2)
    add_para(doc, "must be below 25%)",
             size=12, bold=True, align="center", space_after=30)
    add_para(doc, "[Attach the official Turnitin similarity report PDF / printout here.]",
             size=12, italic=True, align="center", space_after=12)
    add_para(doc, "Similarity index target: Less than 25%.",
             size=12, align="center", space_after=8)

    out_path = r"C:\Users\KIIT0001\Desktop\TwitterSense-AI\TwitterSense-AI_Project_Report.docx"
    doc.save(out_path)
    print(f"Saved: {out_path}")
    return out_path


if __name__ == "__main__":
    build_report()
