from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


BLUE = RGBColor(12, 84, 141)
DARK_BLUE = RGBColor(8, 56, 95)
LIGHT_BLUE = RGBColor(230, 242, 252)
GREEN = RGBColor(39, 158, 83)
ORANGE = RGBColor(233, 140, 28)
GRAY = RGBColor(92, 106, 118)
WHITE = RGBColor(255, 255, 255)
BLACK = RGBColor(30, 30, 30)


def style_text(shape, text, size=20, bold=False, color=BLACK, align=PP_ALIGN.LEFT, italic=False):
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = align
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.italic = italic
    p.font.color.rgb = color
    p.font.name = "Calibri"


def style_bullets(shape, items, size=20, color=BLACK, spacing=7):
    tf = shape.text_frame
    tf.clear()
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.name = "Calibri"
        p.level = 0
        p.space_after = Pt(spacing)


def add_header(slide, title, right_text=None):
    head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.92))
    head.fill.solid()
    head.fill.fore_color.rgb = BLUE
    head.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.35), Inches(0.16), Inches(8.5), Inches(0.56))
    style_text(title_box, title, size=28, bold=True, color=WHITE)

    if right_text:
        right = slide.shapes.add_textbox(Inches(8.7), Inches(0.2), Inches(4.2), Inches(0.5))
        style_text(right, right_text, size=20, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, italic=True)


def add_canvas(slide):
    body = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.92), Inches(13.333), Inches(6.58))
    body.fill.solid()
    body.fill.fore_color.rgb = RGBColor(250, 250, 248)
    body.line.fill.background()

    shade = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(6.98), Inches(13.333), Inches(0.52))
    shade.fill.solid()
    shade.fill.fore_color.rgb = RGBColor(240, 240, 237)
    shade.line.fill.background()


def add_tagline(slide, text="Empowering Teachers, Transforming Learning", color=DARK_BLUE):
    ribbon = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(6.95), Inches(13.333), Inches(0.55))
    ribbon.fill.solid()
    ribbon.fill.fore_color.rgb = color
    ribbon.line.fill.background()
    txt = slide.shapes.add_textbox(Inches(0.3), Inches(7.05), Inches(12.7), Inches(0.35))
    style_text(txt, text, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, italic=True)


def add_image_placeholder(slide, left, top, width, height, label):
    ph = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    ph.fill.solid()
    ph.fill.fore_color.rgb = LIGHT_BLUE
    ph.line.color.rgb = RGBColor(180, 200, 220)
    t = slide.shapes.add_textbox(left + Inches(0.15), top + height / 2 - Inches(0.2), width - Inches(0.3), Inches(0.4))
    style_text(t, label, size=16, bold=True, color=GRAY, align=PP_ALIGN.CENTER)


def slide_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(21, 104, 166)
    bg.line.fill.background()

    cloud = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(5.6), Inches(13.333), Inches(1.9))
    cloud.fill.solid()
    cloud.fill.fore_color.rgb = RGBColor(33, 123, 185)
    cloud.line.fill.background()

    title = s.shapes.add_textbox(Inches(1.0), Inches(1.1), Inches(11.3), Inches(1.2))
    style_text(title, "Teacher Copilot", size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    sub = s.shapes.add_textbox(Inches(1.0), Inches(2.35), Inches(11.3), Inches(0.8))
    style_text(sub, "AI-Powered Teacher Productivity & Student Learning Insights", size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    map_box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.7), Inches(3.0), Inches(5.9), Inches(3.0))
    map_box.fill.solid()
    map_box.fill.fore_color.rgb = RGBColor(178, 224, 173)
    map_box.line.color.rgb = RGBColor(87, 146, 86)
    map_txt = s.shapes.add_textbox(Inches(4.0), Inches(4.25), Inches(5.3), Inches(0.5))
    style_text(map_txt, "Africa rollout map", size=18, bold=True, color=RGBColor(52, 100, 52), align=PP_ALIGN.CENTER)

    for x, y in [
        (4.2, 3.5), (5.3, 3.9), (6.4, 3.4), (7.4, 4.2), (8.2, 3.7),
        (4.9, 4.7), (6.0, 5.0), (7.0, 5.2), (8.1, 4.8),
    ]:
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(0.26), Inches(0.26))
        dot.fill.solid()
        dot.fill.fore_color.rgb = ORANGE
        dot.line.fill.background()


def slide_problem(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "The Problem")
    bullet_box = s.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(6.3), Inches(4.9))
    style_bullets(
        bullet_box,
        [
            "Teachers overloaded with manual grading",
            "Delayed feedback and generic responses",
            "Lack of insights for personalized learning",
            "Students often unaware of learning gaps",
        ],
        size=24,
    )
    add_image_placeholder(s, Inches(7.0), Inches(1.35), Inches(5.6), Inches(5.35), "Teacher under workload pressure")
    add_tagline(s)


def slide_mission(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Our Mission & Vision")

    m = s.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(6.2), Inches(4.7))
    tf = m.text_frame
    tf.clear()

    p1 = tf.paragraphs[0]
    p1.text = "Mission: Empower teachers to transform learning with AI"
    p1.font.size = Pt(24)
    p1.font.bold = True
    p1.font.color.rgb = GREEN
    p1.font.name = "Calibri"

    p2 = tf.add_paragraph()
    p2.text = "Vision: A future where every student receives personalized education"
    p2.font.size = Pt(24)
    p2.font.bold = True
    p2.font.color.rgb = RGBColor(176, 29, 44)
    p2.font.name = "Calibri"
    p2.space_before = Pt(26)

    add_image_placeholder(s, Inches(7.1), Inches(1.55), Inches(5.3), Inches(4.55), "Teacher using AI companion")
    add_tagline(s, color=RGBColor(171, 128, 64))


def slide_solution(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Our Solution:", right_text="Teacher Copilot")

    left_lbl = s.shapes.add_textbox(Inches(1.0), Inches(1.35), Inches(4.8), Inches(0.35))
    style_text(left_lbl, "Current Workflow", size=17, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    right_lbl = s.shapes.add_textbox(Inches(7.5), Inches(1.35), Inches(4.8), Inches(0.35))
    style_text(right_lbl, "Optimized Workflow", size=17, bold=True, color=ORANGE, align=PP_ALIGN.CENTER)

    left_items = ["Manual Grading", "Basic Feedback", "Slow Reports"]
    right_items = ["AI Grading", "Instant Feedback", "Analytics & Alerts"]

    y = 1.9
    for i in range(3):
        l = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(y), Inches(4.9), Inches(1.0))
        l.fill.solid()
        l.fill.fore_color.rgb = GREEN
        l.line.fill.background()
        t = s.shapes.add_textbox(Inches(1.05), Inches(y + 0.3), Inches(4.4), Inches(0.4))
        style_text(t, left_items[i], size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        r = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), Inches(y), Inches(5.1), Inches(1.0))
        r.fill.solid()
        r.fill.fore_color.rgb = ORANGE
        r.line.fill.background()
        t2 = s.shapes.add_textbox(Inches(7.2), Inches(y + 0.3), Inches(4.7), Inches(0.4))
        style_text(t2, right_items[i], size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(5.85), Inches(y + 0.23), Inches(1.0), Inches(0.54))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = BLUE
        arrow.line.fill.background()
        y += 1.45


def slide_market(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Market Opportunity")

    tiers = [
        ("TAM", "2.5M Schools\n$600M+", RGBColor(27, 113, 178), Inches(1.5), Inches(2.0), Inches(4.8), Inches(1.2)),
        ("SAM", "Targetable 32M Schools", RGBColor(54, 168, 142), Inches(1.8), Inches(3.3), Inches(4.2), Inches(1.1)),
        ("SOM", "Pilot 5 Schools\n$20M", RGBColor(221, 147, 34), Inches(2.1), Inches(4.45), Inches(3.6), Inches(1.0)),
    ]
    for label, text, color, l, t, w, h in tiers:
        shp = s.shapes.add_shape(MSO_SHAPE.TRAPEZOID, l, t, w, h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = color
        shp.line.fill.background()
        tb = s.shapes.add_textbox(l + Inches(0.2), t + Inches(0.18), w - Inches(0.4), h - Inches(0.2))
        style_text(tb, f"{label}\n{text}", size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_image_placeholder(s, Inches(7.0), Inches(1.65), Inches(5.6), Inches(4.8), "Education market landscape")


def slide_business(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Business Model")

    cards = [
        (Inches(0.9), RGBColor(217, 133, 29), "B2B (Schools)\n$6,000 per school/year"),
        (Inches(6.9), GREEN, "B2G (Government)\n$600,000 per country/year"),
    ]
    for left, color, text in cards:
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.8), Inches(5.5), Inches(3.2))
        card.fill.solid()
        card.fill.fore_color.rgb = color
        card.line.fill.background()

        banner = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(4.6), Inches(5.5), Inches(0.4))
        banner.fill.solid()
        banner.fill.fore_color.rgb = RGBColor(245, 245, 245)
        banner.line.fill.background()

        tb = s.shapes.add_textbox(left + Inches(0.2), Inches(2.25), Inches(5.1), Inches(1.6))
        style_text(tb, text, size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def slide_value(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Value Proposition")

    panel = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.7), Inches(6.2), Inches(4.9))
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(219, 235, 248)
    panel.line.fill.background()

    box = s.shapes.add_textbox(Inches(0.9), Inches(2.0), Inches(5.7), Inches(4.4))
    style_bullets(
        box,
        [
            "For Teachers: Save time and gain actionable insights",
            "For Students: Personalized learning paths",
            "For Schools & Governments: Scalable, cost-effective solution",
        ],
        size=23,
    )

    add_image_placeholder(s, Inches(7.2), Inches(1.8), Inches(5.4), Inches(4.8), "Teacher + dashboard workflow")


def slide_traction(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Our Traction")

    left = s.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.0), Inches(4.5))
    style_bullets(
        left,
        [
            "5 schools launched in Rwanda & Ghana",
            "Pipeline: 10 schools in Senegal & Kenya",
            "Positive pilot results and teacher retention gains",
        ],
        size=23,
    )

    add_image_placeholder(s, Inches(6.9), Inches(1.7), Inches(5.8), Inches(4.9), "Students collaborating with laptops")


def slide_revenue(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Revenue Projections")

    rows = [("Year 1", "$5.4M", BLUE), ("Year 2", "$14.4M", DARK_BLUE), ("Year 3", "$248M+", ORANGE)]
    y = 1.8
    for year, amt, color in rows:
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(y), Inches(6.2), Inches(1.15))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        t = s.shapes.add_textbox(Inches(0.9), Inches(y + 0.32), Inches(2.2), Inches(0.5))
        style_text(t, year, size=24, bold=True, color=WHITE)
        v = s.shapes.add_textbox(Inches(2.5), Inches(y + 0.3), Inches(3.8), Inches(0.5))
        style_text(v, amt, size=30, bold=True, color=WHITE)
        y += 1.2

    base_x = Inches(8.1)
    heights = [1.2, 1.9, 2.7]
    colors = [RGBColor(221, 146, 62), GREEN, RGBColor(134, 198, 79)]
    for i, h in enumerate(heights):
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, base_x + Inches(0.75 * i), Inches(6.0 - h), Inches(0.55), Inches(h))
        r.fill.solid()
        r.fill.fore_color.rgb = colors[i]
        r.line.fill.background()

    growth = s.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(7.1), Inches(3.3), Inches(4.8), Inches(2.8))
    growth.fill.solid()
    growth.fill.fore_color.rgb = RGBColor(39, 145, 88)
    growth.fill.transparency = 0.18
    growth.line.fill.background()


def slide_gtm(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Go-To-Market Strategy")

    labels = [("Year 1", "$5.4M", Inches(0.9)), ("Year 2", "$14.4M", Inches(4.3)), ("Year 3", "$248M+", Inches(7.8))]
    for year, amount, x in labels:
        y1 = s.shapes.add_textbox(x, Inches(5.45), Inches(2.2), Inches(0.4))
        style_text(y1, year, size=20, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
        y2 = s.shapes.add_textbox(x, Inches(3.25), Inches(2.5), Inches(0.8))
        style_text(y2, amount, size=30, bold=True, color=BLACK, align=PP_ALIGN.CENTER)

    bars = [
        (Inches(1.2), 0.7, BLUE),
        (Inches(2.0), 0.9, DARK_BLUE),
        (Inches(4.8), 1.6, BLUE),
        (Inches(5.6), 1.6, DARK_BLUE),
        (Inches(8.2), 2.1, ORANGE),
        (Inches(9.0), 2.7, RGBColor(200, 76, 40)),
        (Inches(9.8), 3.2, RGBColor(184, 64, 34)),
    ]
    for x, h, color in bars:
        r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, Inches(5.3 - h), Inches(0.55), Inches(h))
        r.fill.solid()
        r.fill.fore_color.rgb = color
        r.line.fill.background()

    arrow = s.shapes.add_shape(MSO_SHAPE.UP_ARROW, Inches(0.8), Inches(2.0), Inches(11.4), Inches(3.2))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = RGBColor(35, 142, 86)
    arrow.fill.transparency = 0.22
    arrow.line.fill.background()


def slide_technology(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Technology Overview")

    tech_items = ["AI Grading Engine", "Analytics Dashboard", "Cloud-Based Platform", "LMS Integrations"]
    y = 1.8
    for item in tech_items:
        icon = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), Inches(y), Inches(0.35), Inches(0.35))
        icon.fill.solid()
        icon.fill.fore_color.rgb = BLUE
        icon.line.fill.background()
        tb = s.shapes.add_textbox(Inches(1.3), Inches(y - 0.03), Inches(5.6), Inches(0.45))
        style_text(tb, item, size=22, bold=True, color=DARK_BLUE)
        y += 1.1

    add_image_placeholder(s, Inches(7.0), Inches(1.8), Inches(5.6), Inches(4.8), "Students using cloud-based learning")


def slide_impact(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Impact & Metrics")

    metrics = s.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.0), Inches(4.8))
    style_bullets(
        metrics,
        [
            "Teacher hours saved",
            "Student engagement +40%",
            "Faster feedback: 75% reduction",
            "Improved outcomes across pilot schools",
        ],
        size=23,
    )

    chart_data = CategoryChartData()
    chart_data.categories = ["Teacher Time", "Engagement", "Other"]
    chart_data.add_series("Impact", (50, 40, 10))
    chart = s.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(7.4),
        Inches(1.7),
        Inches(5.0),
        Inches(4.9),
        chart_data,
    ).chart
    chart.has_legend = False
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.number_format = "0%"
    chart.plots[0].data_labels.position = 2

    points = chart.series[0].points
    points[0].format.fill.solid()
    points[0].format.fill.fore_color.rgb = RGBColor(59, 126, 191)
    points[1].format.fill.solid()
    points[1].format.fill.fore_color.rgb = GREEN
    points[2].format.fill.solid()
    points[2].format.fill.fore_color.rgb = RGBColor(180, 210, 235)


def slide_funding(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Funding Ask")

    info = s.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(5.3), Inches(4.9))
    style_bullets(
        info,
        [
            "Raising $3M",
            "Product development",
            "Sales & marketing",
            "School onboarding",
            "Operations",
        ],
        size=23,
    )

    chart_data = CategoryChartData()
    chart_data.categories = ["Product", "Sales", "Onboarding", "Operations"]
    chart_data.add_series("Allocation", (40, 30, 20, 10))
    pie = s.shapes.add_chart(
        XL_CHART_TYPE.PIE,
        Inches(5.6),
        Inches(2.0),
        Inches(4.0),
        Inches(4.0),
        chart_data,
    ).chart
    pie.has_legend = False
    pie.plots[0].has_data_labels = True
    pie.plots[0].data_labels.number_format = "0%"
    pie.plots[0].data_labels.position = 2

    colors = [RGBColor(220, 145, 29), RGBColor(52, 132, 198), GREEN, RGBColor(166, 196, 231)]
    for i, point in enumerate(pie.series[0].points):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = colors[i]

    add_image_placeholder(s, Inches(9.7), Inches(1.7), Inches(2.8), Inches(4.8), "Africa focus")


def slide_team(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Our Team")

    members = [
        ("Founder & CEO", RGBColor(223, 202, 187)),
        ("CTO", RGBColor(198, 176, 160)),
        ("Head of Partnerships", RGBColor(185, 169, 155)),
        ("Education Advisors", RGBColor(222, 201, 186)),
    ]

    x = 0.8
    for role, color in members:
        card = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.8), Inches(2.8), Inches(4.8))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(244, 244, 244)
        card.line.color.rgb = RGBColor(222, 222, 222)

        img = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x + 0.25), Inches(2.1), Inches(2.3), Inches(2.8))
        img.fill.solid()
        img.fill.fore_color.rgb = color
        img.line.fill.background()

        lbl = s.shapes.add_textbox(Inches(x + 0.1), Inches(5.25), Inches(2.6), Inches(1.0))
        style_text(lbl, role, size=16, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
        x += 3.1


def slide_closing(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_canvas(s)
    add_header(s, "Join Us to", right_text="Transform African Education")

    center = s.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.5), Inches(2.5))
    style_text(center, "Partner with us\nTo empower teachers &\nstudents across Africa", size=27, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER, italic=True)

    add_image_placeholder(s, Inches(7.0), Inches(1.8), Inches(5.8), Inches(3.8), "Regional expansion map")

    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(6.45), Inches(13.333), Inches(1.05))
    bar.fill.solid()
    bar.fill.fore_color.rgb = BLUE
    bar.line.fill.background()

    contact = s.shapes.add_textbox(Inches(0.7), Inches(6.7), Inches(12.0), Inches(0.6))
    style_text(contact, "Contact Us: email@example.com    |    www.teachercopilot.com", size=19, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def build():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_cover(prs)
    slide_problem(prs)
    slide_mission(prs)
    slide_solution(prs)
    slide_market(prs)
    slide_business(prs)
    slide_value(prs)
    slide_traction(prs)
    slide_revenue(prs)
    slide_gtm(prs)
    slide_technology(prs)
    slide_impact(prs)
    slide_funding(prs)
    slide_team(prs)
    slide_closing(prs)

    prs.save("Teacher_Copilot_Pitch_Deck.pptx")


if __name__ == "__main__":
    build()
