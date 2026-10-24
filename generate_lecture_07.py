"""
Lecture 06 — PW Basic Maths and Calculus Class 05 (full class)
10-minute formula recap, then graph, coordinate, distance, slope and straight line.
Class 11 Physics  •  CBSE + NEET Foundation  •  2-hour session
"""
from __future__ import annotations

import math
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

BASE = Path(__file__).resolve().parent
OUTPUT = BASE / "class notes" / "Lecture_07_Graph_Shapes_and_Physics_Applications.pptx"

NAVY = RGBColor(8, 30, 63)
BLUE = RGBColor(64, 156, 255)
TEAL = RGBColor(40, 180, 190)
GREEN = RGBColor(90, 214, 128)
AMBER = RGBColor(255, 204, 72)
ORANGE = RGBColor(255, 158, 64)
RED = RGBColor(255, 96, 96)
TEXT = RGBColor(255, 255, 255)
MUTED = RGBColor(176, 182, 192)
STROKE = RGBColor(72, 78, 90)
WHITE = RGBColor(255, 255, 255)
SOFT = RGBColor(22, 24, 28)
BLACK = RGBColor(0, 0, 0)
GOLD = RGBColor(255, 214, 70)
NCERT_FILL = RGBColor(255, 248, 225)
NCERT_LINE = RGBColor(218, 168, 64)
QCOL = ORANGE
WORK = RGBColor(90, 200, 255)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

XL, CW, Y0, YF = 0.65, 12.05, 1.18, 7.12
LH = 0.32
NG = 0.18
PT_SCALE = 0.80


def _bg(slide, color=BLACK):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color


def _solid(slide, x, y, w, h, fill, no_line=True):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if no_line:
        s.line.fill.background()
    else:
        s.line.color.rgb = fill
    return s


def _card_bg(slide, x, y, w, h, fill, border, line_pt=1.0, rounded=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.color.rgb = border
    s.line.width = Pt(line_pt)
    return s


def tx(slide, text, x, y, w, h, size=15, color=TEXT, bold=False,
       align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP, spa=1, lsp=1.1,
       font="Aptos", scaled=True):
    text = (
        text.replace("PW 04 — ", "")
        .replace("PW Class 05 homework — ", "Homework — ")
        .replace("PW Class 05", "")
        .replace("PW intercept example: ", "")
        .replace("PW example: ", "")
    )
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(0)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_before = Pt(0)
    p.space_after = Pt(spa)
    p.line_spacing = lsp
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(size if not scaled else max(7, round(size * PT_SCALE, 1)))
    r.font.bold = bold
    r.font.color.rgb = color
    return box


def _bar(slide, label, color=BLUE):
    _solid(slide, 0, 0, 13.333, 0.29, color)
    if label:
        tx(slide, label.upper(), XL, 0.05, 9, 0.20, 8, WHITE, True, spa=0)


def _title(slide, text):
    tx(slide, text, XL, 0.34, CW, 0.72, 26, GOLD, True, spa=0)


def _sep(slide, y=1.10):
    _solid(slide, XL, y, CW, 0.018, STROKE)


def _footer(slide, n):
    tx(slide, f"{n}", 12.25, YF, 0.8, 0.25, 11, MUTED, True, PP_ALIGN.RIGHT, spa=0)


def add_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


_CONTENT = 0
_ROUGH_DUE = False
_ANSWER_STEPS = {}


def add_rough_work():
    slide = prs.slides.add_slide(BLANK)
    _bg(slide)
    _bar(slide, "Rough work", MUTED)
    tx(slide, "Rough work  /  extra work", XL, 0.34, CW, 0.50, 22, GOLD, True, spa=0)
    _sep(slide)
    _footer(slide, len(prs.slides))
    add_notes(slide, "Empty board for extra working.")


def after_content_slide(defer_rough=False):
    global _CONTENT, _ROUGH_DUE
    _CONTENT += 1
    if defer_rough and _CONTENT % 5 == 0:
        _ROUGH_DUE = True
    elif not defer_rough and (_ROUGH_DUE or _CONTENT % 5 == 0):
        add_rough_work()
        _ROUGH_DUE = False


def new_slide(section, title, bar_color=BLUE, defer_rough=False):
    slide = prs.slides.add_slide(BLANK)
    _bg(slide)
    _bar(slide, section, bar_color)
    _title(slide, title)
    _sep(slide)
    _footer(slide, len(prs.slides))
    after_content_slide(defer_rough)
    return slide


def question_slide(title):
    return new_slide("Question", title, ORANGE, defer_rough=True)


def answer_slide(title="Simple solution"):
    slide = new_slide("Answer", title, GREEN)
    _ANSWER_STEPS[id(slide)] = 0
    return slide


def ncert(slide, formula, x, y, w, h, size=20):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = NCERT_FILL
    s.line.color.rgb = NCERT_LINE
    s.line.width = Pt(1.5)
    tx(slide, formula, x + 0.18, y + 0.16, w - 0.36, h - 0.32,
       size, NAVY, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, spa=0, lsp=1.15)


def fml(slide, text, x, y, w, h=0.44, size=22):
    if id(slide) in _ANSWER_STEPS:
        _ANSWER_STEPS[id(slide)] += 1
        text = f"Step {_ANSWER_STEPS[id(slide)]}:  {text}"
    return tx(slide, text, x, y, w, h, size, WORK, True, spa=0)


def body(slide, text, x, y, w, h=0.40, size=16):
    return tx(slide, text, x, y, w, h, size, TEXT, False, spa=0, scaled=False)


def qtx(slide, text, x, y, w, h=0.44, size=18):
    return tx(slide, text, x, y, w, h, size, QCOL, True, spa=0, scaled=False)


def result_card(slide, text, x, y, w, h=0.36, tsize=18):
    return tx(slide, text, x, y, w, h, max(tsize, 18), GREEN, True, spa=0)


def trap_card(slide, text, x, y, w, h=0.40):
    return tx(slide, text, x, y, w, h, 16, ORANGE, False, spa=0, scaled=False)


def choices(slide, items, x, y, w=None):
    w = w or CW
    for i, item in enumerate(items):
        cy = y + i * 0.40
        circ = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(cy + 0.04), Inches(0.30), Inches(0.30))
        circ.fill.solid()
        circ.fill.fore_color.rgb = TEAL
        circ.line.fill.background()
        tx(slide, str(i + 1), x, cy + 0.06, 0.30, 0.26, 14, WHITE, True, PP_ALIGN.CENTER, spa=0, scaled=False)
        tx(slide, item, x + 0.42, cy + 0.06, w - 0.42, 0.28, 18, TEXT, False, spa=0, scaled=False)


def grid_table(slide, headers, rows, x, y, col_w, row_h=0.42, hsize=14, csize=15):
    col_x = [x]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w)
    for j, ht in enumerate(headers):
        _solid(slide, col_x[j], y, col_w[j], row_h, RGBColor(28, 36, 52), False)
        tx(slide, ht, col_x[j] + 0.04, y, col_w[j] - 0.08, row_h,
           hsize, WHITE, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, spa=0, scaled=False)
    for i, row in enumerate(rows):
        ry = y + row_h + i * row_h
        fill = RGBColor(16, 16, 20) if i % 2 == 0 else SOFT
        for j, val in enumerate(row):
            s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(col_x[j]), Inches(ry), Inches(col_w[j]), Inches(row_h))
            s.fill.solid()
            s.fill.fore_color.rgb = fill
            s.line.color.rgb = STROKE
            s.line.width = Pt(0.5)
            tx(slide, val, col_x[j] + 0.04, ry, col_w[j] - 0.08, row_h,
               csize, WORK if j == 0 else TEXT, j == 0,
               PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE, spa=0, scaled=False)


def tx_3d(slide, text, x, y, w, h, size, face, edge, font="Arial Black"):
    tx(slide, text, x + 0.045, y + 0.045, w, h, size, edge, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)
    tx(slide, text, x, y, w, h, size, face, True, PP_ALIGN.CENTER, spa=0, scaled=False, font=font)


def draw_line(slide, x1, y1, x2, y2, color=WHITE, pt=1.75):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    c.line.color.rgb = color
    c.line.width = Pt(pt)
    return c


def axes(slide, x=7.5, y=1.55, w=4.5, h=4.5, xlab="x", ylab="y"):
    ox, oy = x + w * 0.40, y + h * 0.58
    draw_line(slide, x, oy, x + w, oy, MUTED, 1.5)
    draw_line(slide, ox, y, ox, y + h, MUTED, 1.5)
    tx(slide, xlab, x + w - 0.15, oy + 0.06, 0.3, 0.22, 13, MUTED, True, spa=0)
    tx(slide, ylab, ox + 0.08, y, 0.3, 0.22, 13, MUTED, True, spa=0)
    return ox, oy, w, h


def line_on_axes(slide, ox, oy, dx1, dy1, dx2, dy2, color=GOLD, pt=2.4):
    return draw_line(slide, ox + dx1, oy - dy1, ox + dx2, oy - dy2, color, pt)


# ── COVER ─────────────────────────────────────────────────────────────────────
# COVER
def curve(slide, points, color=GOLD, pt=2.4):
    for (x1,y1),(x2,y2) in zip(points, points[1:]):
        draw_line(slide, x1, y1, x2, y2, color, pt)


def graph_axes(slide, x, y, w, h, xmin, xmax, ymin, ymax, xlab="x", ylab="y"):
    def sx(value):
        return x + (value - xmin) * w / (xmax - xmin)

    def sy(value):
        return y + h - (value - ymin) * h / (ymax - ymin)

    if ymin <= 0 <= ymax:
        draw_line(slide, x, sy(0), x + w, sy(0), MUTED, 1.4)
    if xmin <= 0 <= xmax:
        draw_line(slide, sx(0), y, sx(0), y + h, MUTED, 1.4)
    tx(slide, xlab, x + w - 0.18, sy(0) + 0.05 if ymin <= 0 <= ymax else y + h - 0.22,
       0.35, 0.22, 12, MUTED, True, spa=0)
    tx(slide, ylab, sx(0) + 0.07 if xmin <= 0 <= xmax else x + 0.05, y,
       0.45, 0.22, 12, MUTED, True, spa=0)
    return sx, sy


def plot_function(slide, fn, sx, sy, xmin, xmax, ymin, ymax, color=GOLD, samples=160, pt=2.2):
    previous = None
    for i in range(samples + 1):
        xv = xmin + (xmax - xmin) * i / samples
        try:
            yv = fn(xv)
        except (ValueError, ZeroDivisionError, OverflowError):
            previous = None
            continue
        if not math.isfinite(yv) or yv < ymin or yv > ymax:
            previous = None
            continue
        current = (sx(xv), sy(yv))
        if previous is not None:
            draw_line(slide, previous[0], previous[1], current[0], current[1], color, pt)
        previous = current


def plot_parametric(slide, fx, fy, sx, sy, t0, t1, color=GOLD, samples=180, pt=2.2):
    points = []
    for i in range(samples + 1):
        t = t0 + (t1 - t0) * i / samples
        points.append((sx(fx(t)), sy(fy(t))))
    curve(slide, points, color, pt)


# COVER
slide = prs.slides.add_slide(BLANK)
_bg(slide)
tx_3d(slide, "EXCELLENCE", XL, 0.55, CW, 1.05, 62, GOLD, RGBColor(210,210,220))
tx_3d(slide, "CLASS 11   -   2026-27", XL, 1.62, CW, 0.55, 30, WHITE, RGBColor(90,90,100), font="Aptos")
_solid(slide, 4.15, 2.28, 5.00, 0.028, WHITE)
bx,by,bw=2.05,2.70,9.20
_card_bg(slide,bx,by,bw,2.35,WHITE,WHITE,0)
_card_bg(slide,bx,by+0.72,bw,1.63,BLACK,RGBColor(220,220,220),1.25)
_solid(slide,bx+bw/2,by+0.16,0.018,0.42,RGBColor(40,40,40))
tx(slide,"Physics",bx,by+0.16,bw/2-0.10,0.44,25,BLACK,True,PP_ALIGN.CENTER,spa=0,scaled=False)
tx(slide,"Lecture - 07",bx+bw/2,by+0.16,bw/2,0.44,25,BLACK,True,PP_ALIGN.CENTER,spa=0,scaled=False)
tx(slide,"Basic Maths and Calculus",bx,by+1.05,bw,0.55,32,WHITE,True,PP_ALIGN.CENTER,spa=0,scaled=False,font="Georgia")
tx(slide,"Graph Shapes and Physics Applications",bx,by+1.62,bw,0.40,18,GOLD,True,PP_ALIGN.CENTER,spa=0,scaled=False)
tx(slide,"By - Excellence Programme",XL,5.30,CW,0.42,25,WHITE,True,PP_ALIGN.CENTER,spa=0,scaled=False)
tx(slide,"11 September 2026",XL,5.80,CW,0.32,18,MUTED,True,PP_ALIGN.CENTER,spa=0,scaled=False)
_footer(slide,1); add_notes(slide,"New graph shapes and physics applications."); after_content_slide()

slide=prs.slides.add_slide(BLANK); _bg(slide)
tx(slide,"Topics to be covered",XL,0.38,CW,0.55,32,GOLD,True,spa=0,scaled=False); _solid(slide,XL,0.98,3.40,0.035,GOLD)
topics=["Spring-force graph and perpendicular lines","Modulus graphs and transformations","Rectangular hyperbola and inverse relations","Parabola, square-root graph and transformations","Graphs used in mechanics, electricity and thermal physics","Circle and S.H.M. homework with guided answers"]
for i,item in enumerate(topics):
    yy=1.18+i*0.78; tx(slide,f"{i+1:02d}",XL,yy,0.70,0.50,23,GOLD,True,spa=0,scaled=False); tx(slide,item,XL+0.90,yy,CW-0.90,0.70,18,WHITE,False,spa=0,scaled=False)
_footer(slide,2); add_notes(slide,"Do not reteach AP, GP, slope or line equations."); after_content_slide()

slide=new_slide("2-minute recall","Only the graph tools we need",BLUE)
fml(slide,"rising line: positive slope     falling line: negative slope",XL,Y0,CW,LH,18)
fml(slide,"horizontal line: slope 0     vertical line: slope undefined",XL,Y0+LH+NG,CW,LH,18)
fml(slide,"To draw any graph: make a small value table, mark points, join smoothly.",XL,Y0+2*LH+2*NG,CW,LH,18)

slide=new_slide("Physics graph","Spring force: F = -kx",TEAL)
body(slide,"Hooke's law says the spring force is opposite to displacement.",XL,Y0,6.1)
fml(slide,"F = -kx",XL,Y0+LH+NG,6.1,LH,24)
body(slide,"The graph is a straight line through the origin with negative slope -k.",XL,Y0+2*LH+2*NG,6.1)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,-2.5,2.5,-2.5,2.5,"x","F")
plot_function(slide,lambda x:-x,sx,sy,-2.5,2.5,-2.5,2.5,RED)

slide=new_slide("How it works","Theory of the spring-force graph",TEAL)
body(slide,"1. x is displacement from the equilibrium position.",XL,Y0,CW)
body(slide,"2. k is the spring constant. A larger k means a stiffer spring.",XL,Y0+LH+NG,CW)
body(slide,"3. The minus sign means force acts opposite to displacement.",XL,Y0+2*LH+2*NG,CW)
fml(slide,"x>0 gives F<0     and     x<0 gives F>0",XL,Y0+3*LH+3*NG,CW,LH,19)
result_card(slide,"Slope of the F-x graph = -k. Its magnitude gives the spring constant.",XL,Y0+4*LH+4*NG,CW,0.42,19)

slide=new_slide("Learn","Perpendicular straight lines",TEAL)
body(slide,"Two perpendicular lines meet at 90 degrees.",XL,Y0,CW)
ncert(slide,"m1 x m2 = -1",XL,Y0+LH+NG,CW,0.95,25)
fml(slide,"If m1 = tan(theta), then m2 = tan(theta+90 degrees) = -1/tan(theta)",XL,Y0+2*LH+2*NG+0.85,CW,LH,17)
trap_card(slide,"This rule is used only when both slopes are defined.",XL,Y0+3*LH+3*NG+0.85,CW)

slide=new_slide("How it works","Theory of perpendicular slopes",TEAL)
body(slide,"Rotating a line by 90 degrees changes its direction completely.",XL,Y0,CW)
fml(slide,"m1=tan(theta)",XL,Y0+LH+NG,CW,LH,20)
fml(slide,"m2=tan(theta+90 degrees)=-cot(theta)=-1/tan(theta)",XL,Y0+2*LH+2*NG,CW,LH,18)
fml(slide,"Therefore m1 x m2 = tan(theta) x [-1/tan(theta)] = -1",XL,Y0+3*LH+3*NG,CW,LH,18)
trap_card(slide,"A horizontal line is perpendicular to a vertical line; the product rule cannot be used because vertical slope is undefined.",XL,Y0+4*LH+4*NG,CW,0.50)

slide=question_slide("Perpendicular-line check")
qtx(slide,"A line has slope 2. What is the slope of a perpendicular line?",XL,Y0,CW,0.45,18)
slide=answer_slide()
fml(slide,"Use m1 x m2 = -1",XL,Y0,CW,LH,20)
fml(slide,"2 x m2 = -1",XL,Y0+LH+NG,CW,LH,20)
fml(slide,"m2 = -1/2",XL,Y0+2*LH+2*NG,CW,LH,22)
result_card(slide,"Perpendicular slope = -1/2",XL,Y0+3*LH+3*NG,CW,0.38,22)

slide=new_slide("Learn","Meaning of modulus",TEAL)
body(slide,"Modulus gives distance from zero, so it can never be negative.",XL,Y0,CW)
ncert(slide,"|x| = x for x>=0     and     |x| = -x for x<0",XL,Y0+LH+NG,CW,1.05,20)
grid_table(slide,["x","-2","-1","0","1","2"],[["|x|","2","1","0","1","2"]],XL,Y0+2*LH+2*NG+1.0,[2,2,2,2,2,2],0.60)

slide=new_slide("How it works","Theory of modulus graphs",TEAL)
body(slide,"Modulus changes every negative output into its positive magnitude.",XL,Y0,CW)
fml(slide,"For x>=0: |x|=x, so use the line y=x.",XL,Y0+LH+NG,CW,LH,18)
fml(slide,"For x<0: |x|=-x, so use the line y=-x.",XL,Y0+2*LH+2*NG,CW,LH,18)
body(slide,"The two straight parts meet where the expression inside the modulus becomes zero.",XL,Y0+3*LH+3*NG,CW)
result_card(slide,"For y=a|x-h|+k: vertex=(h,k); sign of a gives up/down and |a| gives steepness.",XL,Y0+4*LH+4*NG,CW,0.48,18)

slide=new_slide("Graph","Graph of y = |x|",GREEN)
body(slide,"The graph is V-shaped, has vertex (0,0), and lies on or above the x-axis.",XL,Y0,6.1)
fml(slide,"left branch: y=-x     right branch: y=x",XL,Y0+LH+NG,6.1,LH,18)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,-3,3,-0.5,3.5)
plot_function(slide,lambda x:abs(x),sx,sy,-3,3,-0.5,3.5,GOLD)

slide=new_slide("Graph","Graph of y = -|x|",GREEN)
body(slide,"The minus sign reflects y=|x| in the x-axis.",XL,Y0,6.1)
fml(slide,"vertex stays at (0,0); the V opens downward",XL,Y0+LH+NG,6.1,LH,18)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,-3,3,-3.5,0.5)
plot_function(slide,lambda x:-abs(x),sx,sy,-3,3,-3.5,0.5,RED)

slide=question_slide("Draw y = |x-2|")
qtx(slide,"Make a value table and draw y=|x-2|.",XL,Y0,CW,0.45,18)
slide=answer_slide()
grid_table(slide,["x","0","1","2","3","4"],[["y","2","1","0","1","2"]],XL,Y0,[1.0,0.8,0.8,0.8,0.8,0.8],0.62)
fml(slide,"The smallest value is 0 at x=2.",XL,Y0+1.55,5.2,LH,19)
result_card(slide,"Shifted 2 units right; vertex=(2,0).",XL,Y0+2.10,5.2,0.40,18)
sx,sy=graph_axes(slide,7.0,1.40,5.1,4.8,-2,6,-0.5,4.5)
plot_function(slide,lambda x:abs(x-2),sx,sy,-2,6,-0.5,4.5,GOLD)

slide=question_slide("Draw y = |4-2x|")
qtx(slide,"Find the vertex and draw y=|4-2x|.",XL,Y0,CW,0.45,18)
slide=answer_slide()
fml(slide,"Set the inside equal to zero: 4-2x=0",XL,Y0,5.8,LH,18)
fml(slide,"x=2, so the vertex is (2,0)",XL,Y0+LH+NG,5.8,LH,18)
fml(slide,"y=|4-2x|=2|x-2|",XL,Y0+2*LH+2*NG,5.8,LH,18)
result_card(slide,"V-shape, vertex (2,0), steeper.",XL,Y0+3*LH+3*NG,5.8,0.40,18)
sx,sy=graph_axes(slide,7.0,1.40,5.1,4.8,-1,5,-0.5,6)
plot_function(slide,lambda x:abs(4-2*x),sx,sy,-1,5,-0.5,6,TEAL)

slide=new_slide("Compare","Position and steepness of modulus graphs",TEAL)
body(slide,"Gold: |x|     Red: -|x|     Teal: |x-2|     Green: 2|x-2|",XL,Y0,CW)
sx,sy=graph_axes(slide,1.2,1.75,10.8,4.7,-4,6,-5,6)
plot_function(slide,lambda x:abs(x),sx,sy,-4,6,-5,6,GOLD)
plot_function(slide,lambda x:-abs(x),sx,sy,-4,6,-5,6,RED)
plot_function(slide,lambda x:abs(x-2),sx,sy,-4,6,-5,6,TEAL)
plot_function(slide,lambda x:2*abs(x-2),sx,sy,-4,6,-5,6,GREEN)

slide=new_slide("Learn","Rectangular hyperbola: y = k/x",TEAL)
body(slide,"In an inverse relation, one quantity increases while the other decreases.",XL,Y0,6.2)
ncert(slide,"y = k/x     or     xy = k",XL,Y0+LH+NG,6.2,0.90,22)
body(slide,"The curve approaches both axes but never touches them.",XL,Y0+2*LH+2*NG+0.82,6.2)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,-4,4,-4,4)
plot_function(slide,lambda x:1/x,sx,sy,-4,-0.08,-4,4,TEAL)
plot_function(slide,lambda x:1/x,sx,sy,0.08,4,-4,4,TEAL)

slide=new_slide("Graph method","How to draw an inverse graph",AMBER)
body(slide,"1. Write the equation in the form y=k/x.",XL,Y0,CW)
body(slide,"2. Make a value table. Avoid x=0 because division by zero is undefined.",XL,Y0+LH+NG,CW)
body(slide,"3. Mark positive and negative x values.",XL,Y0+2*LH+2*NG,CW)
body(slide,"4. Join smoothly; do not let the curve touch an axis.",XL,Y0+3*LH+3*NG,CW)

slide=question_slide("Inverse-square relation")
qtx(slide,"A force follows F=k/r^2. What happens when r becomes 2r?",XL,Y0,CW,0.45,18)
slide=answer_slide()
fml(slide,"Original force: F=k/r^2",XL,Y0,CW,LH,20)
fml(slide,"New force: F'=k/(2r)^2=k/(4r^2)",XL,Y0+LH+NG,CW,LH,20)
result_card(slide,"F'=F/4. Doubling distance makes force one-fourth.",XL,Y0+2*LH+2*NG,CW,0.40,21)

slide=new_slide("Physics graph","De Broglie wavelength",GREEN)
body(slide,"A moving particle has wavelength inversely proportional to momentum.",XL,Y0,6.2)
ncert(slide,"lambda = h/p",XL,Y0+LH+NG,6.2,0.90,25)
body(slide,"So lambda versus p is a decreasing rectangular hyperbola.",XL,Y0+2*LH+2*NG+0.82,6.2)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,0,5,0,5,"p","lambda")
plot_function(slide,lambda p:1/p,sx,sy,0.05,5,0,5,GOLD)

slide=new_slide("Derivation","Wavelength and kinetic energy",TEAL)
fml(slide,"K=p^2/(2m)  gives  p=sqrt(2mK)",XL,Y0,CW,LH,19)
fml(slide,"lambda=h/p=h/sqrt(2mK)",XL,Y0+LH+NG,CW,LH,19)
ncert(slide,"lambda is proportional to 1/sqrt(K)",XL,Y0+2*LH+2*NG,CW,0.95,22)
body(slide,"Therefore lambda decreases when kinetic energy increases.",XL,Y0+3*LH+2*NG+0.88,CW)

slide=new_slide("Family of curves","y = k/x for different k",TEAL)
body(slide,"Every positive k gives a similar decreasing curve.",XL,Y0,6.2)
body(slide,"A larger k places the curve farther from the origin.",XL,Y0+LH+NG,6.2)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,0,5,0,5)
plot_function(slide,lambda x:1/x,sx,sy,0.05,5,0,5,TEAL)
plot_function(slide,lambda x:2/x,sx,sy,0.05,5,0,5,GOLD)

slide=new_slide("Thermal physics","Pressure-volume isotherm",GREEN)
body(slide,"For a fixed amount of ideal gas at constant temperature:",XL,Y0,6.2)
ncert(slide,"PV=nRT  gives  P=nRT/V",XL,Y0+LH+NG,6.2,0.90,20)
body(slide,"P is inversely proportional to V, so the graph is a rectangular hyperbola.",XL,Y0+2*LH+2*NG+0.82,6.2)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,0,5,0,5,"V","P")
plot_function(slide,lambda v:2/v,sx,sy,0.05,5,0,5,GREEN)

slide=new_slide("How it works","Theory of an isothermal P-V graph",TEAL)
body(slide,"The amount of gas n and temperature T stay constant, so nRT is one constant.",XL,Y0,CW)
fml(slide,"PV=constant     therefore     P=constant/V",XL,Y0+LH+NG,CW,LH,20)
body(slide,"If volume doubles, pressure becomes half. If volume triples, pressure becomes one-third.",XL,Y0+2*LH+2*NG,CW)
body(slide,"Pressure becomes very large near V=0 and approaches zero for very large V.",XL,Y0+3*LH+3*NG,CW)
result_card(slide,"A higher-temperature isotherm lies farther from the origin because nRT is larger.",XL,Y0+4*LH+4*NG,CW,0.42,19)

slide=question_slide("Two gas isotherms")
qtx(slide,"Two P-V curves are drawn for temperatures T1 and T2. The T2 curve is farther from the origin. Which temperature is higher?",XL,Y0,CW,0.62,17)
slide=answer_slide()
fml(slide,"At the same V: P=nRT/V",XL,Y0,5.8,LH,19)
fml(slide,"Higher T gives higher P.",XL,Y0+LH+NG,5.8,LH,19)
result_card(slide,"Farther curve: T2>T1.",XL,Y0+2*LH+2*NG,5.8,0.40,20)
sx,sy=graph_axes(slide,7.0,1.40,5.1,4.8,0,5,0,5,"V","P")
plot_function(slide,lambda v:1/v,sx,sy,0.05,5,0,5,TEAL)
plot_function(slide,lambda v:2/v,sx,sy,0.05,5,0,5,GOLD)
tx(slide,"T2",10.9,2.25,0.5,0.25,13,GOLD,True,spa=0)
tx(slide,"T1",10.9,3.20,0.5,0.25,13,TEAL,True,spa=0)

slide=new_slide("Compare","y=1/x and y=1/x^2",TEAL)
body(slide,"Gold: y=1/x     Teal: y=1/x^2. Both approach the axes without touching them.",XL,Y0,CW)
sx,sy=graph_axes(slide,1.2,1.75,10.8,4.7,-4,4,-4,5)
plot_function(slide,lambda x:1/x,sx,sy,-4,-0.06,-4,5,GOLD)
plot_function(slide,lambda x:1/x,sx,sy,0.06,4,-4,5,GOLD)
plot_function(slide,lambda x:1/(x*x),sx,sy,-4,-0.06,-4,5,TEAL)
plot_function(slide,lambda x:1/(x*x),sx,sy,0.06,4,-4,5,TEAL)

slide=new_slide("Electricity","Electric field versus distance",GREEN)
body(slide,"Gold: 1/r     Teal: 1/r^2     Red: 1/r^3",XL,Y0,CW)
sx,sy=graph_axes(slide,1.2,1.75,10.8,4.7,0,5,0,5,"r","E")
plot_function(slide,lambda r:1/r,sx,sy,0.05,5,0,5,GOLD)
plot_function(slide,lambda r:1/(r*r),sx,sy,0.05,5,0,5,TEAL)
plot_function(slide,lambda r:1/(r*r*r),sx,sy,0.05,5,0,5,RED)

slide=new_slide("How it works","Theory of electric-field falloff",TEAL)
body(slide,"As distance increases, the same electric effect spreads over a larger region.",XL,Y0,CW)
fml(slide,"line charge: E is proportional to 1/r",XL,Y0+LH+NG,CW,LH,19)
fml(slide,"point charge: E is proportional to 1/r^2",XL,Y0+2*LH+2*NG,CW,LH,19)
fml(slide,"dipole far away: E is proportional to 1/r^3",XL,Y0+3*LH+3*NG,CW,LH,19)
result_card(slide,"At large r, 1/r^3 decreases fastest and 1/r decreases slowest.",XL,Y0+4*LH+4*NG,CW,0.42,20)

slide=new_slide("Learn","Parabola: y = x^2",TEAL)
body(slide,"Squaring gives the same y-value for +x and -x.",XL,Y0,6.1)
grid_table(slide,["x","-2","-1","0","1","2"],[["y=x^2","4","1","0","1","4"]],XL,Y0+0.60,[1.6,0.9,0.9,0.9,0.9,0.9],0.56,12,13)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,-3,3,-1,6)
plot_function(slide,lambda x:x*x,sx,sy,-3,3,-1,6,GOLD)

slide=new_slide("How it works","Theory of a parabola",TEAL)
body(slide,"The general vertical parabola is y=a(x-h)^2+k.",XL,Y0,CW)
fml(slide,"vertex=(h,k)     axis of symmetry: x=h",XL,Y0+LH+NG,CW,LH,19)
fml(slide,"a>0: opens upward     a<0: opens downward",XL,Y0+2*LH+2*NG,CW,LH,19)
body(slide,"A larger |a| makes the parabola narrower; a smaller |a| makes it wider.",XL,Y0+3*LH+3*NG,CW)
result_card(slide,"Find the vertex first, then calculate symmetric points on both sides.",XL,Y0+4*LH+4*NG,CW,0.42,20)

slide=question_slide("Draw y = x^2 - 4")
qtx(slide,"Find the vertex and x-intercepts of y=x^2-4, then sketch the graph.",XL,Y0,CW,0.48,18)
slide=answer_slide()
fml(slide,"Vertex: x=0 gives y=-4, so vertex=(0,-4)",XL,Y0,5.8,LH,18)
fml(slide,"For x-intercepts: x^2-4=0",XL,Y0+LH+NG,5.8,LH,18)
fml(slide,"x=+2 or x=-2",XL,Y0+2*LH+2*NG,5.8,LH,19)
result_card(slide,"Cuts x-axis at (-2,0) and (2,0).",XL,Y0+3*LH+3*NG,5.8,0.42,18)
sx,sy=graph_axes(slide,7.0,1.40,5.1,4.8,-4,4,-5,7)
plot_function(slide,lambda x:x*x-4,sx,sy,-4,4,-5,7,GOLD)

slide=new_slide("Transform","Vertical shifts of y=x^2",TEAL)
body(slide,"Gold: y=x^2     Teal: y=x^2+4     Red: y=x^2-4",XL,Y0,CW)
sx,sy=graph_axes(slide,1.2,1.75,10.8,4.7,-4,4,-6,10)
plot_function(slide,lambda x:x*x,sx,sy,-4,4,-6,10,GOLD)
plot_function(slide,lambda x:x*x+4,sx,sy,-4,4,-6,10,TEAL)
plot_function(slide,lambda x:x*x-4,sx,sy,-4,4,-6,10,RED)

slide=new_slide("Transform","Downward parabolas",TEAL)
body(slide,"Gold: -x^2     Teal: -x^2+4     Red: -x^2-4",XL,Y0,CW)
sx,sy=graph_axes(slide,1.2,1.75,10.8,4.7,-4,4,-10,6)
plot_function(slide,lambda x:-x*x,sx,sy,-4,4,-10,6,GOLD)
plot_function(slide,lambda x:-x*x+4,sx,sy,-4,4,-10,6,TEAL)
plot_function(slide,lambda x:-x*x-4,sx,sy,-4,4,-10,6,RED)

slide=new_slide("Learn","Square-root graph: y = sqrt(x)",TEAL)
body(slide,"A real square root needs x>=0.",XL,Y0,6.1)
grid_table(slide,["x","0","1","4","9"],[["sqrt(x)","0","1","2","3"]],XL,Y0+0.58,[1.8,1.1,1.1,1.1,1.1],0.58)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,0,9,0,4)
plot_function(slide,lambda x:math.sqrt(x),sx,sy,0,9,0,4,TEAL)

slide=new_slide("How it works","Theory of the square-root graph",TEAL)
body(slide,"y=sqrt(x) is the inverse of y=x^2 when x is restricted to x>=0.",XL,Y0,CW)
fml(slide,"domain: x>=0     range: y>=0",XL,Y0+LH+NG,CW,LH,20)
body(slide,"It begins at (0,0), rises quickly, and then becomes flatter.",XL,Y0+2*LH+2*NG,CW)
fml(slide,"y=sqrt(x-h)+k starts at (h,k)",XL,Y0+3*LH+3*NG,CW,LH,19)
trap_card(slide,"For real-number graphs, the expression inside the square root cannot be negative.",XL,Y0+4*LH+4*NG,CW)

slide=new_slide("Compare","y=x^2, y=x and y=sqrt(x)",TEAL)
body(slide,"Gold: y=x^2     White: y=x     Teal: y=sqrt(x). All meet at (0,0) and (1,1).",XL,Y0,CW)
sx,sy=graph_axes(slide,1.2,1.75,10.8,4.7,0,4,0,6)
plot_function(slide,lambda x:x*x,sx,sy,0,4,0,6,GOLD)
plot_function(slide,lambda x:x,sx,sy,0,4,0,6,WHITE)
plot_function(slide,lambda x:math.sqrt(x),sx,sy,0,4,0,6,TEAL)

slide=new_slide("Mechanics","Kinetic energy versus speed",GREEN)
body(slide,"For constant mass:",XL,Y0,6.1)
ncert(slide,"K = (1/2)mv^2",XL,Y0+LH+NG,6.1,0.90,24)
body(slide,"K is proportional to v^2, so K versus v is an upward parabola through the origin.",XL,Y0+2*LH+2*NG+0.82,6.1)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,0,4,0,8,"v","K")
plot_function(slide,lambda v:0.5*v*v,sx,sy,0,4,0,8,GOLD)

slide=new_slide("How it works","Theory of the kinetic-energy graph",TEAL)
body(slide,"Mass is constant, so K depends on the square of speed.",XL,Y0,CW)
fml(slide,"v becomes 2v  ->  K becomes 4K",XL,Y0+LH+NG,CW,LH,20)
fml(slide,"v becomes 3v  ->  K becomes 9K",XL,Y0+2*LH+2*NG,CW,LH,20)
body(slide,"Speed cannot be negative, so the physical K-v graph uses only v>=0.",XL,Y0+3*LH+3*NG,CW)
result_card(slide,"The curve starts at the origin and becomes steeper as speed increases.",XL,Y0+4*LH+4*NG,CW,0.42,20)

slide=new_slide("Oscillations","Simple pendulum: T versus length",GREEN)
body(slide,"For small oscillations:",XL,Y0,6.1)
ncert(slide,"T = 2pi sqrt(L/g)",XL,Y0+LH+NG,6.1,0.90,22)
body(slide,"T is proportional to sqrt(L): it rises, but its slope gradually decreases.",XL,Y0+2*LH+2*NG+0.82,6.1)
sx,sy=graph_axes(slide,7.5,1.45,4.6,4.7,0,9,0,4,"L","T")
plot_function(slide,lambda length:math.sqrt(length),sx,sy,0,9,0,4,TEAL)

slide=new_slide("How it works","Theory of the pendulum graph",TEAL)
body(slide,"For one place, g is constant. The period depends on the square root of length.",XL,Y0,CW)
fml(slide,"T is proportional to sqrt(L)",XL,Y0+LH+NG,CW,LH,21)
fml(slide,"L becomes 4L  ->  T becomes 2T",XL,Y0+2*LH+2*NG,CW,LH,20)
fml(slide,"L becomes 9L  ->  T becomes 3T",XL,Y0+3*LH+3*NG,CW,LH,20)
result_card(slide,"T^2 versus L is a straight line through the origin.",XL,Y0+4*LH+4*NG,CW,0.42,21)

slide=new_slide("Gravitation","Kepler's third law",GREEN)
body(slide,"For planets orbiting the same central body:",XL,Y0,5.8)
ncert(slide,"T^2 is proportional to R^3",XL,Y0+LH+NG,5.8,0.95,21)
fml(slide,"T is proportional to R^(3/2)",XL,Y0+2*LH+2*NG+0.88,5.8,LH,19)
body(slide,"T versus R is a rising curved graph.",XL,Y0+3*LH+3*NG+0.88,5.8)
sx,sy=graph_axes(slide,7.0,1.40,5.1,4.8,0,4,0,8,"R","T")
plot_function(slide,lambda r:r**1.5,sx,sy,0,4,0,8,GOLD)

slide=new_slide("How it works","Theory of Kepler's third-law graph",TEAL)
body(slide,"T is orbital period and R is the orbital radius or semi-major axis.",XL,Y0,CW)
fml(slide,"T^2/R^3 = constant for bodies orbiting the same central mass",XL,Y0+LH+NG,CW,LH,18)
body(slide,"Therefore T^2 plotted against R^3 gives a straight line through the origin.",XL,Y0+2*LH+2*NG,CW)
body(slide,"But T plotted directly against R is a curved graph because T is proportional to R^(3/2).",XL,Y0+3*LH+3*NG,CW)
trap_card(slide,"The constant changes if the central mass changes.",XL,Y0+4*LH+4*NG,CW)

slide=question_slide("Kinetic energy versus mass")
qtx(slide,"Momentum p is constant. Draw the relation between kinetic energy K and mass m.",XL,Y0,CW,0.48,18)
slide=answer_slide()
fml(slide,"K=p^2/(2m)",XL,Y0,5.8,LH,20)
fml(slide,"With p constant, K is proportional to 1/m.",XL,Y0+LH+NG,5.8,LH,18)
result_card(slide,"Decreasing rectangular hyperbola.",XL,Y0+2*LH+2*NG,5.8,0.40,19)
sx,sy=graph_axes(slide,7.0,1.40,5.1,4.8,0,5,0,5,"m","K")
plot_function(slide,lambda m:1/m,sx,sy,0.05,5,0,5,GOLD)

slide=question_slide("Homework: draw x^2 + y^2 = 9")
qtx(slide,"Identify the curve, its centre and its radius.",XL,Y0,CW,0.45,18)
slide=answer_slide()
fml(slide,"Compare with x^2+y^2=R^2",XL,Y0,5.8,LH,19)
fml(slide,"R^2=9, so R=3",XL,Y0+LH+NG,5.8,LH,19)
result_card(slide,"Centre (0,0), radius 3.",XL,Y0+2*LH+2*NG,5.8,0.40,20)
sx,sy=graph_axes(slide,7.15,1.40,4.8,4.8,-4,4,-4,4)
plot_parametric(slide,lambda t:3*math.cos(t),lambda t:3*math.sin(t),sx,sy,0,2*math.pi,GOLD)

slide=question_slide("Homework: S.H.M. velocity-position graph")
qtx(slide,"For v=omega sqrt(A^2-x^2), draw the graph between v and x.",XL,Y0,CW,0.48,18)
slide=answer_slide()
fml(slide,"v^2=omega^2(A^2-x^2)",XL,Y0,5.8,LH,18)
fml(slide,"x^2/A^2 + v^2/(omega A)^2 = 1",XL,Y0+LH+NG,5.8,LH,18)
body(slide,"Upper half of an ellipse because speed is non-negative.",XL,Y0+2*LH+2*NG,5.8)
result_card(slide,"Maximum at x=0; zero at x=+A and -A.",XL,Y0+3*LH+3*NG,5.8,0.42,18)
sx,sy=graph_axes(slide,7.0,1.40,5.1,4.8,-1.2,1.2,0,1.2,"x","v")
plot_parametric(slide,lambda t:math.cos(t),lambda t:math.sin(t),sx,sy,0,math.pi,TEAL)

slide=new_slide("5-minute summary","Graph-shape lock lines",TEAL)
fml(slide,"|x| -> V shape     1/x -> rectangular hyperbola",XL,Y0,CW,LH,18)
fml(slide,"x^2 -> parabola     sqrt(x) -> rising curve for x>=0",XL,Y0+LH+NG,CW,LH,18)
fml(slide,"inverse power: larger denominator power -> faster decrease",XL,Y0+2*LH+2*NG,CW,LH,18)
fml(slide,"m1 x m2 = -1 for perpendicular non-vertical lines",XL,Y0+3*LH+3*NG,CW,LH,18)

slide=new_slide("After class","Practice",AMBER)
body(slide,"Draw every graph once using a small value table.",XL,Y0,CW)
body(slide,"Label both axes and mark important points or intercepts.",XL,Y0+LH+NG,CW)
body(slide,"Complete the circle and S.H.M. questions without looking at the answer slide.",XL,Y0+2*LH+2*NG,CW)

prs.save(OUTPUT)
print(f"Created {len(prs.slides)} slides")
print(OUTPUT)
