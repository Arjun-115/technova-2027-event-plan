from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

def rgb(r,g,b): return RGBColor(r,g,b)

BG=rgb(10,15,30); ACCENT=rgb(99,102,241); GREEN=rgb(16,185,129)
YELLOW=rgb(245,158,11); RED=rgb(239,68,68); PURPLE=rgb(139,92,246)
WHITE=rgb(226,232,240); MUTED=rgb(100,116,139); CARD=rgb(22,32,50); DARK=rgb(13,27,46)

def bg(s,c=None):
    c=c or BG
    f=s.background.fill; f.solid(); f.fore_color.rgb=c

def txt(s,t,l,tp,w,h,sz=11,bold=False,col=None,align=PP_ALIGN.LEFT,italic=False):
    col=col or WHITE
    b=s.shapes.add_textbox(l,tp,w,h); tf=b.text_frame; tf.word_wrap=True
    p=tf.paragraphs[0]; p.alignment=align; r=p.add_run(); r.text=t
    r.font.size=Pt(sz); r.font.bold=bold; r.font.italic=italic; r.font.color.rgb=col
    return b

def rect(s,l,tp,w,h,fc,bc=None,bw=None):
    sh=s.shapes.add_shape(1,l,tp,w,h)
    sh.fill.solid(); sh.fill.fore_color.rgb=fc
    if bc: sh.line.color.rgb=bc; sh.line.width=bw or Pt(0.5)
    else: sh.line.fill.background()
    return sh

prs=Presentation()
prs.slide_width=Inches(13.33); prs.slide_height=Inches(7.5)
BL=prs.slide_layouts[6]

# ── SLIDE 1 TITLE ──────────────────────────────────────────────────────
s=prs.slides.add_slide(BL); bg(s)
rect(s,0,0,Inches(4),Inches(7.5),DARK)
rect(s,Inches(4),Inches(7.3),Inches(9.33),Inches(0.12),ACCENT)
txt(s,"ASSIGNMENT II – PART 2",Inches(0.2),Inches(0.3),Inches(3.6),Inches(0.3),sz=8,col=MUTED)
txt(s,"College Event Planning & Review",Inches(0.2),Inches(0.6),Inches(3.6),Inches(0.35),sz=9,col=MUTED)
txt(s,"TECH\nNOVA\n2027",Inches(0.2),Inches(1.3),Inches(3.6),Inches(1.8),sz=36,bold=True,col=WHITE)
txt(s,"5th Annual National\nTechnical Symposium",Inches(0.2),Inches(3.2),Inches(3.6),Inches(0.6),sz=11,col=ACCENT)
txt(s,"Feb 14-15, 2027",Inches(0.2),Inches(3.95),Inches(3.6),Inches(0.35),sz=13,bold=True,col=YELLOW)
txt(s,"Sri Sairam Engineering College, Chennai",Inches(0.2),Inches(4.4),Inches(3.6),Inches(0.4),sz=9,col=MUTED)
txt(s,"Arjun V B  |  711523BCB009",Inches(0.2),Inches(6.9),Inches(3.6),Inches(0.35),sz=9,col=MUTED)

txt(s,"A National-Level Technical Fest",Inches(4.3),Inches(0.5),Inches(8.7),Inches(0.5),sz=22,bold=True,col=WHITE)
txt(s,"Covering: Goal  Audience  Timeline  Budget  Promotion  Approvals  Roles  Recording  Post-Event Review",
    Inches(4.3),Inches(1.1),Inches(8.7),Inches(0.35),sz=9,col=MUTED)

stats=[("800+","Participants",Inches(4.3),ACCENT),("Rs.2.45L","Total Budget",Inches(6.6),GREEN),
       ("Rs.42k","Prize Pool",Inches(8.9),YELLOW),("8","Events",Inches(11.2),PURPLE)]
for v,l,x,c in stats:
    rect(s,x,Inches(1.65),Inches(2.1),Inches(0.85),CARD,c,Pt(0.5))
    txt(s,v,x+Inches(0.1),Inches(1.7),Inches(1.9),Inches(0.4),sz=20,bold=True,col=c)
    txt(s,l,x+Inches(0.1),Inches(2.1),Inches(1.9),Inches(0.3),sz=9,col=MUTED)

events=[("Code Sprint","Rs.5k"),("Paper Presentation","Rs.4k"),("Project Expo","Rs.6k"),
        ("Hackathon 24hr","Rs.15k"),("Tech Quiz","Rs.3k"),("Web Design","Rs.4k"),
        ("Debugging","Rs.2k"),("Circuit Design","Rs.3k")]
xs=[Inches(4.3),Inches(6.6),Inches(8.9),Inches(11.2)]
for i,(n,p) in enumerate(events):
    x=xs[i%4]; y=Inches(2.8) if i<4 else Inches(3.85)
    rect(s,x,y,Inches(2.1),Inches(0.8),rgb(20,30,55),ACCENT,Pt(0.5))
    txt(s,n,x+Inches(0.1),y+Inches(0.07),Inches(1.9),Inches(0.35),sz=10,col=WHITE)
    txt(s,p,x+Inches(0.1),y+Inches(0.42),Inches(1.9),Inches(0.28),sz=9,col=GREEN)

rect(s,Inches(4.3),Inches(5.0),Inches(8.7),Inches(0.7),CARD,PURPLE,Pt(0.5))
txt(s,"40+ Colleges",Inches(4.5),Inches(5.05),Inches(2),Inches(0.3),sz=16,bold=True,col=PURPLE)
txt(s,"Participating",Inches(4.5),Inches(5.35),Inches(2),Inches(0.25),sz=9,col=MUTED)
txt(s,"50 Volunteers",Inches(7.0),Inches(5.05),Inches(2),Inches(0.3),sz=16,bold=True,col=RED)
txt(s,"Deployed",Inches(7.0),Inches(5.35),Inches(2),Inches(0.25),sz=9,col=MUTED)
txt(s,"2 Days",Inches(9.5),Inches(5.05),Inches(2),Inches(0.3),sz=16,bold=True,col=YELLOW)
txt(s,"of Competition",Inches(9.5),Inches(5.35),Inches(2),Inches(0.25),sz=9,col=MUTED)

# ── SLIDE 2 GOAL & TIMELINE ────────────────────────────────────────────
s2=prs.slides.add_slide(BL); bg(s2)
rect(s2,0,0,Inches(13.33),Inches(0.65),DARK)
rect(s2,0,Inches(0.65),Inches(0.07),Inches(6.85),ACCENT)
txt(s2,"2 / 5  |  GOAL  AUDIENCE  TIMING  PREPARATION",Inches(0.2),Inches(0.15),Inches(12),Inches(0.38),sz=9,bold=True,col=ACCENT)
txt(s2,"Goal, Audience & 8-Week Preparation Timeline",Inches(0.2),Inches(0.8),Inches(10),Inches(0.45),sz=20,bold=True,col=WHITE)

rect(s2,Inches(0.2),Inches(1.4),Inches(5.8),Inches(1.45),CARD,ACCENT,Pt(0.5))
txt(s2,"PRIMARY GOAL",Inches(0.35),Inches(1.45),Inches(5.5),Inches(0.3),sz=9,bold=True,col=ACCENT)
txt(s2,"Bridge academia and industry. Provide 800+ students from 40+ colleges\na platform to compete, collaborate and connect with industry professionals\nthrough 8 technical events across 2 days.",
    Inches(0.35),Inches(1.78),Inches(5.5),Inches(0.9),sz=10,col=WHITE)

rect(s2,Inches(6.3),Inches(1.4),Inches(6.7),Inches(1.45),CARD,ACCENT,Pt(0.5))
txt(s2,"AUDIENCE BREAKDOWN",Inches(6.45),Inches(1.45),Inches(6.4),Inches(0.3),sz=9,bold=True,col=ACCENT)
aud=[("External Students (40+ colleges)","600"),("Internal Students","200"),
     ("Industry Judges","15"),("Faculty & Chief Guests","50"),("TOTAL","865")]
for i,(l,v) in enumerate(aud):
    y=Inches(1.78)+i*Inches(0.2)
    c=YELLOW if l=="TOTAL" else WHITE; vc=GREEN if l=="TOTAL" else ACCENT
    txt(s2,l,Inches(6.45),y,Inches(5.0),Inches(0.2),sz=9,col=c,bold=(l=="TOTAL"))
    txt(s2,v,Inches(11.5),y,Inches(1.2),Inches(0.2),sz=9,bold=True,col=vc,align=PP_ALIGN.RIGHT)

txt(s2,"8-WEEK PREPARATION TIMELINE",Inches(0.2),Inches(3.05),Inches(12),Inches(0.32),sz=9,bold=True,col=YELLOW)
tl=[("Dec 21","Form committee, assign coordinators, create comms groups"),
    ("Jan 4","Submit Principal/HOD approval + launch social media + open registration portal"),
    ("Jan 11","Sponsors confirmed + registration opens + chief guest invited"),
    ("Jan 25","Registration closes + participant list finalized + print ID cards"),
    ("Feb 7","3-hr volunteer orientation + venue walkthrough + full dry run"),
    ("Feb 13","Full venue setup + banners + AV/lab tech check"),
    ("Feb 14-15","EVENT DAYS  Inauguration  All 8 Events  Cultural Evening  Valedictory"),
    ("Feb 22","Post-event review + certificates emailed + event report to HOD")]
xs2=[Inches(0.2),Inches(3.4),Inches(6.6),Inches(9.8)]
for i,(dt,ac) in enumerate(tl):
    x=xs2[i%4]; y=Inches(3.45) if i<4 else Inches(5.1)
    ec=GREEN if "EVENT" in ac else ACCENT
    rect(s2,x,y,Inches(3.0),Inches(1.45),CARD,ec,Pt(0.5))
    txt(s2,dt,x+Inches(0.1),y+Inches(0.07),Inches(2.8),Inches(0.28),sz=10,bold=True,col=ec)
    txt(s2,ac,x+Inches(0.1),y+Inches(0.38),Inches(2.8),Inches(0.95),sz=9,col=WHITE)

# ── SLIDE 3 BUDGET & PROMOTION ─────────────────────────────────────────
s3=prs.slides.add_slide(BL); bg(s3)
rect(s3,0,0,Inches(13.33),Inches(0.65),DARK)
rect(s3,0,Inches(0.65),Inches(0.07),Inches(6.85),GREEN)
txt(s3,"3 / 5  |  BUDGET  RESOURCES  PROMOTION  APPROVALS",Inches(0.2),Inches(0.15),Inches(12),Inches(0.38),sz=9,bold=True,col=GREEN)
txt(s3,"Budget, Resources & Promotion Strategy",Inches(0.2),Inches(0.8),Inches(10),Inches(0.45),sz=20,bold=True,col=WHITE)

for v,l,x,c in [("Rs.2,45,000","Total Income",Inches(0.2),GREEN),
                 ("Rs.1,97,250","Expenditure",Inches(2.7),RED),
                 ("Rs.47,750","Net Surplus",Inches(5.2),ACCENT),
                 ("Rs.42,000","Prize Pool",Inches(7.7),YELLOW)]:
    rect(s3,x,Inches(1.45),Inches(2.3),Inches(0.78),CARD,c,Pt(0.5))
    txt(s3,v,x+Inches(0.1),Inches(1.5),Inches(2.1),Inches(0.38),sz=16,bold=True,col=c)
    txt(s3,l,x+Inches(0.1),Inches(1.88),Inches(2.1),Inches(0.25),sz=8,col=MUTED)

txt(s3,"EXPENDITURE BREAKDOWN",Inches(0.2),Inches(2.4),Inches(5),Inches(0.3),sz=8,bold=True,col=MUTED)
exps=[("Prizes Trophies Certificates","Rs.55,000",72,PURPLE),
      ("Food Hospitality (2 days)","Rs.48,000",62,GREEN),
      ("Technical Equipment Rental","Rs.22,000",28,ACCENT),
      ("Printing Branding","Rs.18,000",23,YELLOW),
      ("Stage Decor Lighting","Rs.15,000",19,rgb(251,146,60)),
      ("Photography Video","Rs.12,000",15,rgb(236,72,153)),
      ("Guest Hospitality","Rs.10,000",13,rgb(20,184,166)),
      ("Contingency 5pct","Rs.12,250",16,MUTED)]
for i,(item,amt,pct,c) in enumerate(exps):
    y=Inches(2.78)+i*Inches(0.53)
    txt(s3,item,Inches(0.2),y,Inches(3.2),Inches(0.25),sz=9,col=WHITE)
    txt(s3,amt,Inches(3.4),y,Inches(1.0),Inches(0.25),sz=9,bold=True,col=c,align=PP_ALIGN.RIGHT)
    rect(s3,Inches(0.2),y+Inches(0.27),Inches(4.2),Inches(0.1),rgb(30,58,95))
    rect(s3,Inches(0.2),y+Inches(0.27),Inches(4.2*pct/100),Inches(0.1),c)

txt(s3,"INCOME SOURCES",Inches(5.0),Inches(2.4),Inches(4),Inches(0.3),sz=8,bold=True,col=MUTED)
inc=[("Registration Fees (800 x Rs.150)","Rs.1,20,000"),
     ("Title Sponsor","Rs.50,000"),
     ("Co-Sponsors (2 x Rs.15k)","Rs.30,000"),
     ("College Management Grant","Rs.40,000"),
     ("Exhibition Stall Fees","Rs.5,000")]
for i,(s_,a) in enumerate(inc):
    y=Inches(2.78)+i*Inches(0.56)
    rect(s3,Inches(5.0),y,Inches(4.2),Inches(0.4),CARD,rgb(30,58,95),Pt(0.5))
    txt(s3,s_,Inches(5.1),y+Inches(0.06),Inches(2.8),Inches(0.25),sz=9,col=WHITE)
    txt(s3,a,Inches(7.9),y+Inches(0.06),Inches(1.2),Inches(0.25),sz=9,bold=True,col=GREEN,align=PP_ALIGN.RIGHT)

txt(s3,"PROMOTION CHANNELS",Inches(9.4),Inches(2.4),Inches(3.7),Inches(0.3),sz=8,bold=True,col=MUTED)
pro=[("Instagram","Daily reels countdowns events"),
     ("WhatsApp","Broadcast to 40+ college reps"),
     ("Unstop","External registration listing"),
     ("Email","Outreach to 40 HODs coordinators"),
     ("Physical Posters","5 nearby engineering colleges")]
for i,(ch,desc) in enumerate(pro):
    y=Inches(2.78)+i*Inches(0.56)
    rect(s3,Inches(9.4),y,Inches(3.7),Inches(0.4),CARD,ACCENT,Pt(0.5))
    txt(s3,ch,Inches(9.5),y+Inches(0.04),Inches(1.3),Inches(0.2),sz=9,bold=True,col=ACCENT)
    txt(s3,desc,Inches(10.85),y+Inches(0.06),Inches(2.1),Inches(0.25),sz=8,col=MUTED)

txt(s3,"APPROVAL CHAIN",Inches(0.2),Inches(7.0),Inches(12),Inches(0.3),sz=8,bold=True,col=MUTED)
chain="Student Event Head  -->  Dept. Faculty Coordinator  -->  HOD  -->  Vice Principal  -->  Principal  -->  Management (>Rs.1L)"
txt(s3,chain,Inches(0.2),Inches(7.22),Inches(12.9),Inches(0.2),sz=9,col=WHITE)

# ── SLIDE 4 ROLES & VOLUNTEER ──────────────────────────────────────────
s4=prs.slides.add_slide(BL); bg(s4)
rect(s4,0,0,Inches(13.33),Inches(0.65),DARK)
rect(s4,0,Inches(0.65),Inches(0.07),Inches(6.85),YELLOW)
txt(s4,"4 / 5  |  APPROVALS  EVENT-DAY ROLES  VOLUNTEER DEPLOYMENT",Inches(0.2),Inches(0.15),Inches(12),Inches(0.38),sz=9,bold=True,col=YELLOW)
txt(s4,"Event-Day Roles & Volunteer Operations",Inches(0.2),Inches(0.8),Inches(10),Inches(0.45),sz=20,bold=True,col=WHITE)

txt(s4,"CORE COMMITTEE (9 ROLES)",Inches(0.2),Inches(1.4),Inches(8.5),Inches(0.3),sz=9,bold=True,col=ACCENT)
roles=[("Event Head","Overall coordination, crisis management, final decisions"),
       ("Deputy Head","Logistics, vendor payments, equipment issues"),
       ("Finance","Cash management, sponsor liaison, expense tracking"),
       ("Registration","Check-in, ID cards, attendance, participant data"),
       ("Technical","AV setup, Wi-Fi, lab systems, backup plans"),
       ("Hospitality","Food arrangement, guest reception, VIP transport"),
       ("Stage and MC","Programme anchoring, ceremonies, timing"),
       ("Publicity","Photos, video direction, live social media reels"),
       ("Volunteer Head","Manages all 50 volunteers across 8 stations")]
for i,(r,d) in enumerate(roles):
    col_i=i%3; row_i=i//3
    x=Inches(0.2)+col_i*Inches(4.3)
    y=Inches(1.78)+row_i*Inches(1.0)
    rect(s4,x,y,Inches(4.1),Inches(0.88),CARD,ACCENT,Pt(0.5))
    txt(s4,r,x+Inches(0.12),y+Inches(0.06),Inches(3.85),Inches(0.28),sz=10,bold=True,col=ACCENT)
    txt(s4,d,x+Inches(0.12),y+Inches(0.36),Inches(3.85),Inches(0.42),sz=8.5,col=WHITE)

txt(s4,"VOLUNTEER DEPLOYMENT  (50 total)",Inches(0.2),Inches(4.95),Inches(12),Inches(0.3),sz=9,bold=True,col=YELLOW)
vols=[("Main Gate & Entry","5","ID check, crowd control"),
      ("Registration Desk","8","Participant check-in, kit distribution"),
      ("Seminar Halls A & B","6","Seating, announcements, judge help"),
      ("Computer Labs","8","Monitor participants, technical help"),
      ("Exhibition Hall","6","Project expo setup and supervision"),
      ("Auditorium","6","Ushering, stage assistance, seating"),
      ("Food Area","5","Food distribution, queue management"),
      ("Roving Volunteers","6","General help, message running")]
xs4=[Inches(0.2),Inches(3.4),Inches(6.6),Inches(9.8)]
for i,(st,cnt,duty) in enumerate(vols):
    x=xs4[i%4]; y=Inches(5.35) if i<4 else Inches(6.3)
    rect(s4,x,y,Inches(3.0),Inches(0.75),CARD,YELLOW,Pt(0.5))
    txt(s4,st,x+Inches(0.1),y+Inches(0.05),Inches(2.3),Inches(0.25),sz=9,bold=True,col=WHITE)
    txt(s4,cnt+" volunteers",x+Inches(0.1),y+Inches(0.3),Inches(1.5),Inches(0.2),sz=9,col=YELLOW)
    txt(s4,duty,x+Inches(0.1),y+Inches(0.5),Inches(2.8),Inches(0.2),sz=7.5,col=MUTED)

# ── SLIDE 5 RECORDING REVIEW REFLECTION ───────────────────────────────
s5=prs.slides.add_slide(BL); bg(s5)
rect(s5,0,0,Inches(13.33),Inches(0.65),DARK)
rect(s5,0,Inches(0.65),Inches(0.07),Inches(6.85),PURPLE)
txt(s5,"5 / 5  |  RECORDING  POST-EVENT REVIEW  LESSONS  REFLECTION",Inches(0.2),Inches(0.15),Inches(12),Inches(0.38),sz=9,bold=True,col=PURPLE)
txt(s5,"Documentation, Review & Key Learnings",Inches(0.2),Inches(0.8),Inches(10),Inches(0.45),sz=20,bold=True,col=WHITE)

txt(s5,"RECORDING PLAN",Inches(0.2),Inches(1.4),Inches(4),Inches(0.3),sz=9,bold=True,col=PURPLE)
rec=[("Inauguration & Valedictory","Professional videographer hired"),
     ("All event sessions","3 assigned volunteer photographers"),
     ("Hackathon progress","Time-lapse + periodic photos every hour"),
     ("Attendance","Physical sign-in + Google Form QR at entry"),
     ("Participant feedback","Exit gate QR satisfaction survey"),
     ("Financial records","Real-time expense log by finance coordinator")]
for i,(w,h) in enumerate(rec):
    y=Inches(1.78)+i*Inches(0.53)
    rect(s5,Inches(0.2),y,Inches(4.5),Inches(0.4),CARD,PURPLE,Pt(0.5))
    txt(s5,w,Inches(0.32),y+Inches(0.04),Inches(2.3),Inches(0.2),sz=9,bold=True,col=WHITE)
    txt(s5,h,Inches(0.32),y+Inches(0.22),Inches(4.2),Inches(0.18),sz=8,col=MUTED)

txt(s5,"POST-EVENT SUCCESS METRICS",Inches(4.9),Inches(1.4),Inches(4),Inches(0.3),sz=9,bold=True,col=GREEN)
mets=[("Overall Satisfaction Score","4.0 / 5.0 minimum",GREEN),
      ("Registration to Attendance","80% conversion rate",GREEN),
      ("Event Start Delay","15 minutes maximum",YELLOW),
      ("Technical Incidents","3 incidents maximum",YELLOW),
      ("Social Media Reach","10,000+ impressions",ACCENT),
      ("Budget Adherence","Within 5% of plan",ACCENT)]
for i,(m,t,c) in enumerate(mets):
    y=Inches(1.78)+i*Inches(0.53)
    rect(s5,Inches(4.9),y,Inches(3.8),Inches(0.4),CARD,c,Pt(0.5))
    txt(s5,m,Inches(5.0),y+Inches(0.04),Inches(2.3),Inches(0.2),sz=9,col=WHITE)
    txt(s5,t,Inches(7.3),y+Inches(0.06),Inches(1.3),Inches(0.25),sz=9,bold=True,col=c,align=PP_ALIGN.RIGHT)

txt(s5,"LESSONS FROM TECHNOVA 2026",Inches(8.9),Inches(1.4),Inches(4.2),Inches(0.3),sz=9,bold=True,col=YELLOW)
les=[("Registration bottleneck","200 in 30 min -- assign arrival time slots"),
     ("Wi-Fi failed in coding","Dedicated LAN for labs -- no shared Wi-Fi"),
     ("Lunch 45 min late","Pre-order 3 days ahead + stagger lunch halls"),
     ("12 certificate errors","Participant self-verify portal 1 week before"),
     ("Only 3k social reach","Dedicated reels volunteer + post every 2 hours")]
for i,(issue,fix) in enumerate(les):
    y=Inches(1.78)+i*Inches(0.56)
    rect(s5,Inches(8.9),y,Inches(4.2),Inches(0.45),CARD,YELLOW,Pt(0.5))
    txt(s5,issue,Inches(9.0),y+Inches(0.03),Inches(4.0),Inches(0.2),sz=9,bold=True,col=YELLOW)
    txt(s5,fix,Inches(9.0),y+Inches(0.23),Inches(4.0),Inches(0.18),sz=8,col=MUTED)

rect(s5,Inches(0.2),Inches(5.0),Inches(12.9),Inches(2.3),rgb(20,12,40),PURPLE,Pt(0.8))
txt(s5,"STUDENT REFLECTION",Inches(0.4),Inches(5.08),Inches(5),Inches(0.3),sz=9,bold=True,col=PURPLE)
txt(s5,
    '"Communication gaps cause 80% of event-day problems. Having organized TechNova 2026, the biggest issue was that '
    'volunteer briefing happened 30 minutes before opening -- nobody knew their station. For TechNova 2027 I would '
    'enforce a 3-hour orientation one week before, a broadcast-only WhatsApp group for day-of updates, and a '
    '15-minute morning muster before anything opens. The best part is watching 800 strangers walk in, compete, '
    'make friends, and leave with something they could not find in a classroom."',
    Inches(0.4),Inches(5.45),Inches(9.6),Inches(1.6),sz=9.5,col=WHITE,italic=True)
txt(s5,"-- Arjun V B, 711523BCB009",Inches(10.0),Inches(7.05),Inches(3.0),Inches(0.3),
    sz=9,col=MUTED,align=PP_ALIGN.RIGHT)

prs.save(r"D:\CollegeEvent\TechNova_2027_Event_Plan.pptx")
print("TechNova PPT created successfully!")
