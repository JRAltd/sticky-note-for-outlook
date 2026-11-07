import os
from PIL import Image, ImageDraw, ImageFont

OUTPUT_DIR = "instagram-carousel"
os.makedirs(OUTPUT_DIR, exist_ok=True)

W, H = 1080, 1350

FONT_BOLD_PATH = "C:/Windows/Fonts/segoeuib.ttf"
FONT_REG_PATH = "C:/Windows/Fonts/segoeui.ttf"

fonts = {
    "title_xl": ImageFont.truetype(FONT_BOLD_PATH, 62),
    "title": ImageFont.truetype(FONT_BOLD_PATH, 54),
    "subtitle": ImageFont.truetype(FONT_REG_PATH, 32),
    "heading": ImageFont.truetype(FONT_BOLD_PATH, 40),
    "subheading": ImageFont.truetype(FONT_BOLD_PATH, 30),
    "body": ImageFont.truetype(FONT_REG_PATH, 28),
    "body_bold": ImageFont.truetype(FONT_BOLD_PATH, 28),
    "badge": ImageFont.truetype(FONT_BOLD_PATH, 22),
    "footer": ImageFont.truetype(FONT_BOLD_PATH, 26),
    "small": ImageFont.truetype(FONT_REG_PATH, 24)
}

def draw_gradient_bg(draw, w, h, top_color, bottom_color):
    for y in range(h):
        ratio = y / h
        r = int(top_color[0] * (1 - ratio) + bottom_color[0] * ratio)
        g = int(top_color[1] * (1 - ratio) + bottom_color[1] * ratio)
        b = int(top_color[2] * (1 - ratio) + bottom_color[2] * ratio)
        draw.line([(0, y), (w, y)], fill=(r, g, b))

def draw_rounded_rect(draw, box, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)

def draw_dot(draw, x, y, radius, color):
    draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=color)

def add_header(img, draw, slide_num, total=5):
    # App brand pill
    draw_rounded_rect(draw, (70, 60, 420, 116), radius=28, fill=(240, 246, 255), outline=(219, 234, 254), width=2)
    draw_dot(draw, 102, 88, 7, (26, 115, 232))
    draw.text((122, 74), "EMAIL STICKY NOTES", font=fonts["badge"], fill=(26, 115, 232))
    
    # Slide counter
    counter_text = f"{slide_num} / {total}"
    draw_rounded_rect(draw, (W - 170, 60, W - 70, 116), radius=28, fill=(241, 245, 249), outline=(226, 232, 240), width=1)
    draw.text((W - 146, 74), counter_text, font=fonts["badge"], fill=(100, 116, 139))

def add_footer(img, draw, next_hint="SWIPE LEFT  ->"):
    draw.line([(70, H - 120), (W - 70, H - 120)], fill=(226, 232, 240), width=2)
    draw.text((70, H - 90), "@jraltd  |  notes.jraltdinc.us", font=fonts["small"], fill=(148, 163, 184))
    if next_hint:
        draw.text((W - 290, H - 92), next_hint, font=fonts["footer"], fill=(26, 115, 232))

# ==========================================
# SLIDE 1: COVER
# ==========================================
def create_slide_1():
    img = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(img)
    draw_gradient_bg(draw, W, H, (255, 255, 255), (238, 246, 255))
    add_header(img, draw, 1)

    # Announcement badge
    badge_box = (W//2 - 290, 180, W//2 + 290, 240)
    draw_rounded_rect(draw, badge_box, radius=30, fill=(236, 253, 245), outline=(167, 243, 208), width=2)
    draw_dot(draw, W//2 - 250, 210, 6, (5, 150, 105))
    draw.text((W//2 - 230, 196), "OFFICIALLY APPROVED & LIVE ON GOOGLE", font=fonts["badge"], fill=(5, 150, 105))

    # App Icon
    icon_path = "../Sticky Note for Gmail/assets/icon-512.png"
    if os.path.exists(icon_path):
        icon = Image.open(icon_path).convert("RGBA").resize((210, 210), Image.Resampling.LANCZOS)
        mask = Image.new("L", (210, 210), 0)
        mask_draw = ImageDraw.Draw(mask)
        mask_draw.rounded_rectangle((0, 0, 210, 210), radius=46, fill=255)
        
        # Soft shadow behind icon
        draw_rounded_rect(draw, (W//2 - 120, 280, W//2 + 120, 520), radius=54, fill=(224, 231, 255))
        img.paste(icon, (W//2 - 105, 295), mask)

    # Big Titles
    draw.text((70, 560), "Gmail Just Got", font=fonts["title_xl"], fill=(15, 23, 42))
    draw.text((70, 638), "Sticky Notes.", font=fonts["title_xl"], fill=(26, 115, 232))

    subtitle = "Attach private, color-coded notes and interactive\nchecklists directly to your email threads."
    draw.text((70, 735), subtitle, font=fonts["subtitle"], fill=(71, 85, 105), spacing=12)

    # Feature preview pills
    features = [
        ("Contextual to Active Email Thread", (26, 115, 232)),
        ("6 Color Themes & 4 Priority Levels", (217, 119, 6)),
        ("Interactive Checklists in Sidebar", (5, 150, 105)),
        ("100% Private (No External Database)", (109, 40, 217))
    ]
    
    start_y = 860
    for text, dot_color in features:
        draw_rounded_rect(draw, (70, start_y, W - 70, start_y + 64), radius=14, fill=(255, 255, 255), outline=(226, 232, 240), width=2)
        draw_dot(draw, 105, start_y + 32, 7, dot_color)
        draw.text((130, start_y + 16), text, font=fonts["subheading"], fill=(30, 41, 59))
        start_y += 78

    add_footer(img, draw, "SWIPE LEFT  ->")
    img.save(os.path.join(OUTPUT_DIR, "slide1_cover.png"), quality=95)
    print("Slide 1 generated.")

# ==========================================
# SLIDE 2: THE PROBLEM
# ==========================================
def create_slide_2():
    img = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(img)
    draw_gradient_bg(draw, W, H, (255, 255, 255), (248, 250, 252))
    add_header(img, draw, 2)

    draw.text((70, 170), "The Daily Email Struggle", font=fonts["title"], fill=(15, 23, 42))
    draw.text((70, 248), "Managing tasks inside an inbox is broken.", font=fonts["subtitle"], fill=(100, 116, 139))

    problems = [
        ("Problem 1: Lost Context",
         "You star or flag an email for follow-up, but days later you forget why you flagged it and what needs to be done."),
        ("Problem 2: Cluttered Inboxes",
         "Sending quick emails to yourself just to keep notes on existing conversation threads creates even more inbox chaos."),
        ("Problem 3: Tool Fatigue",
         "Switching back and forth between your email, separate to-do apps, and physical paper notes on your desk.")
    ]

    card_y = 330
    for title, desc in problems:
        draw_rounded_rect(draw, (70, card_y, W - 70, card_y + 215), radius=20, fill=(254, 242, 242), outline=(254, 202, 202), width=2)
        
        # Red tag
        draw_rounded_rect(draw, (100, card_y + 24, 185, card_y + 58), radius=8, fill=(239, 68, 68))
        draw.text((118, card_y + 29), "PAIN", font=fonts["badge"], fill=(255, 255, 255))
        draw.text((205, card_y + 26), title, font=fonts["subheading"], fill=(185, 28, 28))
        
        words = desc.split()
        lines = []
        cur = ""
        for w in words:
            if len(cur + " " + w) > 52:
                lines.append(cur)
                cur = w
            else:
                cur = cur + " " + w if cur else w
        if cur: lines.append(cur)
        draw.text((100, card_y + 85), "\n".join(lines), font=fonts["body"], fill=(71, 85, 105), spacing=8)
        card_y += 245

    # Bridge card
    draw_rounded_rect(draw, (70, card_y, W - 70, card_y + 110), radius=18, fill=(238, 242, 255), outline=(199, 210, 254), width=2)
    draw_dot(draw, 110, card_y + 55, 8, (67, 56, 202))
    draw.text((135, card_y + 38), "What if your notes lived directly inside Gmail?", font=fonts["subheading"], fill=(67, 56, 202))

    add_footer(img, draw, "THE SOLUTION  ->")
    img.save(os.path.join(OUTPUT_DIR, "slide2_problem.png"), quality=95)
    print("Slide 2 generated.")

# ==========================================
# SLIDE 3: FEATURES
# ==========================================
def create_slide_3():
    img = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(img)
    draw_gradient_bg(draw, W, H, (248, 250, 252), (241, 245, 249))
    add_header(img, draw, 3)

    draw.text((70, 160), "Meet Email Sticky Notes", font=fonts["title"], fill=(15, 23, 42))
    draw.text((70, 238), "Native companion sidebar built right into Gmail.", font=fonts["subtitle"], fill=(71, 85, 105))

    shot_path = "../Sticky Note for Gmail/store-assets/screenshot-1-light-1280x800.png"
    if os.path.exists(shot_path):
        shot = Image.open(shot_path).convert("RGB")
        crop_box = (500, 30, 1260, 770)
        shot_crop = shot.crop(crop_box).resize((940, 520), Image.Resampling.LANCZOS)
        
        draw_rounded_rect(draw, (66, 316, W - 66, 864), radius=20, fill=(203, 213, 225))
        img.paste(shot_crop, (70, 320))

    highlights = [
        ("6 Color Themes", "Color-code by project, client, or team category."),
        ("Action Checklists", "Check and uncheck tasks directly inside the sidebar."),
        ("1-Click Templates", "Instant To-Do, Meeting Summary, and Follow-Up.")
    ]
    
    col_w = (W - 140 - 40) // 3
    for i, (h_title, h_sub) in enumerate(highlights):
        x = 70 + i * (col_w + 20)
        draw_rounded_rect(draw, (x, 900, x + col_w, 1180), radius=16, fill=(255, 255, 255), outline=(226, 232, 240), width=2)
        
        # Pill header
        draw_rounded_rect(draw, (x + 18, 920, x + 70, 948), radius=6, fill=(240, 246, 255))
        draw.text((x + 28, 924), f"0{i+1}", font=fonts["badge"], fill=(26, 115, 232))
        
        draw.text((x + 18, 965), h_title, font=fonts["subheading"], fill=(15, 23, 42))
        
        words = h_sub.split()
        sub_lines = []
        cur = ""
        for w in words:
            if len(cur + " " + w) > 17:
                sub_lines.append(cur)
                cur = w
            else:
                cur = cur + " " + w if cur else w
        if cur: sub_lines.append(cur)
        draw.text((x + 18, 1025), "\n".join(sub_lines), font=fonts["small"], fill=(100, 116, 139), spacing=6)

    add_footer(img, draw, "PRIVACY FIRST  ->")
    img.save(os.path.join(OUTPUT_DIR, "slide3_features.png"), quality=95)
    print("Slide 3 generated.")

# ==========================================
# SLIDE 4: PRIVACY & MULTI-PLATFORM
# ==========================================
def create_slide_4():
    img = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(img)
    draw_gradient_bg(draw, W, H, (255, 255, 255), (248, 250, 252))
    add_header(img, draw, 4)

    draw.text((70, 160), "Private. Secure. Everywhere.", font=fonts["title"], fill=(15, 23, 42))
    draw.text((70, 238), "Enterprise-grade privacy with zero compromise.", font=fonts["subtitle"], fill=(71, 85, 105))

    pillars = [
        ("100% Private (Zero External Servers)",
         "Your notes never touch a third-party server. Everything is stored strictly inside your own private Google Account via Apps Script UserProperties.",
         (236, 253, 245), (167, 243, 208), (4, 120, 87)),

        ("Desktop Web + Mobile (iOS & Android)",
         "Access your notes anywhere: in Gmail's right-hand companion bar on desktop, and directly in the official Gmail mobile app on phone or tablet.",
         (239, 246, 255), (191, 219, 254), (29, 78, 216)),

        ("Cross-Platform Portability (Outlook)",
         "Seamlessly export and import your complete notes archive between Gmail and Microsoft Outlook with one click using structured JSON backups.",
         (254, 243, 199), (253, 230, 138), (180, 83, 9))
    ]

    card_y = 320
    for idx, (title, desc, bg_c, border_c, title_c) in enumerate(pillars):
        draw_rounded_rect(draw, (70, card_y, W - 70, card_y + 255), radius=22, fill=bg_c, outline=border_c, width=2)
        
        # Pill number
        draw_rounded_rect(draw, (105, card_y + 28, 160, card_y + 60), radius=6, fill=border_c)
        draw.text((120, card_y + 32), f"0{idx+1}", font=fonts["badge"], fill=title_c)
        
        draw.text((180, card_y + 28), title, font=fonts["subheading"], fill=title_c)
        
        words = desc.split()
        lines = []
        cur = ""
        for w in words:
            if len(cur + " " + w) > 52:
                lines.append(cur)
                cur = w
            else:
                cur = cur + " " + w if cur else w
        if cur: lines.append(cur)
        draw.text((105, card_y + 90), "\n".join(lines), font=fonts["body"], fill=(51, 65, 85), spacing=8)
        card_y += 285

    add_footer(img, draw, "GET IT FREE  ->")
    img.save(os.path.join(OUTPUT_DIR, "slide4_privacy.png"), quality=95)
    print("Slide 4 generated.")

# ==========================================
# SLIDE 5: CTA / INSTALL
# ==========================================
def create_slide_5():
    img = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(img)
    draw_gradient_bg(draw, W, H, (238, 246, 255), (255, 255, 255))
    add_header(img, draw, 5)

    icon_path = "../Sticky Note for Gmail/assets/icon-512.png"
    if os.path.exists(icon_path):
        icon = Image.open(icon_path).convert("RGBA").resize((170, 170), Image.Resampling.LANCZOS)
        mask = Image.new("L", (170, 170), 0)
        mask_draw = ImageDraw.Draw(mask)
        mask_draw.rounded_rectangle((0, 0, 170, 170), radius=38, fill=255)
        img.paste(icon, (W//2 - 85, 180), mask)

    draw.text((W//2 - 240, 380), "Install in 1-Click", font=fonts["title"], fill=(15, 23, 42))
    draw.text((W//2 - 275, 455), "Free on Google Workspace Marketplace", font=fonts["subtitle"], fill=(71, 85, 105))

    draw_rounded_rect(draw, (70, 540, W - 70, 930), radius=24, fill=(255, 255, 255), outline=(226, 232, 240), width=2)
    
    steps = [
        ("01", "Tap the Link in Bio", "Direct link to our verified Google Workspace listing."),
        ("02", "Click 'Install'", "Google automatically connects it to your Gmail companion bar."),
        ("03", "Start Taking Notes!", "Open any email in Gmail and click the Sticky Notes icon.")
    ]
    
    step_y = 575
    for num, s_title, s_sub in steps:
        draw_rounded_rect(draw, (105, step_y, 165, step_y + 45), radius=8, fill=(240, 246, 255))
        draw.text((120, step_y + 8), num, font=fonts["badge"], fill=(26, 115, 232))
        
        draw.text((185, step_y + 4), s_title, font=fonts["subheading"], fill=(15, 23, 42))
        draw.text((185, step_y + 46), s_sub, font=fonts["small"], fill=(100, 116, 139))
        step_y += 115

    # Big CTA Button
    btn_box = (70, 970, W - 70, 1080)
    draw_rounded_rect(draw, btn_box, radius=24, fill=(26, 115, 232))
    draw.text((W//2 - 215, 1005), "TAP LINK IN BIO TO INSTALL", font=fonts["heading"], fill=(255, 255, 255))

    # Social rating
    draw.text((W//2 - 260, 1115), "Found it helpful? Please leave us a 5-star review!", font=fonts["small"], fill=(71, 85, 105))

    add_footer(img, draw, "SHARE WITH A FRIEND")
    img.save(os.path.join(OUTPUT_DIR, "slide5_cta.png"), quality=95)
    print("Slide 5 generated.")

if __name__ == "__main__":
    create_slide_1()
    create_slide_2()
    create_slide_3()
    create_slide_4()
    create_slide_5()
    print("Successfully regenerated all 5 clean slides.")
