#!/usr/bin/env python3
"""
Robo-Root Presentation Generator (Refined & Perfected)
Generates a 10-slide, highly professional, academic presentation based on Final_Report.pdf.
"""

import os
from PIL import Image, ImageChops
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# -------------------------------------------------------------
# Color Palette Constants
# -------------------------------------------------------------
# Primary Dark Theme (Title Slide)
COLOR_DARK_BG = RGBColor(15, 23, 42)        # Slate 900
COLOR_DARK_CARD = RGBColor(30, 41, 59)      # Slate 800
COLOR_DARK_BORDER = RGBColor(51, 65, 85)    # Slate 700
COLOR_DARK_TEXT = RGBColor(248, 250, 252)   # Slate 50
COLOR_DARK_MUTED = RGBColor(148, 163, 184)  # Slate 400
COLOR_CYAN_ACCENT = RGBColor(56, 189, 248)  # Sky 400
COLOR_BLUE_ACCENT = RGBColor(96, 165, 250)  # Blue 400

# Light Modern Theme (Content Slides)
COLOR_LIGHT_BG = RGBColor(248, 250, 252)    # Slate 50
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_PRIMARY = RGBColor(15, 23, 42)        # Slate 900 (Main headings)
COLOR_SECONDARY = RGBColor(37, 99, 235)     # Blue 600 (Accents/links)
COLOR_BODY = RGBColor(30, 41, 59)           # Slate 800 (Body text)
COLOR_MUTED = RGBColor(100, 116, 139)       # Slate 500 (Subtitles/notes)
COLOR_CARD_BORDER = RGBColor(226, 232, 240) # Slate 200
COLOR_TAG_BG = RGBColor(238, 242, 255)      # Indigo 50
COLOR_TAG_TXT = RGBColor(79, 70, 229)       # Indigo 600
COLOR_EMERALD = RGBColor(16, 185, 129)      # Emerald 500
COLOR_AMBER = RGBColor(217, 119, 6)         # Amber 600
COLOR_HIGHLIGHT_BG = RGBColor(239, 246, 255)# Blue 50
COLOR_CALLOUT_BORDER = RGBColor(147, 197, 253) # Blue 300

FONT_MAIN = "Calibri"
FONT_HEADING = "Arial"

# Slide Dimensions (16:9 Widescreen)
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


def trim_whitespace(img_path, out_path):
    """Trims excess uniform background border around an image."""
    try:
        im = Image.open(img_path)
        bg = Image.new(im.mode, im.size, im.getpixel((0,0)))
        diff = ImageChops.difference(im, bg)
        diff = ImageChops.add(diff, diff, 2.0, -100)
        bbox = diff.getbbox()
        if bbox:
            cropped = im.crop(bbox)
            cropped.save(out_path)
            return out_path
    except Exception as e:
        print(f"Trim failed for {img_path}: {e}")
    return img_path


def add_header(slide, tag_text, title_text, subtitle_text):
    """Adds a standard structured header to a content slide."""
    # Tag pill / category
    tag_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.38), Inches(5.2), Inches(0.3))
    tag_box.fill.solid()
    tag_box.fill.fore_color.rgb = COLOR_TAG_BG
    tag_box.line.color.rgb = COLOR_TAG_TXT
    tag_box.line.width = Pt(1)
    
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = False
    tf_tag.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf_tag.margin_left = Inches(0.15)
    tf_tag.margin_top = Inches(0.0)
    p_tag = tf_tag.paragraphs[0]
    p_tag.text = tag_text.upper()
    p_tag.font.name = FONT_MAIN
    p_tag.font.size = Pt(8.5)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_TAG_TXT
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(11.733), Inches(0.55))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    tf_title.margin_left = Inches(0)
    tf_title.margin_top = Inches(0)
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.name = FONT_HEADING
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = COLOR_PRIMARY
    
    # Subtitle
    sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.28), Inches(11.733), Inches(0.35))
    tf_sub = sub_box.text_frame
    tf_sub.word_wrap = True
    tf_sub.margin_left = Inches(0)
    tf_sub.margin_top = Inches(0)
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = subtitle_text
    p_sub.font.name = FONT_MAIN
    p_sub.font.size = Pt(11.5)
    p_sub.font.color.rgb = COLOR_MUTED


def add_footer(slide, current_slide, total_slides=10):
    """Adds a standard subtle footer bar to slides 2-10."""
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.92), Inches(11.733), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = COLOR_CARD_BORDER
    line.line.fill.background()
    
    foot_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.96), Inches(7.0), Inches(0.3))
    tf = foot_box.text_frame
    tf.margin_left = Inches(0)
    p = tf.paragraphs[0]
    p.text = "Robo-Root: Plant-Root-Inspired Pneumatic Tip-Growing Robot | CUSAT Major Project 2026"
    p.font.name = FONT_MAIN
    p.font.size = Pt(9)
    p.font.color.rgb = COLOR_MUTED
    
    page_box = slide.shapes.add_textbox(Inches(10.533), Inches(6.96), Inches(2.0), Inches(0.3))
    tf_p = page_box.text_frame
    tf_p.margin_right = Inches(0)
    p_p = tf_p.paragraphs[0]
    p_p.text = f"Slide {current_slide} of {total_slides}"
    p_p.alignment = PP_ALIGN.RIGHT
    p_p.font.name = FONT_MAIN
    p_p.font.size = Pt(9)
    p_p.font.bold = True
    p_p.font.color.rgb = COLOR_MUTED


def add_image_fitted(slide, img_path, left, top, max_width, max_height, caption=None):
    """Places an image preserving its exact aspect ratio within bounds and optionally adds a card & caption."""
    if not os.path.exists(img_path):
        print(f"Warning: Image not found: {img_path}")
        return None
    
    im = Image.open(img_path)
    im_w, im_h = im.size
    aspect = im_w / im_h
    
    # Available area for image considering caption height if present
    caption_height = Inches(0.32) if caption else Inches(0)
    avail_h = max_height - caption_height - Inches(0.12)
    avail_w = max_width - Inches(0.2)
    
    if (avail_w / avail_h) > aspect:
        # Height is the constraint
        target_h = avail_h
        target_w = target_h * aspect
    else:
        # Width is the constraint
        target_w = avail_w
        target_h = target_w / aspect
    
    # Background frame
    frame = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, max_width, max_height)
    frame.fill.solid()
    frame.fill.fore_color.rgb = COLOR_WHITE
    frame.line.color.rgb = COLOR_CARD_BORDER
    frame.line.width = Pt(1)
    
    # Center picture horizontally & vertically within available image space
    pos_x = left + (max_width - target_w) / 2
    pos_y = top + Inches(0.06) + (avail_h - target_h) / 2
    
    # Picture
    pic = slide.shapes.add_picture(img_path, pos_x, pos_y, width=target_w, height=target_h)
    
    # Caption
    if caption:
        cap_box = slide.shapes.add_textbox(left, top + max_height - caption_height - Inches(0.04), max_width, caption_height)
        tf_cap = cap_box.text_frame
        tf_cap.word_wrap = True
        tf_cap.margin_top = Inches(0)
        p_cap = tf_cap.paragraphs[0]
        p_cap.text = caption
        p_cap.alignment = PP_ALIGN.CENTER
        p_cap.font.name = FONT_MAIN
        p_cap.font.size = Pt(9.5)
        p_cap.font.bold = True
        p_cap.font.color.rgb = COLOR_MUTED
        
    return pic


def create_presentation():
    # Trim diagram 3 to eliminate excess border
    trim_whitespace("PPT/extracted_images/fig3_robot_overall_mechanical_structure.png", 
                    "PPT/extracted_images/fig3_robot_overall_mechanical_structure_trimmed.png")

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT
    blank_layout = prs.slide_layouts[6]
    
    # =========================================================================
    # SLIDE 1: TITLE SLIDE (Dark Academic Theme)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_DARK_BG
    bg1.line.fill.background()
    
    # Top Institution Banner
    inst_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.5), Inches(11.733), Inches(0.8))
    inst_card.fill.solid()
    inst_card.fill.fore_color.rgb = COLOR_DARK_CARD
    inst_card.line.color.rgb = COLOR_DARK_BORDER
    inst_card.line.width = Pt(1)
    
    # CUSAT Logo on Slide 1
    logo_path = "PPT/extracted_images/cusat_logo.png"
    if os.path.exists(logo_path):
        slide1.shapes.add_picture(logo_path, Inches(1.05), Inches(0.57), height=Inches(0.66))
        
    inst_tf = inst_card.text_frame
    inst_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    inst_tf.margin_left = Inches(1.1)
    p_inst = inst_tf.paragraphs[0]
    p_inst.text = "DIVISION OF COMPUTER SCIENCE AND ENGINEERING | SCHOOL OF ENGINEERING"
    p_inst.font.name = FONT_HEADING
    p_inst.font.size = Pt(11)
    p_inst.font.bold = True
    p_inst.font.color.rgb = COLOR_CYAN_ACCENT
    
    p_inst2 = inst_tf.add_paragraph()
    p_inst2.text = "COCHIN UNIVERSITY OF SCIENCE AND TECHNOLOGY (CUSAT) • OCTOBER 2026"
    p_inst2.font.name = FONT_MAIN
    p_inst2.font.size = Pt(9.5)
    p_inst2.font.color.rgb = COLOR_DARK_MUTED
    
    # Main Title Card
    main_title_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.48), Inches(11.733), Inches(2.2))
    main_title_card.fill.solid()
    main_title_card.fill.fore_color.rgb = COLOR_DARK_CARD
    main_title_card.line.color.rgb = COLOR_DARK_BORDER
    main_title_card.line.width = Pt(1)
    
    tf_main = main_title_card.text_frame
    tf_main.margin_left = Inches(0.5)
    tf_main.margin_top = Inches(0.28)
    
    p_tag = tf_main.paragraphs[0]
    p_tag.text = "MAJOR PROJECT DESIGN & ARCHITECTURE PRESENTATION"
    p_tag.font.name = FONT_MAIN
    p_tag.font.size = Pt(10)
    p_tag.font.bold = True
    p_tag.font.color.rgb = COLOR_BLUE_ACCENT
    
    p_maintitle = tf_main.add_paragraph()
    p_maintitle.text = "Robo-Root: Plant-Root-Inspired Soft Growing Robot"
    p_maintitle.font.name = FONT_HEADING
    p_maintitle.font.size = Pt(28)
    p_maintitle.font.bold = True
    p_maintitle.font.color.rgb = COLOR_DARK_TEXT
    
    p_mainsub = tf_main.add_paragraph()
    p_mainsub.text = "Pneumatic Tip Eversion, Base-Mounted Four-Tendon Steering & Vision-Assisted Autonomous Navigation"
    p_mainsub.font.name = FONT_MAIN
    p_mainsub.font.size = Pt(13)
    p_mainsub.font.color.rgb = COLOR_CYAN_ACCENT
    
    # 4 Presenter Cards
    team = [
        ("Mihal Palakkal", "20223071"),
        ("Rose Maria", "20223085"),
        ("Ruvais P", "20223087"),
        ("Mohamed Jaseeb P J", "20223120")
    ]
    card_w = Inches(2.75)
    gap = Inches(0.24)
    start_x = Inches(0.8)
    for i, (name, reg) in enumerate(team):
        cx = start_x + i * (card_w + gap)
        scard = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(3.85), card_w, Inches(1.3))
        scard.fill.solid()
        scard.fill.fore_color.rgb = COLOR_DARK_CARD
        scard.line.color.rgb = COLOR_DARK_BORDER
        scard.line.width = Pt(1)
        
        stf = scard.text_frame
        stf.margin_top = Inches(0.18)
        stf.margin_left = Inches(0.25)
        
        p1 = stf.paragraphs[0]
        p1.text = "PROJECT MEMBER"
        p1.font.name = FONT_MAIN
        p1.font.size = Pt(8.5)
        p1.font.bold = True
        p1.font.color.rgb = COLOR_BLUE_ACCENT
        
        p2 = stf.add_paragraph()
        p2.text = name
        p2.font.name = FONT_HEADING
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = COLOR_DARK_TEXT
        
        p3 = stf.add_paragraph()
        p3.text = f"Reg No: {reg}"
        p3.font.name = FONT_MAIN
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = COLOR_DARK_MUTED

    # Mentorship & Guide Card (No text overlap)
    guide_card = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.32), Inches(11.733), Inches(1.5))
    guide_card.fill.solid()
    guide_card.fill.fore_color.rgb = COLOR_DARK_CARD
    guide_card.line.color.rgb = COLOR_DARK_BORDER
    guide_card.line.width = Pt(1)
    
    # Sub-header box for supervision
    advisory_lbl = slide1.shapes.add_textbox(Inches(1.1), Inches(5.42), Inches(11.0), Inches(0.3))
    tf_adv = advisory_lbl.text_frame
    tf_adv.margin_left = Inches(0)
    tf_adv.margin_top = Inches(0)
    p_adv = tf_adv.paragraphs[0]
    p_adv.text = "ACADEMIC ADVISORY & PROJECT SUPERVISION"
    p_adv.font.name = FONT_MAIN
    p_adv.font.size = Pt(9)
    p_adv.font.bold = True
    p_adv.font.color.rgb = COLOR_CYAN_ACCENT
    
    # 3 distinct columns within the card
    coords = [
        ("Ms. Amrutha S Nair", "Project Guide", "Assistant Professor, Dept. of CSE", Inches(1.1)),
        ("Dr. Sheena S", "Project Coordinator", "Associate Professor, Dept. of CSE", Inches(5.0)),
        ("Mr. Damodaran V", "Head of Division", "Division of Computer Science & Engg", Inches(8.8))
    ]
    for gname, role, desig, gx in coords:
        gbox = slide1.shapes.add_textbox(gx, Inches(5.75), Inches(3.6), Inches(0.95))
        gtf_sub = gbox.text_frame
        gtf_sub.margin_left = Inches(0)
        gtf_sub.margin_top = Inches(0)
        
        pg1 = gtf_sub.paragraphs[0]
        pg1.text = gname
        pg1.font.name = FONT_HEADING
        pg1.font.size = Pt(11)
        pg1.font.bold = True
        pg1.font.color.rgb = COLOR_DARK_TEXT
        
        pg2 = gtf_sub.add_paragraph()
        pg2.text = role
        pg2.font.name = FONT_MAIN
        pg2.font.size = Pt(9.5)
        pg2.font.color.rgb = COLOR_BLUE_ACCENT
        
        pg3 = gtf_sub.add_paragraph()
        pg3.text = desig
        pg3.font.name = FONT_MAIN
        pg3.font.size = Pt(8.5)
        pg3.font.color.rgb = COLOR_DARK_MUTED

    # =========================================================================
    # SLIDE 2: PROJECT OVERVIEW & MOTIVATION
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Background & Motivation | Chapter 01", 
               "Bio-Inspired Soft Robotics: Why Tip-Growing Systems?", 
               "Overcoming the severe physical constraints of rigid locomotion in complex and confined subterranean environments.")
    add_footer(slide2, 2)
    
    # Left Card: Traditional Rigid Robots Limitations
    card2_l = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.7), Inches(4.8))
    card2_l.fill.solid()
    card2_l.fill.fore_color.rgb = COLOR_WHITE
    card2_l.line.color.rgb = COLOR_CARD_BORDER
    card2_l.line.width = Pt(1)
    
    tf2_l = card2_l.text_frame
    tf2_l.margin_left = Inches(0.35)
    tf2_l.margin_right = Inches(0.35)
    tf2_l.margin_top = Inches(0.35)
    
    p = tf2_l.paragraphs[0]
    p.text = "LIMITATIONS OF CONVENTIONAL RIGID ROBOTS"
    p.font.name = FONT_HEADING
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_AMBER
    
    bullets2_l = [
        ("Whole-Body Motion & Friction: ", "Wheeled, tracked, or legged robots must drag or propel their entire body mass forward through the terrain, creating immense frictional drag in narrow passages."),
        ("Wedging & Geometric Trapping: ", "Rigid structures cannot deform dynamically through tortuous or collapsing cavities, leading to entrapment in disaster rubble or tight pipes."),
        ("Subterranean Soil Resistance: ", "Penetrating soil or sand with rigid shafts incurs exponential penetration resistance that scales drastically with penetration depth."),
        ("High In-Tip Actuator Mass: ", "Locating heavy drive motors or articulating joints near the front compromises balance and severely limits miniaturization.")
    ]
    for bold_prefix, text in bullets2_l:
        p = tf2_l.add_paragraph()
        p.space_before = Pt(10)
        run_b = p.add_run()
        run_b.text = "• " + bold_prefix
        run_b.font.bold = True
        run_b.font.size = Pt(10)
        run_b.font.color.rgb = COLOR_PRIMARY
        run_t = p.add_run()
        run_t.text = text
        run_t.font.size = Pt(9.5)
        run_t.font.color.rgb = COLOR_BODY

    # Right Card: The Robo-Root Apical Growth Advantage
    card2_r = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.85), Inches(5.7), Inches(4.8))
    card2_r.fill.solid()
    card2_r.fill.fore_color.rgb = COLOR_WHITE
    card2_r.line.color.rgb = COLOR_CARD_BORDER
    card2_r.line.width = Pt(1)
    
    tf2_r = card2_r.text_frame
    tf2_r.margin_left = Inches(0.35)
    tf2_r.margin_right = Inches(0.35)
    tf2_r.margin_top = Inches(0.35)
    
    p = tf2_r.paragraphs[0]
    p.text = "THE ROBO-ROOT ADVANTAGE: APICAL TIP EVERSION"
    p.font.name = FONT_HEADING
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_SECONDARY
    
    bullets2_r = [
        ("Apical Extension (Zero Wall-Sliding Friction): ", "Inspired by biological plant roots, material extends strictly at the tip through pneumatic eversion. The deployed body remains stationary relative to surrounding walls."),
        ("Passage Through Unstructured Voids: ", "Seamlessly adapts to tortuous pathways, squeezing through apertures narrower than the unpressurized tube diameter."),
        ("Internal Working Lumen Conduit: ", "The hollow central lumen provides a protected interior pathway for electrical wiring, sensor payloads, and air supply without environmental exposure."),
        ("High Expansion-to-Mass Ratio: ", "Lightweight polymer membranes can achieve dramatic extensions (several meters) while base station bears the actuator weight.")
    ]
    for bold_prefix, text in bullets2_r:
        p = tf2_r.add_paragraph()
        p.space_before = Pt(8)
        run_b = p.add_run()
        run_b.text = "✔ " + bold_prefix
        run_b.font.bold = True
        run_b.font.size = Pt(10)
        run_b.font.color.rgb = COLOR_EMERALD
        run_t = p.add_run()
        run_t.text = text
        run_t.font.size = Pt(9.5)
        run_t.font.color.rgb = COLOR_BODY

    # Bottom Callout Box in Right Card
    callout2 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.05), Inches(5.35), Inches(5.2), Inches(1.15))
    callout2.fill.solid()
    callout2.fill.fore_color.rgb = COLOR_HIGHLIGHT_BG
    callout2.line.color.rgb = COLOR_CALLOUT_BORDER
    callout2.line.width = Pt(1)
    tf_c2 = callout2.text_frame
    tf_c2.margin_left = Inches(0.2)
    tf_c2.margin_top = Inches(0.12)
    p = tf_c2.paragraphs[0]
    p.text = "TARGET APPLICATIONS:"
    p.font.bold = True
    p.font.size = Pt(9.5)
    p.font.color.rgb = COLOR_SECONDARY
    p2 = tf_c2.add_paragraph()
    p2.text = "Urban Search & Rescue in collapsed structures, industrial pipeline and duct inspection, underground sensor deployment, and non-destructive soil probing."
    p2.font.size = Pt(9)
    p2.font.color.rgb = COLOR_BODY

    # =========================================================================
    # SLIDE 3: PROBLEM STATEMENT & ENGINEERING CHALLENGES
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Problem Statement | Chapter 02", 
               "Core Engineering & Scientific Challenges", 
               "Balancing pneumatic growth dynamics, tendon-based mechanical steering, and real-time vision perception.")
    add_footer(slide3, 3)
    
    challenges = [
        ("1. DIRECTIONAL STEERING CONTROL", 
         COLOR_SECONDARY,
         "• Uncontrolled Tip Eversion: Pressurized pneumatic tubes naturally extend along the path of least resistance or in a straight line.\n\n"
         "• Planar & Spatial Bending: Imparting deliberate 2D/3D angular deflection requires active, coordinated steering authority without inducing structural buckling or tube kinking.\n\n"
         "• Critical Need: A decoupled 4-tendon antagonistic steering mechanism operating at the tip."),
        
        ("2. TENDON ROUTING & CABLE PAYOUT", 
         COLOR_AMBER,
         "• Growth vs. Cable Tension: As the robot extends from centimeters to meters, steering cables must deploy continuously.\n\n"
         "• Slack & Friction Management: Loose tendons cause steering deadband and backlash, while overtight tendons prevent eversion growth.\n\n"
         "• Critical Need: Base-mounted motorized cable spools that dynamically release tendon slack synchronous with growth."),
        
        ("3. TIP PERCEPTION IN CONFINED SPACES", 
         COLOR_PRIMARY,
         "• Hostile Visual Environments: Confined boreholes and rubble exhibit complete darkness, dust, lens occlusions, and severe glare.\n\n"
         "• Computational Bottlenecks: Real-time image processing cannot rely on heavy GPU hardware mounted on the fragile soft tip.\n\n"
         "• Critical Need: Micro-camera feed streaming to a host ROS 2 machine running optimized YOLO inference."),
        
        ("4. MULTIMODAL SYSTEM SYNCHRONIZATION", 
         COLOR_EMERALD,
         "• Actuator-Sensor Decoupling: High pneumatic pressure, servo pulse modulation, and image capture must be synchronized.\n\n"
         "• Closed-Loop Control Latency: Steering corrections must occur at speeds compatible with pneumatic forward advancement.\n\n"
         "• Critical Need: A modular ROS 2 architecture establishing deterministic communication across all subsystems.")
    ]
    
    col_w = Inches(2.75)
    col_gap = Inches(0.24)
    start_x = Inches(0.8)
    
    for i, (title, accent_col, body_txt) in enumerate(challenges):
        cx = start_x + i * (col_w + col_gap)
        card = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.85), col_w, Inches(4.8))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_WHITE
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)
        
        # Color bar top accent
        top_bar = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, Inches(1.85), col_w, Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = accent_col
        top_bar.line.fill.background()
        
        ctf = card.text_frame
        ctf.margin_top = Inches(0.25)
        ctf.margin_left = Inches(0.2)
        ctf.margin_right = Inches(0.2)
        
        p = ctf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = accent_col
        
        p_body = ctf.add_paragraph()
        p_body.space_before = Pt(12)
        p_body.text = body_txt
        p_body.font.name = FONT_MAIN
        p_body.font.size = Pt(9.5)
        p_body.font.color.rgb = COLOR_BODY

    # =========================================================================
    # SLIDE 4: PROJECT OBJECTIVES & SYSTEM SCOPE
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Project Objectives & Scope | Chapters 01 & 02", 
               "Systematic Objectives & Technical Scope Boundaries", 
               "Defining the engineering milestones and practical boundaries of the laboratory prototype.")
    add_footer(slide4, 4)
    
    # Left Card: Specific Objectives
    card4_l = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.85), Inches(6.5), Inches(4.8))
    card4_l.fill.solid()
    card4_l.fill.fore_color.rgb = COLOR_WHITE
    card4_l.line.color.rgb = COLOR_CARD_BORDER
    card4_l.line.width = Pt(1)
    
    tf4_l = card4_l.text_frame
    tf4_l.margin_left = Inches(0.35)
    tf4_l.margin_right = Inches(0.3)
    tf4_l.margin_top = Inches(0.3)
    
    p = tf4_l.paragraphs[0]
    p.text = "PRIMARY TECHNICAL OBJECTIVES"
    p.font.name = FONT_HEADING
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_SECONDARY
    
    objectives = [
        ("Pneumatic Tip Eversion: ", "Design an airtight, flexible tubular mechanism capable of continuous forward growth driven by regulated pneumatic pressure."),
        ("Four-Tendon Steering: ", "Construct a 4-tendon orthogonal arrangement around the flexible tube for active planar tip steering in the X-Y horizontal plane."),
        ("Base-Station Actuation: ", "Mount four high-torque servo motors at the stationary base to minimize tip payload and enable active tendon spooling."),
        ("Vision Integration: ", "Equip the advancing tip with a micro-camera for real-time forward environmental sensing and obstacle detection."),
        ("YOLO Perception Pipeline: ", "Integrate pretrained YOLO models via OpenCV to detect objects and obstacles without needing massive prior training sets."),
        ("ROS 2 Control Architecture: ", "Establish a modular ROS 2 framework linking image acquisition, path planning, kinematics, and ESP32 low-level firmware.")
    ]
    for i, (pref, desc) in enumerate(objectives, 1):
        p = tf4_l.add_paragraph()
        p.space_before = Pt(8)
        run_n = p.add_run()
        run_n.text = f"{i}. {pref}"
        run_n.font.bold = True
        run_n.font.size = Pt(10)
        run_n.font.color.rgb = COLOR_PRIMARY
        run_d = p.add_run()
        run_d.text = desc
        run_d.font.size = Pt(9.5)
        run_d.font.color.rgb = COLOR_BODY

    # Right Card: Project Scope & Technical Boundaries
    card4_r = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(1.85), Inches(5.033), Inches(4.8))
    card4_r.fill.solid()
    card4_r.fill.fore_color.rgb = COLOR_WHITE
    card4_r.line.color.rgb = COLOR_CARD_BORDER
    card4_r.line.width = Pt(1)
    
    tf4_r = card4_r.text_frame
    tf4_r.margin_left = Inches(0.35)
    tf4_r.margin_right = Inches(0.3)
    tf4_r.margin_top = Inches(0.3)
    
    p = tf4_r.paragraphs[0]
    p.text = "SCOPE BOUNDARIES & ASSUMPTIONS"
    p.font.name = FONT_HEADING
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY
    
    scope_points = [
        ("Base-Mounted Servos: ", "Servos remain stationary at the base rather than moving with the tip, drastically reducing moving inertia but requiring robust cable payout management."),
        ("Pretrained Model Utilization: ", "Leverages established YOLO weights for standard object classes, focusing on the system integration pipeline rather than de novo neural network training."),
        ("Perception Outputs: ", "Visual detection yields 2D bounding boxes and class IDs; metric depth estimation is handled modularly through geometric or additional depth sensors."),
        ("Empirical Validation Focus: ", "Prototype designed for laboratory bench testing to evaluate growth rate, steering deflection angle, and deployment-retrieval cycles.")
    ]
    for pref, desc in scope_points:
        p = tf4_r.add_paragraph()
        p.space_before = Pt(10)
        run_p = p.add_run()
        run_p.text = "• " + pref
        run_p.font.bold = True
        run_p.font.size = Pt(10)
        run_p.font.color.rgb = COLOR_SECONDARY
        run_d = p.add_run()
        run_d.text = desc
        run_d.font.size = Pt(9.5)
        run_d.font.color.rgb = COLOR_BODY

    # =========================================================================
    # SLIDE 5: SYSTEM WORKFLOW & ROS 2 ARCHITECTURE
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "System Workflow & Software Architecture | Chapter 05", 
               "End-to-End System Workflow & ROS 2 Framework", 
               "Modular software nodes communicating across perception, path planning, and embedded actuation.")
    add_footer(slide5, 5)
    
    # Left Card: 9-Stage Execution Workflow
    card5_l = slide5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.4), Inches(4.8))
    card5_l.fill.solid()
    card5_l.fill.fore_color.rgb = COLOR_WHITE
    card5_l.line.color.rgb = COLOR_CARD_BORDER
    card5_l.line.width = Pt(1)
    
    tf5_l = card5_l.text_frame
    tf5_l.margin_left = Inches(0.3)
    tf5_l.margin_right = Inches(0.25)
    tf5_l.margin_top = Inches(0.25)
    
    p = tf5_l.paragraphs[0]
    p.text = "INTEGRATED SYSTEM WORKFLOW LOOP"
    p.font.name = FONT_HEADING
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_SECONDARY
    
    stages = [
        ("1. Initialization: ", "ROS 2 nodes, camera stream, ESP32 UART, & pneumatic manifold initialize."),
        ("2. Pneumatic Growth: ", "Regulated air enters chamber; tubular membrane begins tip eversion."),
        ("3. Frame Acquisition: ", "Tip camera captures visual frames ahead of advancing robot."),
        ("4. YOLO Inference: ", "OpenCV resizes/letterboxes frames; YOLO outputs bounding boxes."),
        ("5. Path Planning: ", "ROS 2 node determines obstacle-free trajectory setpoint."),
        ("6. Steering Kinematics: ", "Controller converts target angle into 4 differential tendon lengths."),
        ("7. Servo Actuation: ", "ESP32 generates PWM pulses for 4 base-mounted servo motors."),
        ("8. Cable Payout: ", "Motorized drums release tendon slack synchronous with tube growth."),
        ("9. Pressure & Safety: ", "Pressure transducer ensures line pressure stays within safe thresholds.")
    ]
    for pref, text in stages:
        p = tf5_l.add_paragraph()
        p.space_before = Pt(4.5)
        run_b = p.add_run()
        run_b.text = pref
        run_b.font.bold = True
        run_b.font.size = Pt(9.5)
        run_b.font.color.rgb = COLOR_PRIMARY
        run_t = p.add_run()
        run_t.text = text
        run_t.font.size = Pt(9)
        run_t.font.color.rgb = COLOR_BODY

    # Right Frame: Embedded Architecture Figure
    img_arch = "PPT/extracted_images/fig2_ros2_system_architecture.png"
    add_image_fitted(slide5, img_arch, Inches(6.4), Inches(1.85), Inches(6.133), Inches(4.8), 
                     caption="Figure 5.1: ROS 2 Software Architecture & Hardware Interface")

    # =========================================================================
    # SLIDE 6: MECHANICAL DESIGN & TIP-EVERSION MECHANISM
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Modular Structure: Mechanical | Chapter 06", 
               "Mechanical Design & Tip-Eversion Mechanism", 
               "Pneumatic inversion dynamics, base-station structural layout, and internal lumen routing.")
    add_footer(slide6, 6)
    
    # Left Card: Detailed Mechanical Analysis
    card6_l = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.2), Inches(4.8))
    card6_l.fill.solid()
    card6_l.fill.fore_color.rgb = COLOR_WHITE
    card6_l.line.color.rgb = COLOR_CARD_BORDER
    card6_l.line.width = Pt(1)
    
    tf6_l = card6_l.text_frame
    tf6_l.margin_left = Inches(0.3)
    tf6_l.margin_right = Inches(0.3)
    tf6_l.margin_top = Inches(0.3)
    
    p = tf6_l.paragraphs[0]
    p.text = "MECHANICAL DESIGN PRINCIPLES"
    p.font.name = FONT_HEADING
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_SECONDARY
    
    mech_points = [
        ("Tip Eversion Dynamics: ", "The robot body consists of a flexible, airtight membrane folded inwards upon itself. Pneumatic pressure applied internally forces new material to unfold exclusively at the tip, pushing the robot forward without sliding friction."),
        ("Base Station Assembly: ", "Houses the primary structural weight: the air supply connection, pressure regulator, four servo actuators, cable payout drums, and control electronics. Isolates delicate parts from hazardous environments."),
        ("Growing Tip Head Assembly: ", "A lightweight, rigid circular ring mounted at the inversion point. Acts as the anchor for the 4 steering tendons, mounts the micro-camera module, and ensures circular aperture maintenance."),
        ("Internal Working Lumen: ", "The hollow central core forms a protected conduit that carries tendon cables, camera wiring, and pneumatic hoses without exposing them to exterior dirt, sand, or debris.")
    ]
    for pref, text in mech_points:
        p = tf6_l.add_paragraph()
        p.space_before = Pt(8)
        run_b = p.add_run()
        run_b.text = "• " + pref
        run_b.font.bold = True
        run_b.font.size = Pt(9.5)
        run_b.font.color.rgb = COLOR_PRIMARY
        run_t = p.add_run()
        run_t.text = text
        run_t.font.size = Pt(9)
        run_t.font.color.rgb = COLOR_BODY

    # Right Frame: Trimmed Mechanical Diagram (fills space much better)
    img_mech = "PPT/extracted_images/fig3_robot_overall_mechanical_structure_trimmed.png"
    add_image_fitted(slide6, img_mech, Inches(6.2), Inches(1.85), Inches(6.333), Inches(4.8), 
                     caption="Figure 6.1: Overall Mechanical Structure of the Pneumatic Tip-Growing Robot")

    # =========================================================================
    # SLIDE 7: FOUR-TENDON STEERING & SERVO DRUM ASSEMBLY
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Modular Structure: Steering & Actuation | Chapter 06", 
               "Four-Tendon Steering & Base-Mounted Servo Mechanism", 
               "Decoupling actuation weight from the tip through base-mounted antagonistic cable spools.")
    add_footer(slide7, 7)
    
    # Left Card: Detailed Steering & Spooling Engineering
    card7_l = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.2), Inches(4.8))
    card7_l.fill.solid()
    card7_l.fill.fore_color.rgb = COLOR_WHITE
    card7_l.line.color.rgb = COLOR_CARD_BORDER
    card7_l.line.width = Pt(1)
    
    tf7_l = card7_l.text_frame
    tf7_l.margin_left = Inches(0.3)
    tf7_l.margin_right = Inches(0.3)
    tf7_l.margin_top = Inches(0.3)
    
    p = tf7_l.paragraphs[0]
    p.text = "DIRECTIONAL STEERING & ACTUATION"
    p.font.name = FONT_HEADING
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_SECONDARY
    
    steer_points = [
        ("Four-Tendon Orthogonal Arrangement: ", "Four axial cables are routed circumferentially at 90° intervals (North, East, South, West). Antagonistic tensioning allows omnidirectional 2D bending in the X-Y plane."),
        ("Differential Bending Moments: ", "Pulling tendon X+ while releasing tendon X- creates a bending curvature in the +X direction without altering overall growth speed or internal pressure."),
        ("Base-Mounted Actuator Spools: ", "Mounting the four high-torque servo motors at the stationary base keeps tip inertia at zero, preventing front-heavy tipping and allowing high steering agility."),
        ("Dynamic Cable Payout Mechanics: ", "Each servo drives a precision cable drum with guide channels. As the robot grows, tendons pay out in lockstep with the advancing membrane, maintaining consistent tension and avoiding tangling.")
    ]
    for pref, text in steer_points:
        p = tf7_l.add_paragraph()
        p.space_before = Pt(8)
        run_b = p.add_run()
        run_b.text = "• " + pref
        run_b.font.bold = True
        run_b.font.size = Pt(9.5)
        run_b.font.color.rgb = COLOR_PRIMARY
        run_t = p.add_run()
        run_t.text = text
        run_t.font.size = Pt(9)
        run_t.font.color.rgb = COLOR_BODY

    # Right: Two Stacked Figures
    # Top Right: Figure 6.2 Directional Steering
    img_steer = "PPT/extracted_images/fig4_four_tendon_directional_steering.png"
    add_image_fitted(slide7, img_steer, Inches(6.2), Inches(1.85), Inches(6.333), Inches(2.32),
                     caption="Figure 6.2: Directional Steering via Four-Tendon Control")
    
    # Bottom Right: Figure 6.3 Servo Spool Assembly
    img_spool = "PPT/extracted_images/fig5_servo_cable_drum_assembly.png"
    add_image_fitted(slide7, img_spool, Inches(6.2), Inches(4.32), Inches(6.333), Inches(2.33),
                     caption="Figure 6.3: Base-Mounted Servo Motors & Cable-Drum Assembly")

    # =========================================================================
    # SLIDE 8: VISION PERCEPTION & PREPROCESSING PIPELINE (Refined Wide Layout)
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "Perception System: YOLO Pipeline | Chapter 04", 
               "Tip-Mounted Computer Vision & YOLO Inference Pipeline", 
               "From raw camera capture to real-time object classification, bounding boxes, and path planner inputs.")
    add_footer(slide8, 8)
    
    # Top 3 Concept Cards across Slide 8
    pipe_cards = [
        ("1. CAMERA ACQUISITION & CONVERSION", COLOR_SECONDARY,
         "• Micro-Camera Tip Capture: High-framerate video captures terrain ahead of growing apex.\n"
         "• ROS 2 Transport & CvBridge: Converts raw sensor buffers into standard OpenCV BGR arrays."),
        
        ("2. PREPROCESSING & LETTERBOXING", COLOR_AMBER,
         "• Aspect-Preserving Letterboxing: Resizes frames to 640x640 with symmetric padding.\n"
         "• Pixel Normalization: Scales RGB values [0, 255] -> [0.0, 1.0] for model tensor batching."),
        
        ("3. YOLO INFERENCE & INTEGRATION", COLOR_EMERALD,
         "• Pretrained Neural Inference: Multi-scale convolutional feature extraction & class prediction.\n"
         "• NMS & Trajectory Node: Non-Maximum Suppression sends obstacle coordinates to ROS 2.")
    ]
    top_c_w = Inches(3.75)
    top_c_gap = Inches(0.24)
    for i, (title, accent, body) in enumerate(pipe_cards):
        cx = Inches(0.8) + i * (top_c_w + top_c_gap)
        c = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.8), top_c_w, Inches(1.65))
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_WHITE
        c.line.color.rgb = COLOR_CARD_BORDER
        c.line.width = Pt(1)
        
        top_bar = slide8.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, Inches(1.8), top_c_w, Inches(0.06))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = accent
        top_bar.line.fill.background()
        
        ctf = c.text_frame
        ctf.margin_top = Inches(0.16)
        ctf.margin_left = Inches(0.2)
        ctf.margin_right = Inches(0.2)
        p = ctf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = accent
        
        pb = ctf.add_paragraph()
        pb.space_before = Pt(4)
        pb.text = body
        pb.font.name = FONT_MAIN
        pb.font.size = Pt(8.5)
        pb.font.color.rgb = COLOR_BODY

    # Bottom Full-Width Frame for Figure 4.1
    img_pipe = "PPT/extracted_images/fig1_image_processing_pipeline.png"
    add_image_fitted(slide8, img_pipe, Inches(0.8), Inches(3.6), Inches(11.733), Inches(3.05),
                     caption="Figure 4.1: End-to-End Image Processing & Pretrained YOLO Object Detection Pipeline")

    # =========================================================================
    # SLIDE 9: ELECTRONICS, EMBEDDED CONTROL & WIRING
    # =========================================================================
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "Embedded Hardware: Electronics | Chapters 06 & 07", 
               "Electronics, Power Distribution & Interfacing", 
               "ESP32 embedded microcontroller bridging high-level ROS 2 commands with physical actuators and sensors.")
    add_footer(slide9, 9)
    
    # Left Card: Detailed Electronics Architecture
    card9_l = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.2), Inches(4.8))
    card9_l.fill.solid()
    card9_l.fill.fore_color.rgb = COLOR_WHITE
    card9_l.line.color.rgb = COLOR_CARD_BORDER
    card9_l.line.width = Pt(1)
    
    tf9_l = card9_l.text_frame
    tf9_l.margin_left = Inches(0.3)
    tf9_l.margin_right = Inches(0.3)
    tf9_l.margin_top = Inches(0.3)
    
    p = tf9_l.paragraphs[0]
    p.text = "HARDWARE & ELECTRICAL INTEGRATION"
    p.font.name = FONT_HEADING
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_SECONDARY
    
    elec_points = [
        ("ESP32 Dual-Core Microcontroller: ", "Manages low-latency real-time actuator control. Receives high-level trajectory commands from host PC via USB/UART and generates high-resolution PWM signals."),
        ("Dedicated High-Current Servo Bus: ", "Steering requires significant torque; four servos are powered by an isolated 5V/6V high-amperage power supply with common ground isolation to prevent logic brownouts."),
        ("Pneumatic Solenoid Valve Driver: ", "Transistor/MOSFET switching circuit controls the air flow solenoid valve, modulating pneumatic growth rate and emergency shutoff."),
        ("Sensory & Safety Interfacing: ", "Integrated pressure sensor continuously transmits pneumatic line pressure to prevent tube over-inflation; supports optional tip IMU sensor for orientation.")
    ]
    for pref, text in elec_points:
        p = tf9_l.add_paragraph()
        p.space_before = Pt(8)
        run_b = p.add_run()
        run_b.text = "• " + pref
        run_b.font.bold = True
        run_b.font.size = Pt(9.5)
        run_b.font.color.rgb = COLOR_PRIMARY
        run_t = p.add_run()
        run_t.text = text
        run_t.font.size = Pt(9)
        run_t.font.color.rgb = COLOR_BODY

    # Right Frame: Wiring Diagram
    img_elec = "PPT/extracted_images/fig6_electronics_wiring_diagram.png"
    add_image_fitted(slide9, img_elec, Inches(6.2), Inches(1.85), Inches(6.333), Inches(4.8), 
                     caption="Figure 6.4: Electronics and Wiring Diagram of the Growing Robot")

    # =========================================================================
    # SLIDE 10: CONCLUSION, EVALUATION ROADMAP & FUTURE SCOPE
    # =========================================================================
    slide10 = prs.slides.add_slide(blank_layout)
    add_header(slide10, "Conclusion & Future Work | Chapter 08", 
               "Project Summary, Evaluation Benchmarks & Future Scope", 
               "Synthesizing key design achievements and outlining empirical test plans for the prototype.")
    add_footer(slide10, 10)
    
    cols = [
        ("PROJECT DESIGN ACHIEVEMENTS", COLOR_SECONDARY, [
            ("Integrated Soft-Growing Architecture: ", "Successfully established a complete design paradigm combining pneumatic tip eversion with 4-tendon steering."),
            ("Actuator Weight Decoupling: ", "Positioned high-torque servos at the base station, eliminating moving tip mass and maximizing navigation agility."),
            ("Vision-Integrated Control: ", "Formulated an end-to-end perception pipeline leveraging pretrained YOLO for real-time obstacle detection."),
            ("Modular ROS 2 Architecture: ", "Created a decoupled framework where mechanical, pneumatic, electronic, and software components can be tested independently.")
        ]),
        ("PLANNED EVALUATION BENCHMARKS", COLOR_AMBER, [
            ("Pneumatic Growth Rate: ", "Measure deployment velocity (cm/s) as a function of air pressure (kPa) and tube material friction."),
            ("Steering Precision & Deflection: ", "Quantify maximum angular deflection range (0°–90°), hysteresis, and repeatable steering accuracy."),
            ("Perception Latency: ", "Benchmark YOLO frame rate (FPS), detection mAP, and classification confidence under dim lighting."),
            ("Deployment & Retrieval Cycles: ", "Test complete forward growth and tendon-assisted retraction without entanglement or membrane rupture.")
        ]),
        ("FUTURE ENHANCEMENTS", COLOR_EMERALD, [
            ("Closed-Loop Visual Servoing: ", "Incorporate stereoscopic / depth camera sensing for exact metric 3D obstacle distance measurement."),
            ("Tension Feedback Control: ", "Integrate inline miniature load cells on cable drums for closed-loop tension regulation."),
            ("Granular Media Burrowing: ", "Evaluate penetration capabilities in noncohesive dry granular media (sand, loose gravel)."),
            ("Autonomous Path Replanning: ", "Implement dynamic obstacle avoidance and SLAM in GPS-denied confined cavities.")
        ])
    ]
    
    c_w = Inches(3.75)
    c_gap = Inches(0.24)
    start_x = Inches(0.8)
    
    for i, (title, accent, bullets) in enumerate(cols):
        cx = start_x + i * (c_w + c_gap)
        card = slide10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, Inches(1.85), c_w, Inches(4.8))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_WHITE
        card.line.color.rgb = COLOR_CARD_BORDER
        card.line.width = Pt(1)
        
        # Color bar top accent
        top_bar = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, cx, Inches(1.85), c_w, Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = accent
        top_bar.line.fill.background()
        
        ctf = card.text_frame
        ctf.margin_top = Inches(0.25)
        ctf.margin_left = Inches(0.22)
        ctf.margin_right = Inches(0.22)
        
        p = ctf.paragraphs[0]
        p.text = title
        p.font.name = FONT_HEADING
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = accent
        
        for pref, desc in bullets:
            p_b = ctf.add_paragraph()
            p_b.space_before = Pt(8)
            run_b = p_b.add_run()
            run_b.text = "• " + pref
            run_b.font.bold = True
            run_b.font.size = Pt(9.5)
            run_b.font.color.rgb = COLOR_PRIMARY
            run_d = p_b.add_run()
            run_d.text = desc
            run_d.font.size = Pt(9)
            run_d.font.color.rgb = COLOR_BODY

    # Output Presentation
    os.makedirs("PPT", exist_ok=True)
    output_path = "PPT/Robo_Root_Project_Presentation.pptx"
    prs.save(output_path)
    print(f"Presentation regenerated successfully: {output_path}")

if __name__ == "__main__":
    create_presentation()
