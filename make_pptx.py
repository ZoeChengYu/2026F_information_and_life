import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def build_presentation(output_path):
    prs = Presentation()
    # 16:9 widescreen: 13.333 x 7.5 inches
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Professional Theme Color Palette
    C_DARK_BG = RGBColor(15, 23, 42)          # Slate 900
    C_DARK_CARD = RGBColor(30, 41, 59)        # Slate 800
    C_LIGHT_BG = RGBColor(248, 250, 252)      # Slate 50
    C_WHITE = RGBColor(255, 255, 255)
    C_CARD_BG = RGBColor(255, 255, 255)
    C_CARD_BORDER = RGBColor(226, 232, 240)   # Slate 200
    C_PRIMARY = RGBColor(37, 99, 235)         # Royal Blue
    C_PRIMARY_LIGHT = RGBColor(239, 246, 255) # Blue 50
    C_TEAL = RGBColor(13, 148, 136)           # Teal
    C_TEAL_LIGHT = RGBColor(240, 253, 250)    # Teal 50
    C_AMBER = RGBColor(217, 119, 6)           # Amber
    C_AMBER_LIGHT = RGBColor(254, 243, 199)   # Amber 50
    C_ROSE = RGBColor(225, 29, 72)            # Rose Red
    C_ROSE_LIGHT = RGBColor(255, 241, 242)    # Rose 50
    C_TEXT_DARK = RGBColor(15, 23, 42)        # Slate 900
    C_TEXT_MUTED = RGBColor(100, 116, 139)    # Slate 500
    C_TEXT_LIGHT = RGBColor(248, 250, 252)

    FONT_MAIN = "Microsoft JhengHei"

    def set_slide_background(slide, color):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.color.rgb = color
        return bg

    def add_header(slide, badge_text, title_text, takeaway_sentence=""):
        # Category Badge Pill
        badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.42), Inches(2.8), Inches(0.32))
        badge.fill.solid()
        badge.fill.fore_color.rgb = C_PRIMARY_LIGHT
        badge.line.color.rgb = RGBColor(191, 219, 254)
        tf_b = badge.text_frame
        tf_b.word_wrap = True
        tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
        p_b = tf_b.paragraphs[0]
        p_b.alignment = PP_ALIGN.CENTER
        run_b = p_b.add_run()
        run_b.text = badge_text
        run_b.font.name = FONT_MAIN
        run_b.font.size = Pt(10.5)
        run_b.font.bold = True
        run_b.font.color.rgb = C_PRIMARY

        # Main Title & Big Takeaway
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.78), Inches(11.733), Inches(0.85))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title_text
        run.font.name = FONT_MAIN
        run.font.size = Pt(21)
        run.font.bold = True
        run.font.color.rgb = C_TEXT_DARK

        if takeaway_sentence:
            p2 = tf.add_paragraph()
            p2.space_before = Pt(3)
            r2 = p2.add_run()
            r2.text = f"💡 核心觀點：{takeaway_sentence}"
            r2.font.name = FONT_MAIN
            r2.font.size = Pt(11.5)
            r2.font.bold = True
            r2.font.color.rgb = C_TEAL

    def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(1)
        else:
            card.line.fill.background()
        return card

    # =========================================================================
    # SLIDE 1: COVER SLIDE (Minimal, Modern, Bold)
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, C_DARK_BG)

    # Accent bar
    glow = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.2), Inches(1.5), Inches(0.08))
    glow.fill.solid()
    glow.fill.fore_color.rgb = C_PRIMARY
    glow.line.fill.background()

    # Category Badge
    badge1 = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.5), Inches(4.2), Inches(0.4))
    badge1.fill.solid()
    badge1.fill.fore_color.rgb = RGBColor(30, 58, 138)
    badge1.line.color.rgb = RGBColor(59, 130, 246)
    tf_b1 = badge1.text_frame
    tf_b1.margin_left = tf_b1.margin_top = tf_b1.margin_right = tf_b1.margin_bottom = 0
    p_b1 = tf_b1.paragraphs[0]
    p_b1.alignment = PP_ALIGN.CENTER
    r_b1 = p_b1.add_run()
    r_b1.text = "國立臺北科技大學　資訊與生活｜期中專題報告"
    r_b1.font.name = FONT_MAIN
    r_b1.font.size = Pt(11)
    r_b1.font.bold = True
    r_b1.font.color.rgb = RGBColor(191, 219, 254)

    # Big Title
    tbox1 = slide1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(2.0))
    tf1 = tbox1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = tf1.margin_top = tf1.margin_right = tf1.margin_bottom = 0

    p_t1 = tf1.paragraphs[0]
    r_t1 = p_t1.add_run()
    r_t1.text = "科技對工作市場的影響"
    r_t1.font.name = FONT_MAIN
    r_t1.font.size = Pt(38)
    r_t1.font.bold = True
    r_t1.font.color.rgb = C_WHITE

    p_sub1 = tf1.add_paragraph()
    p_sub1.space_before = Pt(12)
    r_sub1 = p_sub1.add_run()
    r_sub1.text = "歷史規律・全球數據矛盾・台灣「低失業、大缺工」實證"
    r_sub1.font.name = FONT_MAIN
    r_sub1.font.size = Pt(18)
    r_sub1.font.color.rgb = RGBColor(148, 163, 184)

    # Team Card
    card_group = slide1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.7), Inches(11.333), Inches(1.8))
    card_group.fill.solid()
    card_group.fill.fore_color.rgb = C_DARK_CARD
    card_group.line.color.rgb = RGBColor(51, 65, 85)

    tf_g = card_group.text_frame
    tf_g.word_wrap = True
    tf_g.margin_left = Inches(0.4)
    tf_g.margin_top = Inches(0.25)
    
    p_g_title = tf_g.paragraphs[0]
    r_gt = p_g_title.add_run()
    r_gt.text = "第 16 組"
    r_gt.font.name = FONT_MAIN
    r_gt.font.size = Pt(15)
    r_gt.font.bold = True
    r_gt.font.color.rgb = RGBColor(96, 165, 250)

    p_members = tf_g.add_paragraph()
    p_members.space_before = Pt(10)
    r_m = p_members.add_run()
    r_m.text = "36 黃柏瑋 (113590023)　｜　37 黃熠瑋 (113590044)　｜　38 楊承諭 (113590051)"
    r_m.font.name = FONT_MAIN
    r_m.font.size = Pt(13)
    r_m.font.color.rgb = C_WHITE

    p_date = tf_g.add_paragraph()
    p_date.space_before = Pt(8)
    r_d = p_date.add_run()
    r_d.text = "報告日期：2026 年 10 月 24 日　　版本：ver.Gemini.1.0.1"
    r_d.font.name = FONT_MAIN
    r_d.font.size = Pt(11)
    r_d.font.color.rgb = RGBColor(148, 163, 184)

    # =========================================================================
    # SLIDE 2: 3 PILLARS (全篇三大核心脈絡)
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, C_LIGHT_BG)
    add_header(slide2, "報告導覽", "全篇三大核心脈絡", "從歷史規律、數據矛盾到台灣實況，釐清 AI 時代的真實挑戰")

    pillars = [
        ("01 歷史定律", "消滅規則，創造新職", "長期總就業必然增長，但「轉型期世代」需承擔陣痛。", [
            ("兩百年歷史定律", "七次技術變革終局皆創造更多新工作，總就業未曾萎縮。"),
            ("任務替代本質", "技術永遠先消滅「規則化任務」，而非淘汰整個人類崗位。"),
            ("恩格斯停頓警訊", "生產力飆升至工資分享常落後數十年，轉型痛點落在過渡世代。")
        ], C_PRIMARY, C_PRIMARY_LIGHT),
        ("02 數據解密", "算法不同，結論不同", "指標定義各自為政，高盛、麥肯錫與 WEF 不可混為一談。", [
            ("高盛「工作份量」", "44% 任務可被自動化，代表任務滲透度，絕非裁員 44%。"),
            ("麥肯錫「工時接管」", "估算 20~30% 工時被接管，人力被釋放轉向更高階任務。"),
            ("WEF 與美丹實證", "淨增 1,200 萬職缺；丹麥實證顯示技術更偏向賦能與輔助。")
        ], C_TEAL, C_TEAL_LIGHT),
        ("03 台灣診斷", "結構錯位，缺工轉型", "失業率 3.36% 創 26 年新低，核心問題是「缺工與技能脫節」。", [
            ("顛覆大失業假說", "失業率創 26 年新低，少子化與高齡化使各產業普遍大缺工。"),
            ("雙重結構性錯位", "「有工無人做，有人無工做」，學校培育與實務需求嚴重斷層。"),
            ("企業賦能非裁員", "金融與服務業「導入不等於裁員」，關鍵在流程重組與技能升級。")
        ], C_AMBER, C_AMBER_LIGHT)
    ]

    for i, (tag, headline, highlight, points, color, light_bg) in enumerate(pillars):
        left = Inches(0.8 + i * 4.0)
        card_w = Inches(3.733)
        card_h = Inches(5.1)
        card = add_card(slide2, left, Inches(1.8), card_w, card_h, C_WHITE, RGBColor(226, 232, 240))
        
        # Top color accent bar
        top_bar = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.8), card_w, Inches(0.09))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = color
        top_bar.line.fill.background()

        # Tag / Pillar Badge
        tag_box = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.28), Inches(2.08), Inches(1.5), Inches(0.32))
        tag_box.fill.solid()
        tag_box.fill.fore_color.rgb = light_bg
        tag_box.line.color.rgb = color
        tag_box.line.width = Pt(0.75)
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = True
        tf_tag.margin_left = tf_tag.margin_top = tf_tag.margin_right = tf_tag.margin_bottom = 0
        p_tag = tf_tag.paragraphs[0]
        p_tag.alignment = PP_ALIGN.CENTER
        r_tag = p_tag.add_run()
        r_tag.text = tag
        r_tag.font.name = FONT_MAIN
        r_tag.font.size = Pt(11)
        r_tag.font.bold = True
        r_tag.font.color.rgb = color

        # Text Frame for Headline & Content
        tbox = slide2.shapes.add_textbox(left + Inches(0.28), Inches(2.52), card_w - Inches(0.56), Inches(4.2))
        tf = tbox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        # Headline
        p_h = tf.paragraphs[0]
        r_h = p_h.add_run()
        r_h.text = headline
        r_h.font.name = FONT_MAIN
        r_h.font.size = Pt(19)
        r_h.font.bold = True
        r_h.font.color.rgb = C_TEXT_DARK

        # Highlight Callout
        p_hl = tf.add_paragraph()
        p_hl.space_before = Pt(10)
        r_hl = p_hl.add_run()
        r_hl.text = f"💡 {highlight}"
        r_hl.font.name = FONT_MAIN
        r_hl.font.size = Pt(11.5)
        r_hl.font.bold = True
        r_hl.font.color.rgb = color

        # Divider line
        p_div = tf.add_paragraph()
        p_div.space_before = Pt(8)
        r_div = p_div.add_run()
        r_div.text = "—" * 24
        r_div.font.name = FONT_MAIN
        r_div.font.size = Pt(8)
        r_div.font.color.rgb = RGBColor(203, 213, 225)

        # 3 Detailed Points
        for pt_title, pt_desc in points:
            p_pt = tf.add_paragraph()
            p_pt.space_before = Pt(8)
            r_t = p_pt.add_run()
            r_t.text = f"• {pt_title}：\n  "
            r_t.font.name = FONT_MAIN
            r_t.font.size = Pt(11)
            r_t.font.bold = True
            r_t.font.color.rgb = C_TEXT_DARK

            r_d = p_pt.add_run()
            r_d.text = pt_desc
            r_d.font.name = FONT_MAIN
            r_d.font.size = Pt(10.5)
            r_d.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 2B: AGENDA ROADMAP (全篇簡報章節地圖)
    # =========================================================================
    slide2_agenda = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2_agenda, C_LIGHT_BG)
    add_header(slide2_agenda, "報告導覽", "全篇簡報章節地圖", "循序漸進剖析技術規律、全球數據辨析、勞動結構分化與個人行動指引")

    agenda_chapters = [
        ("第一章", "歷史規律", "兩百年來七次技術變革的共同模式，與生成式 AI 的本質轉變",
         "回顧蒸汽機至智慧手機的替代規律；剖析高學歷「規則化任務」被取代本質與恩格斯停頓", C_PRIMARY, C_PRIMARY_LIGHT),
        ("第二章", "數據解密", "為什麼全球統計互相矛盾？三種算法與美丹實證對決",
         "拆解高盛（任務）、麥肯錫（工時）、WEF（人數）算法差異，以實證釐清真實衝擊", C_TEAL, C_TEAL_LIGHT),
        ("第三章", "職場分化", "消失的一端（規則化任務）vs 成長的一端（分析思辨與 AI 溢價）",
         "勞動市場兩極化走勢：中階文書與代碼遭受擠壓，複合思辨與人際協調溢價提升", RGBColor(99, 102, 241), RGBColor(238, 242, 255)),
        ("第四章", "台灣實況", "失業率 26 年新低卻大缺工；金融業與服務業「導入不等於裁員」",
         "檢視台灣少子化與大缺工並存實況；企業 AI 導入以「賦能與效率化」為核心而非裁員", C_AMBER, C_AMBER_LIGHT),
        ("第五至七章", "落差與啟示", "產學思辨脫鉤、政策資源分配與未來個人的行動指引",
         "探討教育端與產業界技能落差、政府政策資源精準挹注，以及個人終身學習三維指引", C_ROSE, C_ROSE_LIGHT)
    ]

    for i, (ch_tag, ch_title, ch_headline, ch_desc, color, light_bg) in enumerate(agenda_chapters):
        row_top = Inches(1.80 + i * 1.02)
        row_h = Inches(0.92)
        row_w = Inches(11.733)
        left = Inches(0.8)

        # Outer Card
        card = add_card(slide2_agenda, left, row_top, row_w, row_h, C_WHITE, RGBColor(226, 232, 240))

        # Left Accent Line
        bar = slide2_agenda.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, row_top, Inches(0.08), row_h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        # Left Chapter Badge Box
        badge_box = slide2_agenda.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.22), row_top + Inches(0.16), Inches(2.1), Inches(0.60))
        badge_box.fill.solid()
        badge_box.fill.fore_color.rgb = light_bg
        badge_box.line.color.rgb = color
        badge_box.line.width = Pt(0.75)
        tf_bb = badge_box.text_frame
        tf_bb.word_wrap = True
        tf_bb.margin_left = tf_bb.margin_top = tf_bb.margin_right = tf_bb.margin_bottom = 0
        
        p_b1 = tf_bb.paragraphs[0]
        p_b1.alignment = PP_ALIGN.CENTER
        r_b1 = p_b1.add_run()
        r_b1.text = ch_tag
        r_b1.font.name = FONT_MAIN
        r_b1.font.size = Pt(10)
        r_b1.font.bold = True
        r_b1.font.color.rgb = color

        p_b2 = tf_bb.add_paragraph()
        p_b2.alignment = PP_ALIGN.CENTER
        r_b2 = p_b2.add_run()
        r_b2.text = ch_title
        r_b2.font.name = FONT_MAIN
        r_b2.font.size = Pt(13)
        r_b2.font.bold = True
        r_b2.font.color.rgb = C_TEXT_DARK

        # Right Content Box
        tbox = slide2_agenda.shapes.add_textbox(left + Inches(2.55), row_top + Inches(0.12), row_w - Inches(2.8), Inches(0.68))
        tf = tbox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p_title = tf.paragraphs[0]
        r_tit = p_title.add_run()
        r_tit.text = ch_headline
        r_tit.font.name = FONT_MAIN
        r_tit.font.size = Pt(12.5)
        r_tit.font.bold = True
        r_tit.font.color.rgb = C_TEXT_DARK

        p_sub = tf.add_paragraph()
        p_sub.space_before = Pt(3)
        r_sub = p_sub.add_run()
        r_sub.text = ch_desc
        r_sub.font.name = FONT_MAIN
        r_sub.font.size = Pt(10.5)
        r_sub.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 3: 7 REVOLUTIONS (High-impact Timeline Grid)
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, C_LIGHT_BG)
    add_header(slide3, "第一章・歷史規律", "兩百年來七次技術變革：重複的規律", "每次變革都先消滅「規則化任務」，總就業長期回升，但轉型期需付代價")

    revs = [
        ("1. 工業革命", "18-19 世紀", "蒸汽機與紡織機", "手工紡織工失業，勞動力轉入工廠；生活水準停滯數十年（恩格斯停頓）"),
        ("2. 電力與內燃機", "19-20 世紀初", "汽車與裝配線", "馬伕鐵匠萎縮；福特流水線降低門檻，創造製造、公路與物流業"),
        ("3. 農業機械化", "20 世紀中", "拖拉機與收割機", "美國農業就業從 40% 降至 2% 以下；勞動力成功移轉至製造與服務業"),
        ("4. 自動化機器人", "1960s-90s", "組裝與焊接機器人", "中低技能製造業萎縮；結合全球化分工，形成美國「鏽帶」衰退"),
        ("5. 資訊科技與電腦", "1970s-00s", "PC 與套裝軟體", "打字簿記員消失，就業極化；高技能薪資上升（技能偏向型進步）"),
        ("6. 網際網路與電商", "1990s-10s", "電商與平台經濟", "實體零售重創（百視達倒閉）；誕生數位行銷、現代物流與外送零工"),
        ("7. 智慧型手機", "2007 年至今", "App 與行動支付", "催生行動應用生態圈，重構交通出行、餐飲配送與自媒體創作")
    ]

    card_w = Inches(3.733)
    card_h = Inches(2.1)
    
    # Grid: 3 top, 3 bottom, plus 1 big takeaway card on the bottom-right
    positions = [
        (Inches(0.8), Inches(1.85)),
        (Inches(4.8), Inches(1.85)),
        (Inches(8.8), Inches(1.85)),
        (Inches(0.8), Inches(4.2)),
        (Inches(4.8), Inches(4.2)),
        (Inches(8.8), Inches(4.2)),
    ]

    for idx in range(6):
        title, time_span, tech, impact = revs[idx]
        pos_x, pos_y = positions[idx]
        card = add_card(slide3, pos_x, pos_y, card_w, card_h, C_WHITE, C_CARD_BORDER)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.18)

        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = f"{title} ({time_span})"
        r.font.name = FONT_MAIN
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = C_PRIMARY

        p_t = tf.add_paragraph()
        r_t = p_t.add_run()
        r_t.text = f"關鍵技術：{tech}"
        r_t.font.name = FONT_MAIN
        r_t.font.size = Pt(10)
        r_t.font.bold = True
        r_t.font.color.rgb = C_TEXT_DARK

        p_i = tf.add_paragraph()
        p_i.space_before = Pt(3)
        r_i = p_i.add_run()
        r_i.text = impact
        r_i.font.name = FONT_MAIN
        r_i.font.size = Pt(9.5)
        r_i.font.color.rgb = C_TEXT_MUTED

    # Bottom Banner replacing card 7 + takeaway
    bottom_banner = add_card(slide3, Inches(0.8), Inches(6.45), Inches(11.733), Inches(0.75), C_PRIMARY_LIGHT, RGBColor(191, 219, 254))
    tf_bb = bottom_banner.text_frame
    tf_bb.word_wrap = True
    tf_bb.margin_left = Inches(0.3)
    tf_bb.margin_top = Inches(0.12)
    p_bb = tf_bb.paragraphs[0]
    r_bb = p_bb.add_run()
    r_bb.text = "📌 第 7 次智慧手機（2007~）與歷史核心教訓： "
    r_bb.font.name = FONT_MAIN
    r_bb.font.size = Pt(10.5)
    r_bb.font.bold = True
    r_bb.font.color.rgb = C_PRIMARY

    r_bb2 = p_bb.add_run()
    r_bb2.text = "技術不會消滅總就業，但「能被寫成明確規則」的工作最先被淘汰，轉型成本落在來不及換技能的人身上。"
    r_bb2.font.name = FONT_MAIN
    r_bb2.font.size = Pt(10)
    r_bb2.font.color.rgb = C_TEXT_DARK

    # =========================================================================
    # SLIDE 4: GENERATIVE AI - WHAT'S UNIQUE (High Contrast Comparison)
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, C_LIGHT_BG)
    add_header(slide4, "第一章・當前變革", "這次有何不同？生成式 AI 的本質區別", "從體力勞動轉向「認知型工作」——迭代以月計算，規則化任務首當其衝")

    # Left: Past vs Now Card
    card_l4 = add_card(slide4, Inches(0.8), Inches(1.85), Inches(5.7), Inches(5.0), C_WHITE, C_CARD_BORDER)
    tf_l4 = card_l4.text_frame
    tf_l4.word_wrap = True
    tf_l4.margin_left = tf_l4.margin_right = Inches(0.35)
    tf_l4.margin_top = Inches(0.3)

    pl4 = tf_l4.paragraphs[0]
    rl4 = pl4.add_run()
    rl4.text = "維度對比：過去變革 vs. 生成式 AI"
    rl4.font.name = FONT_MAIN
    rl4.font.size = Pt(14)
    rl4.font.bold = True
    rl4.font.color.rgb = C_PRIMARY

    diffs = [
        ("受影響族群", "過去：體力勞動、工廠藍領、基層文書", "現在：白領階級、知識工作者（認知型任務）"),
        ("典型任務", "過去：產線組裝、馬車駕駛、謄寫記帳", "現在：客服對答、翻譯、文案、初階程式、插畫"),
        ("演進速度", "過去：以數十年計算，勞工有世代緩衝", "現在：以月份與季度迭代，工具滲透極快"),
        ("替代分界", "過去：低技術 vs 高技術", "現在：能否被拆解為「明確規則與步驟」")
    ]

    for title, past, now in diffs:
        p_t = tf_l4.add_paragraph()
        p_t.space_before = Pt(10)
        r_t = p_t.add_run()
        r_t.text = f"【{title}】"
        r_t.font.name = FONT_MAIN
        r_t.font.size = Pt(11)
        r_t.font.bold = True
        r_t.font.color.rgb = C_TEXT_DARK

        p_v = tf_l4.add_paragraph()
        r_p = p_v.add_run()
        r_p.text = f"• 歷史：{past}\n• AI：{now}"
        r_p.font.name = FONT_MAIN
        r_p.font.size = Pt(10)
        r_p.font.color.rgb = C_TEXT_MUTED

    # Right: 2 Big Takeaway Blocks
    card_r4_1 = add_card(slide4, Inches(6.8), Inches(1.85), Inches(5.733), Inches(2.35), C_TEAL_LIGHT, RGBColor(94, 234, 212))
    tf_r1 = card_r4_1.text_frame
    tf_r1.word_wrap = True
    tf_r1.margin_left = tf_r1.margin_right = Inches(0.35)
    tf_r1.margin_top = Inches(0.25)

    pr1 = tf_r1.paragraphs[0]
    rr1 = pr1.add_run()
    rr1.text = "打破迷思：高學歷不再是避風港"
    rr1.font.name = FONT_MAIN
    rr1.font.size = Pt(14)
    rr1.font.bold = True
    rr1.font.color.rgb = C_TEAL

    pr1_d = tf_r1.add_paragraph()
    pr1_d.space_before = Pt(8)
    rr1_d = pr1_d.add_run()
    rr1_d.text = "• AI 具備強大符號處理與模式生成能力。\n• 初階程式撰寫與法律文書首當其衝。\n• 關鍵不是「學歷高低」，而是工作能否被演算法規則化封裝。"
    rr1_d.font.name = FONT_MAIN
    rr1_d.font.size = Pt(11)
    rr1_d.font.color.rgb = C_TEXT_DARK

    card_r4_2 = add_card(slide4, Inches(6.8), Inches(4.5), Inches(5.733), Inches(2.35), C_AMBER_LIGHT, RGBColor(252, 211, 77))
    tf_r2 = card_r4_2.text_frame
    tf_r2.word_wrap = True
    tf_r2.margin_left = tf_r2.margin_right = Inches(0.35)
    tf_r2.margin_top = Inches(0.25)

    pr2 = tf_r2.paragraphs[0]
    rr2 = pr2.add_run()
    rr2.text = "轉型痛點：成本落在誰身上？"
    rr2.font.name = FONT_MAIN
    rr2.font.size = Pt(14)
    rr2.font.bold = True
    rr2.font.color.rgb = C_AMBER

    pr2_d = tf_r2.add_paragraph()
    pr2_d.space_before = Pt(8)
    rr2_d = pr2_d.add_run()
    rr2_d.text = "• 歷史七次變革最終都創造了更多工作。\n• 但在劇變轉型期，來不及換技能的人將承擔失業與低薪代價。\n• 這次速度更快，給社會的適應時間更短。"
    rr2_d.font.name = FONT_MAIN
    rr2_d.font.size = Pt(11)
    rr2_d.font.color.rgb = C_TEXT_DARK

    # =========================================================================
    # SLIDE 5: 3 MEASUREMENTS PARADOX (Data Cleaned)
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, C_LIGHT_BG)
    add_header(slide5, "第二章・數據解密", "全球統計為何矛盾？三種算法的真相", "指標不同，不能比較：高盛算上限、麥肯錫算工時、WEF 算人數")

    algo_cards = [
        ("算法一：碰到的程度", "衡量「技術理論上限」", C_PRIMARY, [
            ("代表機構", "ILO (國際勞工組織)、高盛"),
            ("數據指標", "1/4 勞工受影響，但僅 3.3% 屬最高等級"),
            ("真相拆解", "高盛「3 億個工作受影響」是指【受影響的任務份量】，並非 3 億人失業。")
        ]),
        ("算法二：工時被自動化比例", "衡量「工時自動化潛能」", C_TEAL, [
            ("代表機構", "麥肯錫 (McKinsey)"),
            ("數據指標", "2030 年美國最多 30% 工時具自動化潛力"),
            ("真相拆解", "算的是【時間 (Hours)】而非職位。30% 時間被 AI 代勞，代表生產力提升，不是裁員 30%。")
        ]),
        ("算法三：就業人數淨增減", "衡量「市場總體增減」", C_AMBER, [
            ("代表機構", "WEF (世界經濟論壇)"),
            ("數據指標", "新增 1.7 億、消滅 9,200 萬、淨增 7,800 萬"),
            ("真相拆解", "只看淨增會低估衝擊！消失的與新增的職缺，在【地點與技能】上完全脫節。")
        ])
    ]

    for i, (m_title, m_sub, color, items) in enumerate(algo_cards):
        left = Inches(0.8 + i * 4.0)
        card = add_card(slide5, left, Inches(1.85), Inches(3.733), Inches(5.0), C_WHITE, C_CARD_BORDER)

        top_strip = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.85), Inches(3.733), Inches(0.08))
        top_strip.fill.solid()
        top_strip.fill.fore_color.rgb = color
        top_strip.line.fill.background()

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.22)

        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = m_title
        r.font.name = FONT_MAIN
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = color

        p_sub = tf.add_paragraph()
        r_sub = p_sub.add_run()
        r_sub.text = m_sub
        r_sub.font.name = FONT_MAIN
        r_sub.font.size = Pt(10)
        r_sub.font.color.rgb = C_TEXT_MUTED

        for label, content in items:
            p_lbl = tf.add_paragraph()
            p_lbl.space_before = Pt(12)
            r_l = p_lbl.add_run()
            r_l.text = f"• {label}：\n"
            r_l.font.name = FONT_MAIN
            r_l.font.size = Pt(10.5)
            r_l.font.bold = True
            r_l.font.color.rgb = C_TEXT_DARK

            r_c = p_lbl.add_run()
            r_c.text = f"  {content}"
            r_c.font.name = FONT_MAIN
            r_c.font.size = Pt(10)
            r_c.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 6: EMPIRICAL CONTRAST (US VS DENMARK) - Simplified
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6, C_LIGHT_BG)
    add_header(slide6, "第二章・實證研究", "兩個相反的研究：微觀企業 vs 宏觀國家", "為什麼企業追蹤初階驟減 19%，而全國資料卻顯示「衝擊小於 2%」？")

    # Left: US Canaries
    card_us = add_card(slide6, Inches(0.8), Inches(1.85), Inches(5.7), Inches(2.9), C_WHITE, C_CARD_BORDER)
    top_us = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.7), Inches(0.08))
    top_us.fill.solid()
    top_us.fill.fore_color.rgb = C_ROSE
    top_us.line.fill.background()

    tf_us = card_us.text_frame
    tf_us.word_wrap = True
    tf_us.margin_left = tf_us.margin_right = Inches(0.3)
    tf_us.margin_top = Inches(0.2)

    pu = tf_us.paragraphs[0]
    ru = pu.add_run()
    ru.text = "研究一：美國 Canaries 企業追蹤（微觀前線）"
    ru.font.name = FONT_MAIN
    ru.font.size = Pt(13)
    ru.font.bold = True
    ru.font.color.rgb = C_ROSE

    p_stat_u = tf_us.add_paragraph()
    p_stat_u.space_before = Pt(4)
    r_su = p_stat_u.add_run()
    r_su.text = "初階就業：-19%（資深反而增加）"
    r_su.font.name = FONT_MAIN
    r_su.font.size = Pt(16)
    r_su.font.bold = True
    r_su.font.color.rgb = C_ROSE

    pu_items = [
        "對象：已導入生成式 AI 的前線企業",
        "現象：組織內部初階入門職缺被顯著壓縮",
        "本質：反映特定前線企業的【內部職位重構】"
    ]
    for txt in pu_items:
        p = tf_us.add_paragraph()
        p.space_before = Pt(3)
        r = p.add_run()
        r.text = f"• {txt}"
        r.font.name = FONT_MAIN
        r.font.size = Pt(10)
        r.font.color.rgb = C_TEXT_DARK

    # Right: Denmark Registry
    card_dk = add_card(slide6, Inches(6.8), Inches(1.85), Inches(5.733), Inches(2.9), C_WHITE, C_CARD_BORDER)
    top_dk = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), Inches(1.85), Inches(5.733), Inches(0.08))
    top_dk.fill.solid()
    top_dk.fill.fore_color.rgb = C_TEAL
    top_dk.line.fill.background()

    tf_dk = card_dk.text_frame
    tf_dk.word_wrap = True
    tf_dk.margin_left = tf_dk.margin_right = Inches(0.3)
    tf_dk.margin_top = Inches(0.2)

    pdk = tf_dk.paragraphs[0]
    rdk = pdk.add_run()
    rdk.text = "研究二：丹麥全國人口登記（宏觀總體）"
    rdk.font.name = FONT_MAIN
    rdk.font.size = Pt(13)
    rdk.font.bold = True
    rdk.font.color.rgb = C_TEAL

    p_stat_d = tf_dk.add_paragraph()
    p_stat_d.space_before = Pt(4)
    r_sd = p_stat_d.add_run()
    r_sd.text = "就業衝擊：< 2%（完全看不出影響）"
    r_sd.font.name = FONT_MAIN
    r_sd.font.size = Pt(16)
    r_sd.font.bold = True
    r_sd.font.color.rgb = C_TEAL

    pdk_items = [
        "對象：涵蓋 11 個職業、2.5 萬名勞工、7,000 個場所",
        "現象：宏觀勞動市場未出現統計顯著之失業衝擊",
        "本質：反映全體國家市場的【宏觀淨效果】"
    ]
    for txt in pdk_items:
        p = tf_dk.add_paragraph()
        p.space_before = Pt(3)
        r = p.add_run()
        r.text = f"• {txt}"
        r.font.name = FONT_MAIN
        r.font.size = Pt(10)
        r.font.color.rgb = C_TEXT_DARK

    # Bottom: 3 Core Reasons
    card_diff = add_card(slide6, Inches(0.8), Inches(4.95), Inches(11.733), Inches(1.9), C_PRIMARY_LIGHT, RGBColor(191, 219, 254))
    tf_diff = card_diff.text_frame
    tf_diff.word_wrap = True
    tf_diff.margin_left = tf_diff.margin_right = Inches(0.35)
    tf_diff.margin_top = Inches(0.2)

    pd = tf_diff.paragraphs[0]
    rd = pd.add_run()
    rd.text = "為何兩者落差如此巨大？三大關鍵解釋"
    rd.font.name = FONT_MAIN
    rd.font.size = Pt(12)
    rd.font.bold = True
    rd.font.color.rgb = C_PRIMARY

    reasons = [
        ("1. 勞動體制", "丹麥採「彈性安全制 (Flexicurity)」，雇主多調整工時與內部轉職，不輕易解僱。"),
        ("2. 樣本聚焦", "美國研究鎖定已導入 AI 企業（放大人力調整）；丹麥涵蓋全市場（衝擊被整體稀釋）。"),
        ("3. 調整週期", "兩年追蹤期遠短於企業全面重組週期；因此「19% 是微觀結構衝擊，2% 是宏觀淨效果」。")
    ]
    for r_title, r_desc in reasons:
        p = tf_diff.add_paragraph()
        p.space_before = Pt(4)
        r1 = p.add_run()
        r1.text = f"• {r_title}： "
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = r_desc
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_TEXT_DARK

    # =========================================================================
    # SLIDE 7: DISAPPEARING VS EXPANDING (Concise & Table)
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide7, C_LIGHT_BG)
    add_header(slide7, "第三章・職場分化", "哪些工作在消失、哪些在增加？", "能被拆解為規則的工作減少，分析思辨與 AI 技能享有高額溢價")

    # Left: Disappearing Jobs
    card_l7 = add_card(slide7, Inches(0.8), Inches(1.85), Inches(5.7), Inches(5.0), C_WHITE, C_CARD_BORDER)
    tf_l7 = card_l7.text_frame
    tf_l7.word_wrap = True
    tf_l7.margin_left = tf_l7.margin_right = Inches(0.3)
    tf_l7.margin_top = Inches(0.2)

    pl7 = tf_l7.paragraphs[0]
    rl7 = pl7.add_run()
    rl7.text = "減少端：規則化明確的職業"
    rl7.font.name = FONT_MAIN
    rl7.font.size = Pt(13)
    rl7.font.bold = True
    rl7.font.color.rgb = C_ROSE

    pl7_s = tf_l7.add_paragraph()
    rl7_s = pl7_s.add_run()
    rl7_s.text = "美國勞工統計局 (BLS) 2025–2035 推估減少前三名："
    rl7_s.font.name = FONT_MAIN
    rl7_s.font.size = Pt(10)
    rl7_s.font.color.rgb = C_TEXT_MUTED

    # Table
    table_shape = slide7.shapes.add_table(4, 3, Inches(1.0), Inches(2.6), Inches(5.3), Inches(1.9))
    table = table_shape.table
    table.columns[0].width = Inches(1.7)
    table.columns[1].width = Inches(1.3)
    table.columns[2].width = Inches(2.3)

    t_headers = ["職業別", "10 年變化", "主要成因"]
    t_rows = [
        ["客服代表", "-5% (-14.1萬)", "生成式 AI 處理標準查詢"],
        ["銀行櫃員", "-13% (-4.4萬)", "臨櫃業務線上化"],
        ["電腦程式員", "-7% (-8,100)", "程式碼自動生成工具普及"]
    ]

    for col_idx, h_text in enumerate(t_headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(241, 245, 249)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h_text
        r.font.name = FONT_MAIN
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = C_TEXT_DARK

    for r_idx, row_data in enumerate(t_rows):
        for c_idx, val in enumerate(row_data):
            cell = table.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_WHITE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if c_idx == 1 else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = val
            r.font.name = FONT_MAIN
            r.font.size = Pt(9)
            if c_idx == 1:
                r.font.bold = True
                r.font.color.rgb = C_ROSE
            else:
                r.font.color.rgb = C_TEXT_DARK

    # Left Bottom Note
    p_sc = tf_l7.add_paragraph()
    p_sc.space_before = Pt(85) # push below table
    r_sc = p_sc.add_run()
    r_sc.text = "🏢 企業實例：渣打集團 (Standard Chartered)\n"
    r_sc.font.name = FONT_MAIN
    r_sc.font.size = Pt(10)
    r_sc.font.bold = True
    r_sc.font.color.rgb = C_TEXT_DARK
    r_sc2 = p_sc.add_run()
    r_sc2.text = "2030 前裁減後勤 >15% (>7,000人)，同時上調獲利指標至 15% 以上。後勤規律性工作面臨全面壓縮。"
    r_sc2.font.name = FONT_MAIN
    r_sc2.font.size = Pt(9.5)
    r_sc2.font.color.rgb = C_TEXT_MUTED

    # Right: Expanding Skills
    card_r7 = add_card(slide7, Inches(6.8), Inches(1.85), Inches(5.733), Inches(5.0), C_WHITE, C_CARD_BORDER)
    tf_r7 = card_r7.text_frame
    tf_r7.word_wrap = True
    tf_r7.margin_left = tf_r7.margin_right = Inches(0.3)
    tf_r7.margin_top = Inches(0.2)

    pr7 = tf_r7.paragraphs[0]
    rr7 = pr7.add_run()
    rr7.text = "增加端：技能重塑與薪資溢價"
    rr7.font.name = FONT_MAIN
    rr7.font.size = Pt(13)
    rr7.font.bold = True
    rr7.font.color.rgb = C_TEAL

    r_stats = [
        ("39% 核心技能過時", "雇主預期到 2030 年，員工近 4 成核心技能將被轉型或淘汰。"),
        ("70% 企業必備思辨", "10 家企業中有 7 家將「分析性思考 (Analytical Thinking)」列為招募必備。"),
        ("+56% AI 技能溢價", "同一職業中，會用 AI 的勞工薪資溢價達 56%（前一年調查僅 25%）。"),
        ("技能迭代加快 66%", "AI 接觸度最高職業，技能演進速度比最低者快 66%。")
    ]

    for stat_h, stat_d in r_stats:
        p = tf_r7.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = f"• {stat_h}\n"
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = f"  {stat_d}"
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 8: MECHANISM - STANFORD STUDY (Big Numbers & Clear Insights)
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide8, C_LIGHT_BG)
    add_header(slide8, "第三章・機制剖析", "史丹佛 5,179 名客服實驗：成長階梯的消失", "AI 讓新手快速及格，但吃掉了入門學習價值，壓縮專業晉升路徑")

    # 3 Big Stat Cards
    stat_boxes = [
        ("整體平均效率", "+14%", "5,179 名客服導入 AI 後平均問題解決率提升", C_PRIMARY),
        ("新手菜鳥效率", "+34%", "新手效率爆發成長，大幅跳過摸索磨合期", C_TEAL),
        ("資深老手效率", "幾乎無感", "AI 無法超越資深專家的情境判斷與直覺", C_AMBER)
    ]

    for i, (label, val, desc, color) in enumerate(stat_boxes):
        left = Inches(0.8 + i * 4.0)
        card = add_card(slide8, left, Inches(1.85), Inches(3.733), Inches(2.0), C_WHITE, C_CARD_BORDER)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.18)

        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = label
        r.font.name = FONT_MAIN
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = C_TEXT_MUTED

        p_v = tf.add_paragraph()
        r_v = p_v.add_run()
        r_v.text = val
        r_v.font.name = FONT_MAIN
        r_v.font.size = Pt(28)
        r_v.font.bold = True
        r_v.font.color.rgb = color

        p_d = tf.add_paragraph()
        r_d = p_d.add_run()
        r_d.text = desc
        r_d.font.name = FONT_MAIN
        r_d.font.size = Pt(9.5)
        r_d.font.color.rgb = C_TEXT_DARK

    # Bottom In-depth Takeaways Card
    card_b8 = add_card(slide8, Inches(0.8), Inches(4.1), Inches(11.733), Inches(2.8), C_WHITE, C_CARD_BORDER)
    tf_b8 = card_b8.text_frame
    tf_b8.word_wrap = True
    tf_b8.margin_left = tf_b8.margin_right = Inches(0.35)
    tf_b8.margin_top = Inches(0.25)

    pb = tf_b8.paragraphs[0]
    rb = pb.add_run()
    rb.text = "深層機制：被壓縮的成長階梯 (Flattened Growth Ladder)"
    rb.font.name = FONT_MAIN
    rb.font.size = Pt(13)
    rb.font.bold = True
    rb.font.color.rgb = C_PRIMARY

    mechanisms = [
        ("學習價值被剝離", "以往新人靠處理基礎件累積「暗默知識 (Tacit Knowledge)」。AI 代勞後，新人失去了從實戰摸索沉澱的機會。"),
        ("起聘門檻提高", "企業不再需要大量新人做基礎工作，而是直接要求能審查 AI、具備綜合思辨的即戰力，新鮮人起聘門檻提升。"),
        ("專家斷層隱憂", "若所有人都依賴 AI 產出「及格答案」，未來誰能成長為能糾正 AI 偏誤、解決非典型極端問題的頂尖專家？")
    ]

    for title_m, desc_m in mechanisms:
        p = tf_b8.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = f"• {title_m}： "
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = desc_m
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 9: TAIWAN REALITY (Clean Numbers & Structure)
    # =========================================================================
    slide9 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide9, C_LIGHT_BG)
    add_header(slide9, "第四章・台灣現況", "台灣的真實狀況：失業率 26 年新低，卻又大缺工", "不是職位太少，而是「人與職位對不上」——結構性不匹配")

    # 3 Big Stat Cards
    tw_data = [
        ("失業率創 26 年新低", "3.36%", "季節調整後 3.33%，完全不存在所謂「AI 造成大失業潮」", C_TEAL),
        ("職缺總數高檔成長", "29.6 萬", "年增 2.9 萬個；製造業缺 11.1 萬 (+27%)、服務業缺 16.4 萬", C_PRIMARY),
        ("民生服務深陷大缺工", "67.3 萬", "住宿餐飲缺 27.3 萬、零售缺 20.3 萬；求供比僅 0.28（4 缺搶 1 人）", C_AMBER)
    ]

    for i, (title_t, val_t, desc_t, color_t) in enumerate(tw_data):
        left = Inches(0.8 + i * 4.0)
        card = add_card(slide9, left, Inches(1.85), Inches(3.733), Inches(2.0), C_WHITE, C_CARD_BORDER)
        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.18)

        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title_t
        r.font.name = FONT_MAIN
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = C_TEXT_MUTED

        p_v = tf.add_paragraph()
        r_v = p_v.add_run()
        r_v.text = val_t
        r_v.font.name = FONT_MAIN
        r_v.font.size = Pt(28)
        r_v.font.bold = True
        r_v.font.color.rgb = color_t

        p_d = tf.add_paragraph()
        r_d = p_d.add_run()
        r_d.text = desc_t
        r_d.font.name = FONT_MAIN
        r_d.font.size = Pt(9.5)
        r_d.font.color.rgb = C_TEXT_DARK

    # Bottom Diagnosis Card
    card_b9 = add_card(slide9, Inches(0.8), Inches(4.1), Inches(11.733), Inches(2.8), C_WHITE, C_CARD_BORDER)
    tf_b9 = card_b9.text_frame
    tf_b9.word_wrap = True
    tf_b9.margin_left = tf_b9.margin_right = Inches(0.35)
    tf_b9.margin_top = Inches(0.25)

    pb9 = tf_b9.paragraphs[0]
    rb9 = pb9.add_run()
    rb9.text = "核心病因：結構性不匹配 (Structural Mismatch)"
    rb9.font.name = FONT_MAIN
    rb9.font.size = Pt(13)
    rb9.font.bold = True
    rb9.font.color.rgb = C_PRIMARY

    tw_points = [
        ("AI 高階人才極度稀缺", "104 統計每月 AI 人才招募需求約 10 萬人，五年暴增 97.6%（遠高於全體之 22.9%）。高階技術人才嚴重供不應求。"),
        ("基層民生服務無人可招", "少子化衝擊勞動供給，青年求職者不願投入低薪或重複性體力勞動之民生服務業，造成餐廳、旅宿業陷入空前人力荒。"),
        ("結論：雙重結構斷裂", "台灣面對的不是「AI 奪走工作」，而是「高端技術人手不足」與「基層缺工無人問津」的人職錯位與技能斷層。")
    ]

    for title_p, desc_p in tw_points:
        p = tf_b9.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = f"• {title_p}： "
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = desc_p
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 10: TAIWAN EMPIRICAL - AI ADOPTION != LAYOFFS (Clean)
    # =========================================================================
    slide10 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide10, C_LIGHT_BG)
    add_header(slide10, "第四章・產業實務", "反直覺的實證：導入 AI 不等於裁員", "企業導入 AI 是為了「補足人力缺口」與「釋放服務量能」，而非精簡編制")

    # Left: Financial Survey
    card_fsc = add_card(slide10, Inches(0.8), Inches(1.85), Inches(5.7), Inches(5.0), C_WHITE, C_CARD_BORDER)
    top_fsc = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.7), Inches(0.08))
    top_fsc.fill.solid()
    top_fsc.fill.fore_color.rgb = C_PRIMARY
    top_fsc.line.fill.background()

    tf_fsc = card_fsc.text_frame
    tf_fsc.word_wrap = True
    tf_fsc.margin_left = tf_fsc.margin_right = Inches(0.3)
    tf_fsc.margin_top = Inches(0.25)

    pfsc = tf_fsc.paragraphs[0]
    rfsc = pfsc.add_run()
    rfsc.text = "金管會 2026 金融業 AI 調查"
    rfsc.font.name = FONT_MAIN
    rfsc.font.size = Pt(13)
    rfsc.font.bold = True
    rfsc.font.color.rgb = C_PRIMARY

    fsc_stats = [
        ("43% 金融業已導入 AI", "383 家業者中 164 家導入；生成式 AI 滲透率自 27% 飆至 76%"),
        ("55% 自動化決策為零", "核心決策仍由專業人員把關，AI 僅定位為資料檢索與內部輔助"),
        ("56% 明確未精簡人力", "214 家表示無裁員；多數採遇缺不補或轉調支援高附加價值業務"),
        ("官方權威聲明", "目前台灣查無任何可靠數據證明 AI 造成實質就業流失")
    ]
    for lbl, txt in fsc_stats:
        p = tf_fsc.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = f"• {lbl}\n"
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = f"  {txt}"
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Right: Service Cases
    card_svc = add_card(slide10, Inches(6.8), Inches(1.85), Inches(5.733), Inches(5.0), C_WHITE, C_CARD_BORDER)
    top_svc = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), Inches(1.85), Inches(5.733), Inches(0.08))
    top_svc.fill.solid()
    top_svc.fill.fore_color.rgb = C_TEAL
    top_svc.line.fill.background()

    tf_svc = card_svc.text_frame
    tf_svc.word_wrap = True
    tf_svc.margin_left = tf_svc.margin_right = Inches(0.3)
    tf_svc.margin_top = Inches(0.25)

    psvc = tf_svc.paragraphs[0]
    rsvc = psvc.add_run()
    rsvc.text = "民生服務業落地案例"
    rsvc.font.name = FONT_MAIN
    rsvc.font.size = Pt(13)
    rsvc.font.bold = True
    rsvc.font.color.rgb = C_TEAL

    svc_cases = [
        ("漢來海港 AI 語音客服", "同時接聽 10 條訂位線路，半年處理 15 萬通電話。\n現場外場人力釋放 20%，員工專注於桌邊優質服務。"),
        ("全家便利商店自助結帳", "逾 1,000 家門市普及自助機台。\n董事長強調：目的在「紓解尖峰排隊與分擔店員多工」，絕非裁員。"),
        ("核心洞察總結", "在少子化嚴重的台灣，AI 是填補基層缺口的「救火隊」，維持運營而非取代勞工。")
    ]
    for brand, detail in svc_cases:
        p = tf_svc.add_paragraph()
        p.space_before = Pt(10)
        r1 = p.add_run()
        r1.text = f"• {brand}\n"
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = f"  {detail}"
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 11: EDUCATION-WORKPLACE GAP (Punchy)
    # =========================================================================
    slide11 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide11, C_LIGHT_BG)
    add_header(slide11, "第五章・教育落差", "學校教的 vs 職場要的：三層結構性脫節", "只教工具操作，回應不了職場真正渴望的「系統性思辨能力」")

    layers = [
        ("第一層：技能折舊", "技術半衰期急遽縮短", C_PRIMARY, [
            ("39% 技能過時", "2030 年前近 4 成核心專業技能將轉型淘汰"),
            ("教材滯後", "課堂軟體與代碼在畢業前即被新工具代換")
        ]),
        ("第二層：職缺變形", "需求與恐慌兩極化", C_TEAL, [
            ("AI 職缺逆勢大增", "總職缺 -9%，但要求 AI 職缺 +28.1%"),
            ("87.2% 勞工焦慮", "擔憂被取代，不知如何有效重塑技能")
        ]),
        ("第三層：供給脫節", "預算投入與需求錯位", C_AMBER, [
            ("大學百億投入", "教育部在 41 所大學設 AI 學程，砸 117 億元"),
            ("71.2% 缺乏思辨", "人資主管指出：最缺的是【系統性思辨】，單教 Prompt 毫無意義！")
        ])
    ]

    for i, (l_title, l_sub, color_l, pts) in enumerate(layers):
        left = Inches(0.8 + i * 4.0)
        card = add_card(slide11, left, Inches(1.85), Inches(3.733), Inches(5.0), C_WHITE, C_CARD_BORDER)

        top_bar = slide11.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.85), Inches(3.733), Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = color_l
        top_bar.line.fill.background()

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.22)

        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = l_title
        r.font.name = FONT_MAIN
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = color_l

        p_s = tf.add_paragraph()
        r_s = p_s.add_run()
        r_s.text = l_sub
        r_s.font.name = FONT_MAIN
        r_s.font.size = Pt(10)
        r_s.font.color.rgb = C_TEXT_MUTED

        for p_name, p_desc in pts:
            p_item = tf.add_paragraph()
            p_item.space_before = Pt(14)
            r1 = p_item.add_run()
            r1.text = f"• {p_name}：\n"
            r1.font.name = FONT_MAIN
            r1.font.size = Pt(10.5)
            r1.font.bold = True
            r1.font.color.rgb = C_TEXT_DARK

            r2 = p_item.add_run()
            r2.text = f"  {p_desc}"
            r2.font.name = FONT_MAIN
            r2.font.size = Pt(10)
            r2.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 12: POLICY COMPARISON & BUDGET SKEW (Clean)
    # =========================================================================
    slide12 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide12, C_LIGHT_BG)
    add_header(slide12, "第六章・政策工具", "政府做了什麼？台灣三大工具箱與國際借鑑", "台灣「先補人、再補技能」；審計部警示經費 100% 偏袒電子業，傳產為零")

    # Left: Taiwan Policy
    card_tw_pol = add_card(slide12, Inches(0.8), Inches(1.85), Inches(5.7), Inches(5.0), C_WHITE, C_CARD_BORDER)
    top_tw_p = slide12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.85), Inches(5.7), Inches(0.08))
    top_tw_p.fill.solid()
    top_tw_p.fill.fore_color.rgb = C_PRIMARY
    top_tw_p.line.fill.background()

    tf_tp = card_tw_pol.text_frame
    tf_tp.word_wrap = True
    tf_tp.margin_left = tf_tp.margin_right = Inches(0.3)
    tf_tp.margin_top = Inches(0.2)

    ptp = tf_tp.paragraphs[0]
    rtp = ptp.add_run()
    rtp.text = "台灣勞動部三大政策工具"
    rtp.font.name = FONT_MAIN
    rtp.font.size = Pt(13)
    rtp.font.bold = True
    rtp.font.color.rgb = C_PRIMARY

    tw_tools = [
        ("補技能", "產業人才投資方案（3 年 10 萬）、青年學習獎勵月領 8,000 元"),
        ("補就業", "青年尋職津貼 1.8 萬，就業滿 90 天加發 3 萬（已助 1.5 萬人）"),
        ("保在職", "減班休息再充電計畫，每月最高補助 1 萬 6,300 元津貼避免裁員")
    ]
    for t_name, t_detail in tw_tools:
        p = tf_tp.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = f"• {t_name}： "
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = t_detail
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # Audit Alert
    p_aud = tf_tp.add_paragraph()
    p_aud.space_before = Pt(12)
    raud = p_aud.add_run()
    raud.text = "⚠️ 審計部重大警訊：經費分配嚴重失衡\n"
    raud.font.name = FONT_MAIN
    raud.font.size = Pt(11)
    raud.font.bold = True
    raud.font.color.rgb = C_ROSE

    raud_d = p_aud.add_run()
    raud_d.text = "經濟部三大 AI 計畫中，兩項受惠對象【100% 為電子業，傳產為 0%】！資源過度向高科技傾斜，傳統產業轉型嚴重被邊緣化。"
    raud_d.font.name = FONT_MAIN
    raud_d.font.size = Pt(9.5)
    raud_d.font.color.rgb = C_TEXT_DARK

    # Right: International Benchmarks
    card_intl = add_card(slide12, Inches(6.8), Inches(1.85), Inches(5.733), Inches(5.0), C_WHITE, C_CARD_BORDER)
    top_intl = slide12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.8), Inches(1.85), Inches(5.733), Inches(0.08))
    top_intl.fill.solid()
    top_intl.fill.fore_color.rgb = C_TEAL
    top_intl.line.fill.background()

    tf_intl = card_intl.text_frame
    tf_intl.word_wrap = True
    tf_intl.margin_left = tf_intl.margin_right = Inches(0.3)
    tf_intl.margin_top = Inches(0.2)

    pi = tf_intl.paragraphs[0]
    ri = pi.add_run()
    ri.text = "國際借鑑：新加坡、德國與歐盟"
    ri.font.name = FONT_MAIN
    ri.font.size = Pt(13)
    ri.font.bold = True
    ri.font.color.rgb = C_TEAL

    intl_models = [
        ("新加坡 SkillsFuture 全民學習帳戶", "個人專屬終身培訓基金，政府直接注資。\n46% 合格國民使用過，打造全民學習文化。"),
        ("德國 Kurzarbeit 短時工作津貼", "政府補貼薪資差額以減少裁員。\n2025 年吸收 1.33 億工時，方案延至 2026 年底保住熟練工。"),
        ("歐盟 AI Act 法律責任明定", "第 26 條明訂：導入高風險 AI 系統之企業，必須負起【人類監督 (Human Oversight)】責任，杜絕無限制替代。")
    ]

    for country, details in intl_models:
        p = tf_intl.add_paragraph()
        p.space_before = Pt(8)
        r1 = p.add_run()
        r1.text = f"• {country}\n"
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = f"  {details}"
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 13: CONCLUSIONS & ACTIONABLE TAKEAWAYS (Clean)
    # =========================================================================
    slide13 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide13, C_LIGHT_BG)
    add_header(slide13, "第七章・總結啟示", "研究總結：三大定論與未來個人的行動指引", "面對 AI 浪潮，不需要恐慌失業，關鍵在於培養「無法被規則化的思辨力」")

    # 3 Conclusion Cards
    conclusions = [
        ("定論一：先釐清算法", "上限 ≠ 工時 ≠ 失業人數", "辨析指標口徑，勿被「AI 奪走 3 億工作」的聳動標題誤導。", C_PRIMARY),
        ("定論二：體制決定衝擊", "微觀劇烈 vs 宏觀稀釋", "美國特定企業初階減少 19%，丹麥宏觀衝擊 < 2%，體制緩衝是關鍵。", C_TEAL),
        ("定論三：台灣重在缺工", "結構錯位而非大裁員", "歷史定律未變；台灣核心挑戰是少子化缺工與高階思辨人才匱乏。", C_AMBER)
    ]

    for i, (title_c, sub_c, desc_c, color_c) in enumerate(conclusions):
        left = Inches(0.8 + i * 4.0)
        card = add_card(slide13, left, Inches(1.85), Inches(3.733), Inches(2.0), C_WHITE, C_CARD_BORDER)
        top_bar = slide13.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.85), Inches(3.733), Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = color_c
        top_bar.line.fill.background()

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title_c
        r.font.name = FONT_MAIN
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = color_c

        p_s = tf.add_paragraph()
        r_s = p_s.add_run()
        r_s.text = sub_c
        r_s.font.name = FONT_MAIN
        r_s.font.size = Pt(13)
        r_s.font.bold = True
        r_s.font.color.rgb = C_TEXT_DARK

        p_d = tf.add_paragraph()
        p_d.space_before = Pt(4)
        r_d = p_d.add_run()
        r_d.text = desc_c
        r_d.font.name = FONT_MAIN
        r_d.font.size = Pt(9.5)
        r_d.font.color.rgb = C_TEXT_MUTED

    # Bottom Student Action Cards
    card_rec = add_card(slide13, Inches(0.8), Inches(4.1), Inches(11.733), Inches(2.8), C_WHITE, C_CARD_BORDER)
    tf_rec = card_rec.text_frame
    tf_rec.word_wrap = True
    tf_rec.margin_left = tf_rec.margin_right = Inches(0.35)
    tf_rec.margin_top = Inches(0.25)

    prec = tf_rec.paragraphs[0]
    rrec = prec.add_run()
    rrec.text = "給新世代大學生與職場工作者的三大行動指引"
    rrec.font.name = FONT_MAIN
    rrec.font.size = Pt(13)
    rrec.font.bold = True
    rrec.font.color.rgb = C_PRIMARY

    actions = [
        ("從「執行者」轉型為「架構決策者」", "只會照本宣科的執行者容易被取代；但能定義問題、拆解流程、審查 AI 結果並承擔責任的人，薪資享有 +56% 溢價。"),
        ("厚植不可替代的「系統性思辨力」", "71.2% 企業最渴望思辨力。別只滿足於當個會下 Prompt 的操作員，更要掌握業務邏輯、批判性思考與人際同理心。"),
        ("建立終身「跨域技能重塑 (Reskilling)」心態", "技能 5 年內有 4 成面臨過時。學習新加坡 SkillsFuture 精神，將終身學習化為主動習慣，隨技術共同進化。")
    ]

    for title_a, desc_a in actions:
        p = tf_rec.add_paragraph()
        p.space_before = Pt(6)
        r1 = p.add_run()
        r1.text = f"• {title_a}： "
        r1.font.name = FONT_MAIN
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = C_TEXT_DARK
        r2 = p.add_run()
        r2.text = desc_a
        r2.font.name = FONT_MAIN
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 14: THANK YOU & QA
    # =========================================================================
    slide14 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide14, C_DARK_BG)

    card_ty = slide14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.2), Inches(1.5), Inches(8.933), Inches(4.5))
    card_ty.fill.solid()
    card_ty.fill.fore_color.rgb = C_DARK_CARD
    card_ty.line.color.rgb = RGBColor(51, 65, 85)

    tf_ty = card_ty.text_frame
    tf_ty.word_wrap = True
    tf_ty.margin_left = tf_ty.margin_right = Inches(0.5)
    tf_ty.margin_top = Inches(0.5)

    p_badge = tf_ty.paragraphs[0]
    p_badge.alignment = PP_ALIGN.CENTER
    r_b = p_badge.add_run()
    r_b.text = "國立臺北科技大學　資訊與生活"
    r_b.font.name = FONT_MAIN
    r_b.font.size = Pt(13)
    r_b.font.bold = True
    r_b.font.color.rgb = RGBColor(96, 165, 250)

    p_thx = tf_ty.add_paragraph()
    p_thx.alignment = PP_ALIGN.CENTER
    p_thx.space_before = Pt(16)
    r_thx = p_thx.add_run()
    r_thx.text = "感謝聆聽・敬請指教"
    r_thx.font.name = FONT_MAIN
    r_thx.font.size = Pt(32)
    r_thx.font.bold = True
    r_thx.font.color.rgb = C_WHITE

    p_qa = tf_ty.add_paragraph()
    p_qa.alignment = PP_ALIGN.CENTER
    p_qa.space_before = Pt(10)
    r_qa = p_qa.add_run()
    r_qa.text = "Q & A　問 答 與 交 流 討 論"
    r_qa.font.name = FONT_MAIN
    r_qa.font.size = Pt(18)
    r_qa.font.bold = True
    r_qa.font.color.rgb = RGBColor(203, 213, 225)

    p_info = tf_ty.add_paragraph()
    p_info.alignment = PP_ALIGN.CENTER
    p_info.space_before = Pt(24)
    r_info = p_info.add_run()
    r_info.text = "第 16 組：36 黃柏瑋 ｜ 37 黃熠瑋 ｜ 38 楊承諭\n簡報版本：期中報告_ver.Gemini.1.0.1.pptx"
    r_info.font.name = FONT_MAIN
    r_info.font.size = Pt(12)
    r_info.font.color.rgb = RGBColor(148, 163, 184)

    # Save presentation
    prs.save(output_path)
    print(f"Refactored presentation saved: {output_path}")

if __name__ == "__main__":
    out_file = "期中報告_ver.Gemini.1.0.1.pptx"
    build_presentation(out_file)
