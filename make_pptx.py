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
    add_header(slide2, "導讀指南", "全篇報告核心：看清 AI 與工作的真實樣貌", "與其被「AI 搶飯碗」的新聞嚇自己，不如從歷史規律、全球數據與台灣現場看清真相")

    pillars = [
        ("01 歷史定律", "淘汰的是雜事，不是人類", "長遠來看工作一定會變多，但轉型期的那代人總要扛下代價。", [
            ("兩百年的鐵律", "七次技術革命每次都先淘汰舊職缺，但最後創造的新工作永遠比消失的更多。"),
            ("誰最容易被取代", "技術最先吃掉的，永遠是「能照 SOP 照抄」的固定流程，而不是人類本身。"),
            ("殘酷的現實", "歷史上的「恩格斯停頓」告訴我們：工廠賺翻了，工人的薪水卻可能卡住幾十年。")
        ], C_PRIMARY, C_PRIMARY_LIGHT),
        ("02 數據解密", "算盤打法不同，結論差很多", "高盛算任務、麥肯錫算時間、WEF 數人頭，千萬別混在一起談。", [
            ("高盛被誤解的 3 億", "報告說 3 億工作「受影響」，是指工作內容被科技滲透，不是開除 3 億人。"),
            ("麥肯錫算的工時", "AI 幫忙省下 20~30% 的作業時間，代表工作變輕鬆高效，不是裁掉三成人。"),
            ("看淨增數會盲目", "WEF 說全球淨增 1,200 萬個職缺，但消失的工作跟新開的技能完全對不上。")
        ], C_TEAL, C_TEAL_LIGHT),
        ("03 台灣現場", "根本沒大失業，而是到處缺人", "失業率創 26 年新低，台灣真正的痛點是「缺工與學用脫節」。", [
            ("打破失業迷思", "失業率跌到 3.36% 歷史新低，少子化與缺工並存，壓根沒有所謂「AI 裁員潮」。"),
            ("結構嚴重錯位", "現場是「企業找不到懂用 AI 的人，畢業生找不到合適的高薪缺」，兩頭對不上。"),
            ("公司導入不是裁員", "台灣金融與餐飲業引進 AI，全是因為人力吃緊拿來當救火隊，不是為了精簡編制。")
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
    add_header(slide2_agenda, "章節導覽", "全篇簡報章節地圖", "五大板塊循序漸進：從歷史規律出發，辨析統計迷思，深挖台灣現場與個人生存之道")

    agenda_chapters = [
        ("第一章", "歷史規律", "兩百年來七次技術變革的共同劇本，與這次生成式 AI 的本質跳躍",
         "回顧蒸汽機到智慧手機的衝擊史；為什麼高學歷白領這次不能再拿文憑當免死金牌？", C_PRIMARY, C_PRIMARY_LIGHT),
        ("第二章", "數據解密", "為什麼全球各大機構的預測互相打架？拆解三種算法的真實意圖",
         "看懂高盛、麥肯錫與 WEF 到底在算什麼；以美國微觀前線對比丹麥宏觀人口實證。", C_TEAL, C_TEAL_LIGHT),
        ("第三章", "職場分化", "誰的工作在消失、誰在身價翻倍？史丹佛客服實驗揭開的殘酷真相",
         "規則明確的初階工作慘遭壓縮，具備批判思辨與 AI 協同能力的專業人才享有巨大溢價。", RGBColor(99, 102, 241), RGBColor(238, 242, 255)),
        ("第四章", "台灣實況", "失業率創 26 年新低卻大喊缺工！台灣各行各業導入 AI 的真實心態",
         "金管會金融業大調查與漢來、全家實證：AI 在少子化時代成了撐住營運的關鍵救火隊。", C_AMBER, C_AMBER_LIGHT),
        ("第五至七章", "落差與啟示", "大學開課開錯方向？政府經費偏心？面對 AI 時代的大學生該怎麼自保？",
         "直指產學斷層與政策資源失衡，為新世代職場人整理出無法被演算法取代的三大行動指南。", C_ROSE, C_ROSE_LIGHT)
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
    # SLIDE 3A: 歷史技術變革（一）：動力革命與機械化時代
    # =========================================================================
    slide3_p1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3_p1, C_LIGHT_BG)
    add_header(slide3_p1, "第一章・歷史規律", "歷史技術變革（一）：蒸汽動力到農業機械化", "人類第一次從粗重體力勞動中被解放：每一次恐慌，最終都換來了更大的產業舞台")

    stage1_revs = [
        ("01 工業革命", "18-19 世紀", "機器搶生計，工人砸織布機", "蒸汽機・機械紡織機", "手工織布工一夕失業，爆發憤怒的「盧德運動」", [
            ("衝擊現場", "手工織布師傅被無情取代，勞工被迫擠進環境惡劣的初期現代工廠。"),
            ("新職崛起", "蒸汽火車、鐵路網絡、現代採礦與機械維修應運而生，催生了現代工業。"),
            ("殘酷教訓", "歷史上的「恩格斯停頓」——工廠賺大錢，勞工生活水準卻苦等幾十年才跟上。")
        ], C_PRIMARY, C_PRIMARY_LIGHT),
        ("02 電力與內燃機", "19-20 世紀初", "汽車跑起來，馬伕去哪了？", "發電機・汽車・福特流水線", "流水線把複雜工藝拆解，沒受訓的人也能上工", [
            ("衝擊現場", "馬車伕、鐵匠與馬具行紛紛關門，整個以馬為主的傳統運輸生態圈瓦解。"),
            ("新職崛起", "汽車工廠、加油站、公路養護、跨州物流與各種家用電器走入千家萬戶。"),
            ("翻轉階級", "福特把工作拆成簡單步驟，反而讓百萬名沒經驗的基層農民拿到穩定薪水。")
        ], C_TEAL, C_TEAL_LIGHT),
        ("03 農業機械化", "20 世紀中葉", "農田不再需要那麼多人", "拖拉機・聯合收割機・化肥", "四成人口的工作不見了，總體失業率卻沒因此暴跌", [
            ("衝擊現場", "美國農耕人口占總就業比例從 1900 年的 40%，一路崩落到今天的不到 2%。"),
            ("新職崛起", "離開農田的勞動力，剛好無縫接軌成為二戰後製造業與現代服務業的中堅。"),
            ("核心啟發", "單一傳統產業的萎縮並不可怕，科技帶來的富裕會長出更多全新行業。")
        ], C_AMBER, C_AMBER_LIGHT)
    ]

    for i, (tag, period, headline, tech, highlight, points, color, light_bg) in enumerate(stage1_revs):
        left = Inches(0.8 + i * 4.0)
        card_w = Inches(3.733)
        card_h = Inches(5.1)
        card = add_card(slide3_p1, left, Inches(1.8), card_w, card_h, C_WHITE, RGBColor(226, 232, 240))
        
        # Top color accent bar
        top_bar = slide3_p1.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.8), card_w, Inches(0.09))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = color
        top_bar.line.fill.background()

        # Tag / Era Badge
        tag_box = slide3_p1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.28), Inches(2.08), Inches(2.3), Inches(0.32))
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
        r_tag.text = f"{tag} ({period})"
        r_tag.font.name = FONT_MAIN
        r_tag.font.size = Pt(10.5)
        r_tag.font.bold = True
        r_tag.font.color.rgb = color

        # Content Box
        tbox = slide3_p1.shapes.add_textbox(left + Inches(0.28), Inches(2.52), card_w - Inches(0.56), Inches(4.2))
        tf = tbox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        # Headline
        p_h = tf.paragraphs[0]
        r_h = p_h.add_run()
        r_h.text = headline
        r_h.font.name = FONT_MAIN
        r_h.font.size = Pt(17)
        r_h.font.bold = True
        r_h.font.color.rgb = C_TEXT_DARK

        # Key Tech Subtitle
        p_tech = tf.add_paragraph()
        p_tech.space_before = Pt(3)
        r_tech = p_tech.add_run()
        r_tech.text = f"關鍵技術：{tech}"
        r_tech.font.name = FONT_MAIN
        r_tech.font.size = Pt(10.5)
        r_tech.font.bold = True
        r_tech.font.color.rgb = color

        # Highlight Callout
        p_hl = tf.add_paragraph()
        p_hl.space_before = Pt(8)
        r_hl = p_hl.add_run()
        r_hl.text = f"💡 {highlight}"
        r_hl.font.name = FONT_MAIN
        r_hl.font.size = Pt(10.5)
        r_hl.font.bold = True
        r_hl.font.color.rgb = C_TEXT_DARK

        # Divider line
        p_div = tf.add_paragraph()
        p_div.space_before = Pt(6)
        r_div = p_div.add_run()
        r_div.text = "—" * 24
        r_div.font.name = FONT_MAIN
        r_div.font.size = Pt(8)
        r_div.font.color.rgb = RGBColor(203, 213, 225)

        # 3 Points
        for pt_title, pt_desc in points:
            p_pt = tf.add_paragraph()
            p_pt.space_before = Pt(7)
            r_t = p_pt.add_run()
            r_t.text = f"• {pt_title}：\n  "
            r_t.font.name = FONT_MAIN
            r_t.font.size = Pt(10.5)
            r_t.font.bold = True
            r_t.font.color.rgb = C_TEXT_DARK

            r_d = p_pt.add_run()
            r_d.text = pt_desc
            r_d.font.name = FONT_MAIN
            r_d.font.size = Pt(10)
            r_d.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 3B: 歷史技術變革（二）：自動化、電腦與網路浪潮
    # =========================================================================
    slide3_p2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3_p2, C_LIGHT_BG)
    add_header(slide3_p2, "第一章・歷史規律", "歷史技術變革（二）：機械手臂到網際網路", "機器開始學會「照章辦事」：凡是能寫成 SOP 的例行任務，陸續被晶片與程式碼接管")

    stage2_revs = [
        ("04 自動化機器人", "1960s-1990s", "機械手臂上線，產線換血", "組裝機器人・焊接設備", "流水線機器人加上全球化，造就了蕭條的「鏽帶」", [
            ("衝擊現場", "汽車工廠重複性組裝、焊接工大批被撤換，基層藍領工人首當其衝。"),
            ("新職崛起", "誕生了自動化工程師、機器人設備調試、精密模具與跨國供應鏈管理。"),
            ("轉型陣痛", "技能跟不上的傳統工人被迫提早退休，造就歐美傳統工業城鎮走向凋零。")
        ], RGBColor(99, 102, 241), RGBColor(238, 242, 255)),
        ("05 電腦與資訊科技", "1970s-2000s", "PC 擺上桌，打字員走入歷史", "PC・辦公軟體・關聯式資料庫", "懂軟體的人身價飆漲，中階文書工作被徹底掏空", [
            ("衝擊現場", "專門打字、記帳簿記、人工總機接線生的身影，在幾年內迅速從辦公室絕跡。"),
            ("新職崛起", "軟體工程師、IT 資訊網管、系統分析師成了全市場最搶手的百萬金童。"),
            ("職場兩極化", "「技能偏向型進步」——會用電腦的人身價狂飆，只做文書的人薪資停滯。")
        ], C_PRIMARY, C_PRIMARY_LIGHT),
        ("06 網際網路與電商", "1990s-2010s", "滑鼠點下送到家，店面迎浩劫", "WWW・電商平台・數位支付", "買東西不再需要出門，徹底重構零售與物流遊戲規則", [
            ("衝擊現場", "實體唱片行、百視達錄影帶租借店徹底絕跡，各大傳統連鎖商場門市大縮水。"),
            ("新職崛起", "捧紅了電商營運、數位行銷、現代物流倉儲司機，更誕生了全新的零工經濟。"),
            ("邊際成本歸零", "資訊與商品的接觸管道直接搬上雲端，為接下來的行動網路時代鋪平道路。")
        ], C_TEAL, C_TEAL_LIGHT)
    ]

    for i, (tag, period, headline, tech, highlight, points, color, light_bg) in enumerate(stage2_revs):
        left = Inches(0.8 + i * 4.0)
        card_w = Inches(3.733)
        card_h = Inches(5.1)
        card = add_card(slide3_p2, left, Inches(1.8), card_w, card_h, C_WHITE, RGBColor(226, 232, 240))
        
        # Top color accent bar
        top_bar = slide3_p2.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, Inches(1.8), card_w, Inches(0.09))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = color
        top_bar.line.fill.background()

        # Tag / Era Badge
        tag_box = slide3_p2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.28), Inches(2.08), Inches(2.3), Inches(0.32))
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
        r_tag.text = f"{tag} ({period})"
        r_tag.font.name = FONT_MAIN
        r_tag.font.size = Pt(10.5)
        r_tag.font.bold = True
        r_tag.font.color.rgb = color

        # Content Box
        tbox = slide3_p2.shapes.add_textbox(left + Inches(0.28), Inches(2.52), card_w - Inches(0.56), Inches(4.2))
        tf = tbox.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        # Headline
        p_h = tf.paragraphs[0]
        r_h = p_h.add_run()
        r_h.text = headline
        r_h.font.name = FONT_MAIN
        r_h.font.size = Pt(17)
        r_h.font.bold = True
        r_h.font.color.rgb = C_TEXT_DARK

        # Key Tech Subtitle
        p_tech = tf.add_paragraph()
        p_tech.space_before = Pt(3)
        r_tech = p_tech.add_run()
        r_tech.text = f"關鍵技術：{tech}"
        r_tech.font.name = FONT_MAIN
        r_tech.font.size = Pt(10.5)
        r_tech.font.bold = True
        r_tech.font.color.rgb = color

        # Highlight Callout
        p_hl = tf.add_paragraph()
        p_hl.space_before = Pt(8)
        r_hl = p_hl.add_run()
        r_hl.text = f"💡 {highlight}"
        r_hl.font.name = FONT_MAIN
        r_hl.font.size = Pt(10.5)
        r_hl.font.bold = True
        r_hl.font.color.rgb = C_TEXT_DARK

        # Divider line
        p_div = tf.add_paragraph()
        p_div.space_before = Pt(6)
        r_div = p_div.add_run()
        r_div.text = "—" * 24
        r_div.font.name = FONT_MAIN
        r_div.font.size = Pt(8)
        r_div.font.color.rgb = RGBColor(203, 213, 225)

        # 3 Points
        for pt_title, pt_desc in points:
            p_pt = tf.add_paragraph()
            p_pt.space_before = Pt(7)
            r_t = p_pt.add_run()
            r_t.text = f"• {pt_title}：\n  "
            r_t.font.name = FONT_MAIN
            r_t.font.size = Pt(10.5)
            r_t.font.bold = True
            r_t.font.color.rgb = C_TEXT_DARK

            r_d = p_pt.add_run()
            r_d.text = pt_desc
            r_d.font.name = FONT_MAIN
            r_d.font.size = Pt(10)
            r_d.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 3C: 兩百年來七次技術變革：重複的規律（最後才點名！）
    # =========================================================================
    slide3_p3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3_p3, C_LIGHT_BG)
    add_header(slide3_p3, "第一章・歷史規律", "兩百年來七次技術變革：重複的規律", "歷史教給我們的同一堂課：科技從未消滅就業總量，但能被寫成 SOP 的工作必死無疑")

    # Table Container
    table_shape = slide3_p3.shapes.add_table(8, 4, Inches(0.8), Inches(1.8), Inches(11.733), Inches(3.35))
    table = table_shape.table
    table.columns[0].width = Inches(2.35)
    table.columns[1].width = Inches(2.2)
    table.columns[2].width = Inches(3.4)
    table.columns[3].width = Inches(3.783)

    t_headers = ["變革階段與時代", "核心突破技術", "🔻 被淘汰的「規則化任務」", "🔺 催生的「全新職業與產業」"]
    t_rows = [
        ["01 工業革命 (18-19 世紀)", "蒸汽機・機械紡織機", "手工織布工、傳統工匠（盧德砸機器抗爭）", "工廠流水線工人、鐵路司機、煤礦與機修師"],
        ["02 電力與內燃機 (19-20 世紀初)", "發電機・汽車・福特裝配線", "馬伕、鐵匠、馬車打造與馬具業全面沒落", "汽車製造、加油站員、現代物流、家電維修"],
        ["03 農業機械化 (20 世紀中葉)", "拖拉機・收割機・化肥", "傳統農耕粗重勞力（就業占比 40% 掉到 2%）", "戰後製造業大軍、現代商業營運與服務業"],
        ["04 自動化機器人 (1960s-1990s)", "機械手臂・自動化組裝焊接", "汽車產線裝配工、焊接工等例行體力操作", "自動化維護工程師、機器人調試、供應鏈管理"],
        ["05 電腦資訊 (1970s-2000s)", "個人電腦 (PC)・辦公軟體", "打字員、謄寫員、記帳員、人工電話接線生", "軟體工程師、IT 網管工程師、資料庫管理員"],
        ["06 網際網路 (1990s-2010s)", "全球資訊網・電商・數位金流", "實體唱片行、百視達影帶出租、傳統門市零售", "電商營運、數位行銷、倉儲撿貨員、外送零工"],
        ["07 智慧手機 (2007 年至今)", "App 生態系・4G/5G・行動支付", "傳統按鍵手機通訊、傳統路邊叫車、紙本票券", "App 開發者、社群自媒體創作者、即時外送出行"]
    ]

    for col_idx, h_text in enumerate(t_headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(234, 240, 250)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h_text
        r.font.name = FONT_MAIN
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = C_PRIMARY

    for row_idx, row_data in enumerate(t_rows):
        for col_idx, text in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            if row_idx == 6:  # Highlight 07 Smartphone
                cell.fill.fore_color.rgb = RGBColor(240, 249, 255)
            elif row_idx % 2 == 1:
                cell.fill.fore_color.rgb = RGBColor(248, 250, 252)
            else:
                cell.fill.fore_color.rgb = C_WHITE
            
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = text
            r.font.name = FONT_MAIN
            r.font.size = Pt(9.5)
            if col_idx == 0:
                r.font.bold = True
                r.font.color.rgb = C_PRIMARY if row_idx < 6 else RGBColor(30, 64, 175)
            else:
                r.font.color.rgb = C_TEXT_DARK

    # Bottom Takeaway Card: 3 Laws
    takeaway_card = add_card(slide3_p3, Inches(0.8), Inches(5.35), Inches(11.733), Inches(1.5), C_PRIMARY_LIGHT, RGBColor(191, 219, 254))
    tf_tc = takeaway_card.text_frame
    tf_tc.word_wrap = True
    tf_tc.margin_left = tf_tc.margin_right = Inches(0.35)
    tf_tc.margin_top = Inches(0.16)

    ptc_h = tf_tc.paragraphs[0]
    rtc_h = ptc_h.add_run()
    rtc_h.text = "📌 貫穿兩百年變革的「三大歷史鐵律」"
    rtc_h.font.name = FONT_MAIN
    rtc_h.font.size = Pt(12)
    rtc_h.font.bold = True
    rtc_h.font.color.rgb = C_PRIMARY

    laws = [
        ("鐵律一・總就業長期只會增加", "生產力暴增會把餅做大，創造出我們先前完全想像不到的新產業與新慾望。"),
        ("鐵律二・淘汰的是「任務」不是「人」", "只要事情可以被寫成固定步驟、按鈕執行的，早晚都會被工具接管。"),
        ("鐵律三・轉型期的痛苦最為真實", "科技跑得永遠比教育快，轉型成本永遠是落在「來不及換技能的這代人」身上。")
    ]

    for l_title, l_desc in laws:
        p_law = tf_tc.add_paragraph()
        p_law.space_before = Pt(3)
        r_lt = p_law.add_run()
        r_lt.text = f"• {l_title}： "
        r_lt.font.name = FONT_MAIN
        r_lt.font.size = Pt(10)
        r_lt.font.bold = True
        r_lt.font.color.rgb = C_TEXT_DARK

        r_ld = p_law.add_run()
        r_ld.text = l_desc
        r_ld.font.name = FONT_MAIN
        r_ld.font.size = Pt(10)
        r_ld.font.color.rgb = C_TEXT_MUTED

    # =========================================================================
    # SLIDE 4: GENERATIVE AI - WHAT'S UNIQUE (High Contrast Comparison)
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, C_LIGHT_BG)
    add_header(slide4, "第一章・當前變革", "這次到底有何不同？生成式 AI 的本質區別", "以前是機器搶走勞力，這次是演算法挑高學歷白領下手——速度以月計算，規則化任務首當其衝")

    # Left: Past vs Now Card
    card_l4 = add_card(slide4, Inches(0.8), Inches(1.85), Inches(5.7), Inches(5.0), C_WHITE, C_CARD_BORDER)
    tf_l4 = card_l4.text_frame
    tf_l4.word_wrap = True
    tf_l4.margin_left = tf_l4.margin_right = Inches(0.35)
    tf_l4.margin_top = Inches(0.3)

    pl4 = tf_l4.paragraphs[0]
    rl4 = pl4.add_run()
    rl4.text = "維度對照：過去歷次變革 vs. 這次生成式 AI"
    rl4.font.name = FONT_MAIN
    rl4.font.size = Pt(14)
    rl4.font.bold = True
    rl4.font.color.rgb = C_PRIMARY

    diffs = [
        ("誰先被波及？", "過去：工廠藍領、體力勞動、基層打字員", "現在：寫 code 的、打官司的、做行銷設計的知識白領"),
        ("典型被取代任務", "過去：螺絲組裝、趕馬車、算帳謄寫", "現在：客服對答、公文潤飾、初階寫 code、翻譯插畫"),
        ("演變速度快慢", "過去：給社會幾十年慢慢調適換代", "現在：每個月模型都在升級，幾季內工具就全面普及"),
        ("替代的真界線", "過去：拼的是「有沒有念大學」", "現在：拼的是工作能不能被「拆解成明確 prompt 與步驟」")
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
    rr1.text = "打破迷思：名校高學歷不再是免死金牌"
    rr1.font.name = FONT_MAIN
    rr1.font.size = Pt(14)
    rr1.font.bold = True
    rr1.font.color.rgb = C_TEAL

    pr1_d = tf_r1.add_paragraph()
    pr1_d.space_before = Pt(8)
    rr1_d = pr1_d.add_run()
    rr1_d.text = "• AI 專門擅長符號排列與現成套路模仿。\n• 初階程式碼撰寫、常規合約審查，往往最先被演算法給包辦。\n• 關鍵不是讀到碩士還是博士，而是工作能不能被寫成 SOP。"
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
    rr2.text = "現實痛點：轉型的代價這次落在誰身上？"
    rr2.font.name = FONT_MAIN
    rr2.font.size = Pt(14)
    rr2.font.bold = True
    rr2.font.color.rgb = C_AMBER

    pr2_d = tf_r2.add_paragraph()
    pr2_d.space_before = Pt(8)
    rr2_d = pr2_d.add_run()
    rr2_d.text = "• 歷史教訓證明工作最後一定會變多，但「青黃不接」的陣痛非常難熬。\n• 被淘汰技能的老手如果跟不上，很容易淪為低薪或過渡時期的犧牲者。\n• 這次技術演進實在太快，社會體制跟學校根本來不及替勞工做好準備。"
    rr2_d.font.name = FONT_MAIN
    rr2_d.font.size = Pt(11)
    rr2_d.font.color.rgb = C_TEXT_DARK

    # =========================================================================
    # SLIDE 5: 3 MEASUREMENTS PARADOX (Data Cleaned)
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, C_LIGHT_BG)
    add_header(slide5, "第二章・數據解密", "全球數據為何天天打架？看懂三種算法背後的把戲", "算盤定義不同，結果天差地遠：高盛算理論上限、麥肯錫算工作時間、WEF 數人頭淨額")

    algo_cards = [
        ("算法一：摸到的邊界有多廣", "衡量的是「技術理論上的天花板」", C_PRIMARY, [
            ("代表機構", "ILO 國際勞工組織、高盛集團 (Goldman Sachs)"),
            ("統計數據", "全球 1/4 勞工工作會被碰到，但只有 3.3% 屬於高度替代"),
            ("大白話真相", "高盛說「3 億個工作受影響」，是指 3 億份工作裡「有任務會碰到 AI」，根本不是要開除 3 億人！")
        ]),
        ("算法二：能幫人省下多少工時", "衡量的是「自動化潛在時間效益」", C_TEAL, [
            ("代表機構", "麥肯錫顧問公司 (McKinsey Global Institute)"),
            ("統計數據", "到 2030 年，美國市場最高有 30% 的工作時數具備自動化潛能"),
            ("大白話真相", "算的是【小時 (Hours)】不是【人頭】！省下 30% 時間是讓你不用加班做雜事，不代表公司要裁員 30%。")
        ]),
        ("算法三：到底會多幾個或少幾個人", "衡量的是「全市場工作崗位淨增減」", C_AMBER, [
            ("代表機構", "WEF 世界經濟論壇"),
            ("數據指標", "預估未來幾年新增 1.7 億個工作、淘汰 9,200 萬個，淨增 7,800 萬職缺"),
            ("大白話真相", "別光看淨增加就樂觀！消失的工作跟新開的職缺，在「地點、年齡跟能力」上根本搭不上線。")
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
    add_header(slide6, "第二章・實證研究", "兩個完全相反的權威研究：前線開槍 vs 全國無感", "企業前線跟全體社會看到的完全是兩回事：19% 是微觀換血的陣痛，2% 是宏觀稀釋後的平靜")

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
    r_su.text = "初階入門新人職缺：-19%"
    r_su.font.name = FONT_MAIN
    r_su.font.size = Pt(16)
    r_su.font.bold = True
    r_su.font.color.rgb = C_ROSE

    pu_items = [
        "調查對象：已經搶先全線導入生成式 AI 的前鋒企業",
        "發生什麼事：新人職缺被砍掉近兩成，資深老鳥名額反而變多",
        "真實意涵：反映走在最前面的企業，正在悄悄「用 AI 把底層苦力給打包替換」"
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
    rdk.text = "研究二：丹麥全體人口登記追蹤（宏觀國家）"
    rdk.font.name = FONT_MAIN
    rdk.font.size = Pt(13)
    rdk.font.bold = True
    rdk.font.color.rgb = C_TEAL

    p_stat_d = tf_dk.add_paragraph()
    p_stat_d.space_before = Pt(4)
    r_sd = p_stat_d.add_run()
    r_sd.text = "全國就業淨衝擊：< 2%"
    r_sd.font.name = FONT_MAIN
    r_sd.font.size = Pt(16)
    r_sd.font.bold = True
    r_sd.font.color.rgb = C_TEAL

    pdk_items = [
        "調查對象：橫跨 11 個大行業、2.5 萬名真實勞工、7,000 多家企業",
        "發生什麼事：全國統計完全看不出大規模解僱或失業升高的跡象",
        "真實意涵：反映放眼到整個國家社會，AI 的衝擊被龐大的傳統經濟給均勻稀釋掉了"
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
    rd.text = "為什麼同一件事情，兩邊看到的風景差這麼多？三大關鍵解密"
    rd.font.name = FONT_MAIN
    rd.font.size = Pt(12)
    rd.font.bold = True
    rd.font.color.rgb = C_PRIMARY

    reasons = [
        ("1. 勞動法規與體制", "丹麥實行「彈性安全制」，老闆習慣調配工時和內部轉調，不隨便開除人。"),
        ("2. 挑的樣本不一樣", "美國研究挑的是「最激進導入 AI 的軟體前線」；丹麥研究看的是「買菜看病的真實全國」。"),
        ("3. 組織換血需要時間", "兩年時間只夠前線企業砍掉新人名額，一般企業要全面重組流程，少說要好幾年。")
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
    add_header(slide7, "第三章・職場分化", "職場殘酷分化：誰在被默默淘汰？誰在身價狂飆？", "規矩清楚的基層被快速壓縮，具備問題定義與系統思辨能力的人享有高額薪資溢價")

    # Left: Disappearing Jobs
    card_l7 = add_card(slide7, Inches(0.8), Inches(1.85), Inches(5.7), Inches(5.0), C_WHITE, C_CARD_BORDER)
    tf_l7 = card_l7.text_frame
    tf_l7.word_wrap = True
    tf_l7.margin_left = tf_l7.margin_right = Inches(0.3)
    tf_l7.margin_top = Inches(0.2)

    pl7 = tf_l7.paragraphs[0]
    rl7 = pl7.add_run()
    rl7.text = "退場端：能照著 SOP 辦事的崗位正在縮水"
    rl7.font.name = FONT_MAIN
    rl7.font.size = Pt(13)
    rl7.font.bold = True
    rl7.font.color.rgb = C_ROSE

    pl7_s = tf_l7.add_paragraph()
    rl7_s = pl7_s.add_run()
    rl7_s.text = "美國勞工統計局 (BLS) 推估未來十年減幅前三名："
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
        ["客服代表", "-5% (-14.1萬)", "標準常見問題被 AI 機器人接管"],
        ["銀行櫃員", "-13% (-4.4萬)", "臨櫃業務幾乎全部轉向線上與 App"],
        ["初階程式員", "-7% (-8,100)", "程式碼自動生成工具迅速普及"]
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
    r_sc2.text = "預計 2030 年前裁減後勤超過 15%（逾 7,000 人），同時上調獲利指標。後勤純文書與常規審核的位子正面臨全面大退縮。"
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
    rr7.text = "搶手端：重新定義自我與 AI 協同的專家"
    rr7.font.name = FONT_MAIN
    rr7.font.size = Pt(13)
    rr7.font.bold = True
    rr7.font.color.rgb = C_TEAL

    r_stats = [
        ("39% 專業技能面臨洗牌", "企業老闆預估五年內，員工現有近 4 成的專業本事都要換新。"),
        ("七成企業最想要「會動腦」的人", "10 家企業有 7 家在招募時把「批判性與分析性思辨」列為第一優先。"),
        ("會用 AI 的薪水多拿 +56%", "做同樣一份工作，能把 AI 用得好的工作者，薪資溢價高達 56%（前一年才 25%）。"),
        ("技能更換速度加快 66%", "每天高強度接觸 AI 的行業，其專業技能過時與更新的速度比傳統行業快了三分之二。")
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
    add_header(slide8, "第三章・機制剖析", "史丹佛 5,179 人實測：新手變強了，但階梯卻被抽走了", "AI 讓剛進公司的菜鳥快速及格，但同時也吃掉了「在底層摸爬滾打熬成老手」的歷練機會")

    # 3 Big Stat Cards
    stat_boxes = [
        ("整體客服團隊表現", "+14%", "導入 AI 輔助後，每小時成功解決客戶問題的比例明顯增加", C_PRIMARY),
        ("剛入職的新手菜鳥", "+34%", "新手進步最快！直接套用老鳥經驗庫，大幅跳過自己撞牆的時期", C_TEAL),
        ("十年經驗的資深老手", "幾乎沒差", "AI 對專家幾乎沒有幫助，電腦目前還學不會老手的直覺與判斷力", C_AMBER)
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
    rb.text = "深層省思：被悄悄抽走的職場成長階梯 (Flattened Growth Ladder)"
    rb.font.name = FONT_MAIN
    rb.font.size = Pt(13)
    rb.font.bold = True
    rb.font.color.rgb = C_PRIMARY

    mechanisms = [
        ("實戰磨練的價值被剝奪了", "以前新人是靠一件一件處理鳥事雜事，才磨出「只能意會不能言傳的隱性知識」。現在雜事全被 AI 做完了，新人去哪裡練功？"),
        ("新鮮人起聘門檻不降反升", "老闆不再需要招一批人來做基礎資料整理，而是一開口就要「懂業務邏輯、能抓出 AI 錯誤的即戰力」，菜鳥越來越難入行。"),
        ("未來的頂尖專家會從哪裡來？", "如果大家都習慣按一個鍵看 AI 給的及格答案，十年後還有誰能憑硬實力去修正 AI 的重大盲點？")
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
    add_header(slide9, "第四章・台灣現況", "台灣的真實寫照：失業率創 26 年新低，卻到處大缺工", "台灣根本不是「沒有工作給人做」，而是「企業缺的人找嘸人，找工的人能力對不上」")

    # 3 Big Stat Cards
    tw_data = [
        ("失業率創 26 年新低", "3.36%", "季調後 3.33%，台灣目前壓根不存在任何「AI 造成萬人流落街頭」的痕跡", C_TEAL),
        ("全台職缺數維持高檔", "29.6 萬", "一年內職缺多開了近 3 萬個；製造業缺 11 萬人，服務業更缺了 16 萬人", C_PRIMARY),
        ("民生服務業陷入大乾旱", "67.3 萬", "餐飲旅宿缺 27 萬、零售缺 20 萬；求供比 0.28，等於 4 個職缺搶 1 個人", C_AMBER)
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
    rb9.text = "真正的痛點：兩頭對不上的「雙重結構性錯位」"
    rb9.font.name = FONT_MAIN
    rb9.font.size = Pt(13)
    rb9.font.bold = True
    rb9.font.color.rgb = C_PRIMARY

    tw_points = [
        ("高階懂 AI 的專業人才搶破頭", "104 統計每月 AI 相關職缺高達 10 萬個，五年翻了快一倍，市場上完全找不到足夠的技術人才。"),
        ("基層辛苦的民生服務沒人要去", "少子化浪潮撲面而來，年輕人不再願意進入低薪又高壓的實體服務業，使得餐廳飯店天天為排班焦頭爛額。"),
        ("結論：不是飯碗被敲破，而是技能與願景斷裂", "台灣當前面臨的最大危機，是高端技術缺腦袋、基層缺體力，中間的人不知道往哪裡去的嚴重錯位。")
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
    add_header(slide10, "第四章・產業實務", "跌破眼鏡的實情：在台灣，導入 AI 根本不是為了裁員", "面對少子化與找不到人的窘境，企業買 AI 是拿來「當替補救火隊」，而不是為了開除員工")

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
    rfsc.text = "金管會 2026 金融業 AI 普查實況"
    rfsc.font.name = FONT_MAIN
    rfsc.font.size = Pt(13)
    rfsc.font.bold = True
    rfsc.font.color.rgb = C_PRIMARY

    fsc_stats = [
        ("43% 金融機構已引進 AI", "383 家業者中 164 家導入；生成式 AI 內部使用比例一年內從 27% 暴衝到 76%"),
        ("55% 關鍵決策依然零自動化", "放款、風控、高額理財等核心業務，全部都要專人把關蓋章，AI 只是幫忙抓資料的書僮"),
        ("56% 業者白紙黑字保證沒裁員", "逾兩百家金融機構明確表示不裁員；多數採遇缺不補，把省下的人力拉去支援高收益客製服務"),
        ("主管機關直接給結論", "官方正式定調：目前台灣沒有任何事實或數據能佐證 AI 造成金融業勞工失業")
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
    rsvc.text = "走進街頭：民生餐飲與便利商店的現場"
    rsvc.font.name = FONT_MAIN
    rsvc.font.size = Pt(13)
    rsvc.font.bold = True
    rsvc.font.color.rgb = C_TEAL

    svc_cases = [
        ("漢來海港 AI 語音訂位管家", "一次接 10 通電話，半年幫忙接聽 15 萬通訂位。\n餐廳說得很白：省下來的外場電話人力，是為了讓員工專心招呼現場客人提升品質。"),
        ("全家便利商店自助結帳普及全台", "逾 1,000 家門市普及自助機台。\n董事長公開強調：這是幫店員分擔煮咖啡、取包裹的多工負擔，舒緩尖峰排隊，絕非裁員。"),
        ("台灣現場的核心心態", "在少子化嚴重的台灣，AI 不是「搶飯碗的競爭者」，而是「讓店還能開下去的救命稻草」。")
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
    add_header(slide11, "第五章・教育落差", "學校教的 vs 職場要的：三層嚴重的結構性脫節", "大學如果只教學生怎樣下 Prompt 提示詞，根本是把學生推向畢業即被淘汰的火坑")

    layers = [
        ("第一層：本事過期的速度太快", "專業技術的半衰期被腰斬", C_PRIMARY, [
            ("近四成技能快過期", "國際調查顯示，到 2030 年職場上有 39% 的核心能力將面臨失效"),
            ("教材永遠趕不上改版", "大一認真學的寫 code 工具，還沒熬到大四畢業就已被新一代模型取代")
        ]),
        ("第二層：大家想招 AI 但很焦慮", "市場開缺與勞工心態極度分裂", C_TEAL, [
            ("傳統職缺微降，AI 職缺暴增", "一般職缺減少 9%，但只要工作要求會 AI 的，職缺逆向暴增 28.1%"),
            ("高達 87% 的上班族發慌", "知道 AI 很強、害怕被取代，但每天上班忙翻天，根本不知道該從何學起")
        ]),
        ("第三層：政府錢砸了方向搞錯", "狂開操作課，卻沒人懂動腦思辨", C_AMBER, [
            ("教育部砸 117 億大開學程", "全台 41 所大專院校一窩蜂開了一堆標榜生成式 AI 的訓練班"),
            ("企業人資搖頭：最缺思辨！", "71.2% 的主管反映：學生 Prompt 背很熟，但連最基礎的邏輯拆解與思辨都做不好")
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
    add_header(slide12, "第六章・政策工具", "政府的工具箱管用嗎？台灣現狀與國際借鑑", "台灣習慣「發錢補貼就業」，但審計部重砲警告：預算幾乎百分之百給了半導體，傳產快被餓死")

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
    rtp.text = "台灣勞動部現行的三大補貼手段"
    rtp.font.name = FONT_MAIN
    rtp.font.size = Pt(13)
    rtp.font.bold = True
    rtp.font.color.rgb = C_PRIMARY

    tw_tools = [
        ("補技能", "產業人才投資方案（3 年 10 萬元學費補助）、青年參加 AI 認證每月領 8,000 元獎勵金"),
        ("補就業", "青年尋職津貼給 1.8 萬，穩定上工滿 90 天再加發 3 萬元（已發給 1.5 萬人）"),
        ("保在職", "景氣不好放無薪假時，透過充電再出發計畫補貼訓練津貼，最高月領 1.6 萬元避免裁員")
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
    raud.text = "⚠️ 審計部查帳爆出重大警訊：經費全偏向電子大廠！\n"
    raud.font.name = FONT_MAIN
    raud.font.size = Pt(11)
    raud.font.bold = True
    raud.font.color.rgb = C_ROSE

    raud_d = p_aud.add_run()
    raud_d.text = "經濟部主導的三大 AI 轉型經費中，居然有兩項的補助對象【100% 全被電子資訊大廠拿走，傳統製造與中小服務業拿到的比例是 0%】！資源極度失衡。"
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
    ri.text = "外國政府怎麼做？新加坡、德國與歐盟的借鏡"
    ri.font.name = FONT_MAIN
    ri.font.size = Pt(13)
    ri.font.bold = True
    ri.font.color.rgb = C_TEAL

    intl_models = [
        ("新加坡 SkillsFuture 全民學習存摺", "政府直接在每位國民帳戶裡存進進修培訓金。\n46% 的新加坡人都用過這筆錢去學新技能，把終身學習變成國民習慣。"),
        ("德國 Kurzarbeit 短工工時津貼制度", "景氣差或技術轉型時，政府直接掏錢補貼企業工人的薪資差額，保住老師傅的飯碗與技術，不讓工廠隨便裁員。"),
        ("歐盟 AI Act 嚴格明訂法律人責", "歐盟法律第 26 條講得很硬：任何引進高風險 AI 的企業，背後都必須有「真人全程監督把關」，絕不允許把人丟給演算法盲目決定。")
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
    add_header(slide13, "第七章・總結啟示", "全篇研究總結：看透三大定論，掌握未來個人的生存之道", "不要為虛構的失業潮杞人憂天；真正的護城河，是把演算法當工具、活出無法被 SOP 封裝的思辨力")

    # 3 Conclusion Cards
    conclusions = [
        ("定論一：先看懂數字背後的把戲", "摸到邊界 ≠ 省下時間 ≠ 開除員工", "看到新聞標題說「幾億人要丟工作」，先冷靜看它是算任務還是算人頭，別被標題黨嚇倒。", C_PRIMARY),
        ("定論二：你所在的社會體制決定命運", "前線軟體換血 vs 全國經濟平靜吸收", "美國特定前線企業少掉 19% 新人，丹麥宏觀衝擊不到 2%；國家的法規與產業結構才是真正的安全氣囊。", C_TEAL),
        ("定論三：台灣當下最大的考驗是缺工", "根本沒有大裁員，而是技能嚴重脫鉤", "兩百年歷史鐵律依然管用：少子化讓台灣大缺工，重點不是沒工作，而是我們有沒有準備好企業需要的真本事。", C_AMBER)
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
    rrec.text = "寫給新世代大學生與職場新鮮人的三句真心話"
    rrec.font.name = FONT_MAIN
    rrec.font.size = Pt(13)
    rrec.font.bold = True
    rrec.font.color.rgb = C_PRIMARY

    actions = [
        ("從聽命操作的工具人，轉型為「定義問題、驗收成果、扛責任的人」", "只會照著指令做事的人最容易被換掉；但懂得把模糊目標拆解成步驟、檢查 AI 產出漏洞並為最後結果負責的人，薪水直接多拿 56%。"),
        ("把力氣花在「無法被演算法規則化的系統性思辨力」", "七成以上的企業最渴望的是分析與邏輯思考。千萬別只把時間花在背 Prompt，去搞懂商業邏輯、培養人際同理心與批判性眼光才是真本事。"),
        ("習慣這輩子都要「邊做邊學、持續技能升級」的心態", "技術跑得太快，大學四年學的本事大概五年就有一半要洗牌。學學新加坡的精神，主動跟著工具一起進化，AI 就不是淘汰你的敵人，而是幫你放大的外掛！")
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
