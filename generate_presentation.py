from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.dml import MSO_THEME_COLOR


TITLE = "Ethereum & Smart Contracts"
SUBTITLE = "Team 1 Presentation | Blockchain Technology | Decentralized Applications"
OUTPUT_FILE = "Ethereum_Smart_Contracts_Team_1_Deck.pptx"

DARK_BG = RGBColor(10, 15, 25)
CARD_BG = RGBColor(18, 26, 39)
PANEL_BG = RGBColor(24, 34, 49)
ACCENT = RGBColor(70, 224, 180)
ACCENT_2 = RGBColor(90, 155, 255)
ACCENT_3 = RGBColor(255, 181, 72)
TEXT = RGBColor(234, 239, 245)
MUTED = RGBColor(170, 182, 201)
WHITE = RGBColor(255, 255, 255)
RED = RGBColor(255, 110, 110)
GREEN = RGBColor(120, 230, 160)


def set_background(slide):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = DARK_BG

    # Add subtle network grid
    for x in range(0, 14):
        line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(x * 1.0), 0, Inches(0.02), Inches(7.5))
        line.fill.background()
        line.line.fill.background()
        line.shadow.inherit = False

    for y in range(0, 9):
        line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, Inches(y * 0.9), Inches(13.33), Inches(0.02))
        line.fill.background()
        line.line.fill.background()

    # top accent bar
    accent = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, Inches(13.33), Inches(0.18))
    accent.fill.solid()
    accent.fill.fore_color.rgb = ACCENT
    accent.line.fill.background()


def add_title(slide, title, subtitle=None):
    title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.45), Inches(9.5), Inches(0.8))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.name = 'Aptos'
    run.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.LEFT

    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.72), Inches(1.1), Inches(10.5), Inches(0.4))
        tf2 = sub_box.text_frame
        p2 = tf2.paragraphs[0]
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.size = Pt(10)
        r2.font.name = 'Aptos'
        r2.font.color.rgb = MUTED


def add_textbox(slide, x, y, w, h, text, font_size=18, color=TEXT, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.name = 'Aptos'
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def add_card(slide, x, y, w, h, title=None, body=None, accent_color=ACCENT, title_size=18, body_size=12):
    card = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG
    card.line.color.rgb = RGBColor(45, 60, 80)
    card.line.width = Pt(1)

    accent_bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, Inches(0.12))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = accent_color
    accent_bar.line.fill.background()

    if title:
        add_textbox(slide, x + Inches(0.2), y + Inches(0.22), w - Inches(0.4), Inches(0.5), title, font_size=title_size, color=WHITE, bold=True)
    if body:
        add_textbox(slide, x + Inches(0.2), y + Inches(0.7), w - Inches(0.4), h - Inches(1), body, font_size=body_size, color=MUTED)

    return card


def add_bullet_list(slide, x, y, w, h, items, font_size=18):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    for idx, item in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = f"• {item}"
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(10)
        run = p.runs[0]
        run.font.size = Pt(font_size)
        run.font.name = 'Aptos'
        run.font.color.rgb = TEXT


def add_circle(slide, cx, cy, radius, color=ACCENT, text=None, label=None):
    circle = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.OVAL, cx - radius, cy - radius, radius * 2, radius * 2)
    circle.fill.solid()
    circle.fill.fore_color.rgb = color
    circle.line.color.rgb = color
    circle.line.width = Pt(1.5)
    if text:
        add_textbox(slide, cx - radius + Inches(0.12), cy - Inches(0.22), radius * 2 - Inches(0.24), Inches(0.5), text, font_size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    if label:
        add_textbox(slide, cx - Inches(0.8), cy + Inches(0.5), Inches(1.6), Inches(0.25), label, font_size=9, color=MUTED, align=PP_ALIGN.CENTER)


def add_blockquote(slide, x, y, w, h, text):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = PANEL_BG
    shape.line.color.rgb = ACCENT_2
    shape.line.width = Pt(1.5)
    add_textbox(slide, x + Inches(0.25), y + Inches(0.25), w - Inches(0.5), h - Inches(0.5), text, font_size=16, color=TEXT)


def add_header_band(slide, text):
    band = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.6), Inches(0.5), Inches(3.7), Inches(0.32))
    band.fill.solid()
    band.fill.fore_color.rgb = ACCENT
    band.line.fill.background()
    add_textbox(slide, Inches(0.8), Inches(0.52), Inches(3.0), Inches(0.25), text, font_size=13, color=WHITE, bold=True)


def build_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides = [
        {
            "title": "Ethereum & Smart Contracts",
            "subtitle": "Team 1 Presentation | Blockchain Technology | Decentralized Applications",
            "kind": "title"
        },
        {
            "title": "Agenda",
            "subtitle": "What we will cover",
            "kind": "agenda"
        },
        {
            "title": "Why Blockchain Matters",
            "subtitle": "Trust, transparency, and decentralized coordination",
            "kind": "content"
        },
        {
            "title": "Ethereum at a Glance",
            "subtitle": "The programmable blockchain",
            "kind": "content"
        },
        {
            "title": "How Ethereum Works",
            "subtitle": "Nodes, accounts, and transactions",
            "kind": "content"
        },
        {
            "title": "Transactions & Gas",
            "subtitle": "How state changes are processed",
            "kind": "content"
        },
        {
            "title": "Smart Contracts",
            "subtitle": "Self-executing logic on-chain",
            "kind": "content"
        },
        {
            "title": "Solidity Fundamentals",
            "subtitle": "Writing smart contracts in Ethereum",
            "kind": "content"
        },
        {
            "title": "Remix IDE",
            "subtitle": "Fastest way to write, test, and deploy",
            "kind": "content"
        },
        {
            "title": "Local Development Workflow",
            "subtitle": "From idea to deployment",
            "kind": "content"
        },
        {
            "title": "Ganache",
            "subtitle": "A personal Ethereum blockchain",
            "kind": "content"
        },
        {
            "title": "DApp Architecture",
            "subtitle": "Frontend + blockchain backend",
            "kind": "content"
        },
        {
            "title": "Wallets and Identity",
            "subtitle": "MetaMask, accounts, and signatures",
            "kind": "content"
        },
        {
            "title": "Use Case: DeFi",
            "subtitle": "Financial services without intermediaries",
            "kind": "content"
        },
        {
            "title": "Use Case: NFTs",
            "subtitle": "Unique digital ownership and collectibles",
            "kind": "content"
        },
        {
            "title": "Use Case: Supply Chain",
            "subtitle": "Traceability and auditability",
            "kind": "content"
        },
        {
            "title": "Security Best Practices",
            "subtitle": "Designing safe and reliable smart contracts",
            "kind": "content"
        },
        {
            "title": "Challenges & Scalability",
            "subtitle": "Speed, cost, and network load",
            "kind": "content"
        },
        {
            "title": "Future of Ethereum",
            "subtitle": "Layer 2, upgrades, and new opportunities",
            "kind": "content"
        },
        {
            "title": "Thank You",
            "subtitle": "Questions and discussion",
            "kind": "closing"
        }
    ]

    for idx, slide_info in enumerate(slides, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_background(slide)

        if slide_info["kind"] == "title":
            # Title slide
            shapes = slide.shapes
            title_box = shapes.add_textbox(Inches(0.9), Inches(1.2), Inches(7.5), Inches(1.2))
            tf = title_box.text_frame
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = slide_info["title"]
            r.font.size = Pt(30)
            r.font.bold = True
            r.font.name = 'Aptos'
            r.font.color.rgb = WHITE

            subtitle_box = shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(8.5), Inches(0.7))
            tf2 = subtitle_box.text_frame
            p2 = tf2.paragraphs[0]
            r2 = p2.add_run()
            r2.text = slide_info["subtitle"]
            r2.font.size = Pt(12)
            r2.font.name = 'Aptos'
            r2.font.color.rgb = MUTED

            # blockchain-like graphic
            for x in [1.8, 4.8, 7.8, 10.8]:
                add_circle(slide, Inches(x), Inches(4.8), Inches(0.55), color=ACCENT_2)
            for x in [3.3, 6.3, 9.3]:
                add_circle(slide, Inches(x), Inches(5.8), Inches(0.4), color=ACCENT)
            # connecting lines approximated as shapes
            line1 = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(2.35), Inches(4.9), Inches(2.45), Inches(0.02))
            line1.fill.solid(); line1.fill.fore_color.rgb = ACCENT; line1.line.fill.background()
            line2 = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(5.35), Inches(4.9), Inches(2.45), Inches(0.02))
            line2.fill.solid(); line2.fill.fore_color.rgb = ACCENT_2; line2.line.fill.background()
            line3 = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(8.35), Inches(4.9), Inches(2.45), Inches(0.02))
            line3.fill.solid(); line3.fill.fore_color.rgb = ACCENT_3; line3.line.fill.background()

        else:
            add_header_band(slide, slide_info["title"])

            if slide_info["kind"] == "agenda":
                add_bullet_list(slide, Inches(1.0), Inches(1.4), Inches(5.5), Inches(4.8), [
                    "Blockchain foundations and Ethereum overview",
                    "Smart contracts and Solidity basics",
                    "Remix IDE and local blockchain development",
                    "Ganache, wallets, and DApp architecture",
                    "Use cases, security, and future outlook"
                ], font_size=20)

                # quick metrics cards
                add_card(slide, Inches(7.2), Inches(1.8), Inches(4.5), Inches(1.2), "20 Slides", "End-to-end overview of Ethereum and DApps", ACCENT)
                add_card(slide, Inches(7.2), Inches(3.3), Inches(4.5), Inches(1.2), "Core Stack", "Ethereum + Solidity + Remix + Ganache", ACCENT_2)
                add_card(slide, Inches(7.2), Inches(4.8), Inches(4.5), Inches(1.2), "Use Cases", "DeFi, NFTs, Supply Chain, Identity", ACCENT_3)

            elif slide_info["title"] == "Why Blockchain Matters":
                add_bullet_list(slide, Inches(1.0), Inches(1.5), Inches(5.6), Inches(4.8), [
                    "Eliminates single points of failure",
                    "Maintains a shared, tamper-resistant ledger",
                    "Facilitates trust in peer-to-peer systems",
                    "Improves transparency and auditability",
                    "Supports programmable workflows and digital assets"
                ], font_size=20)
                add_blockquote(slide, Inches(7.2), Inches(1.8), Inches(4.8), Inches(2.4), "Blockchain is a system for coordinating trust without relying on a central authority.")
                add_card(slide, Inches(7.2), Inches(4.6), Inches(4.8), Inches(1.4), "Key idea", "Consensus + cryptography + distributed validation = trust in code", ACCENT)

            elif slide_info["title"] == "Ethereum at a Glance":
                add_card(slide, Inches(1.0), Inches(1.6), Inches(3.4), Inches(1.9), "Ethereum", "A programmable blockchain that supports smart contracts and decentralized apps.", ACCENT)
                add_card(slide, Inches(4.8), Inches(1.6), Inches(3.4), Inches(1.9), "Smart Contracts", "Code that runs on-chain and enforces logic without a central server.", ACCENT_2)
                add_card(slide, Inches(8.6), Inches(1.6), Inches(3.4), Inches(1.9), "Ether (ETH)", "Native asset used to pay transaction fees and secure the network.", ACCENT_3)
                add_bullet_list(slide, Inches(1.0), Inches(4.2), Inches(11.0), Inches(2.0), [
                    "Open and permissionless network",
                    "Supports decentralized finance, DAOs, NFTs, identity, and supply chains",
                    "Runs state transitions through a global virtual machine"
                ], font_size=18)

            elif slide_info["title"] == "How Ethereum Works":
                add_card(slide, Inches(1.0), Inches(1.8), Inches(2.4), Inches(1.6), "Users", "Wallets, apps, and transactions", ACCENT)
                add_card(slide, Inches(3.8), Inches(1.8), Inches(2.4), Inches(1.6), "Nodes", "Validate and propagate blocks", ACCENT_2)
                add_card(slide, Inches(6.6), Inches(1.8), Inches(2.4), Inches(1.6), "Consensus", "Proof of Stake secures the network", ACCENT_3)
                add_card(slide, Inches(9.4), Inches(1.8), Inches(2.4), Inches(1.6), "State", "Updated ledger of balances and storage", GREEN)
                add_bullet_list(slide, Inches(1.4), Inches(4.1), Inches(10.5), Inches(2.4), [
                    "Accounts submit signed transactions",
                    "Validators verify block validity and ordering",
                    "The EVM executes smart contract logic and updates state"
                ], font_size=18)

            elif slide_info["title"] == "Transactions & Gas":
                add_card(slide, Inches(1.0), Inches(1.6), Inches(3.4), Inches(2.0), "Transaction", "A signed request to change blockchain state.", ACCENT)
                add_card(slide, Inches(5.0), Inches(1.6), Inches(3.4), Inches(2.0), "Gas", "Unit for computing work and setting network fees.", ACCENT_2)
                add_card(slide, Inches(9.0), Inches(1.6), Inches(3.4), Inches(2.0), "Block Time", "New blocks are proposed and finalized in intervals.", ACCENT_3)
                add_bullet_list(slide, Inches(1.3), Inches(4.4), Inches(10.8), Inches(2.2), [
                    "Gas price reflects demand and network congestion",
                    "The more computation, the more gas is required",
                    "Fees prevent spam and allocate resources fairly"
                ], font_size=18)

            elif slide_info["title"] == "Smart Contracts":
                add_blockquote(slide, Inches(1.0), Inches(1.8), Inches(5.3), Inches(2.2), "Smart contracts are code deployed to the blockchain that defines rules and executes automatically when triggered.")
                add_bullet_list(slide, Inches(1.2), Inches(4.5), Inches(5.3), Inches(2.4), [
                    "Automated execution",
                    "Deterministic behavior",
                    "Transparent rules",
                    "Immutable once deployed"
                ], font_size=18)
                add_card(slide, Inches(7.1), Inches(1.8), Inches(4.8), Inches(2.0), "Examples", "Token standards, auctions, voting, escrow, lending, governance", ACCENT_2)
                add_card(slide, Inches(7.1), Inches(4.3), Inches(4.8), Inches(1.8), "Trade-Off", "Trustless automation vs. complexity and auditability risk", ACCENT_3)

            elif slide_info["title"] == "Solidity Fundamentals":
                add_card(slide, Inches(0.9), Inches(1.6), Inches(3.7), Inches(1.6), "Pragmas", "Versioning and compiler specification", ACCENT)
                add_card(slide, Inches(4.9), Inches(1.6), Inches(3.7), Inches(1.6), "State Variables", "Persistent values stored on-chain", ACCENT_2)
                add_card(slide, Inches(8.9), Inches(1.6), Inches(3.7), Inches(1.6), "Functions", "Public, internal, payable, view, pure", ACCENT_3)
                add_blockquote(slide, Inches(1.2), Inches(3.8), Inches(10.8), Inches(2.2), "pragma solidity ^0.8.0; contract Token { uint256 totalSupply; function transfer(address to, uint256 amount) public {} }")
                add_bullet_list(slide, Inches(1.3), Inches(6.1), Inches(10.6), Inches(0.8), [
                    "Strongly typed language",
                    "Supports inheritance, events, libraries, and modifiers",
                    "Security discipline is essential"
                ], font_size=15)

            elif slide_info["title"] == "Remix IDE":
                add_bullet_list(slide, Inches(1.0), Inches(1.7), Inches(5.6), Inches(3.8), [
                    "Browser-based Solidity editor",
                    "Compile and deploy contracts quickly",
                    "Use test accounts and transaction logs",
                    "Ideal for learning, prototyping, and demos",
                    "Integrates with local and test networks"
                ], font_size=20)
                add_card(slide, Inches(7.4), Inches(1.8), Inches(4.5), Inches(1.6), "Workflow", "Write → Compile → Deploy → Test → Debug", ACCENT)
                add_card(slide, Inches(7.4), Inches(3.8), Inches(4.5), Inches(1.6), "Best For", "Fast prototyping and educational use", ACCENT_2)

            elif slide_info["title"] == "Local Development Workflow":
                add_card(slide, Inches(1.0), Inches(1.6), Inches(2.4), Inches(1.3), "1. Write", "Create the contract logic", ACCENT)
                add_card(slide, Inches(3.8), Inches(1.6), Inches(2.4), Inches(1.3), "2. Compile", "Check syntax and ABI", ACCENT_2)
                add_card(slide, Inches(6.6), Inches(1.6), Inches(2.4), Inches(1.3), "3. Test", "Use unit tests, scripts, or Remix", ACCENT_3)
                add_card(slide, Inches(9.4), Inches(1.6), Inches(2.4), Inches(1.3), "4. Deploy", "Run on local or test network", GREEN)
                add_blockquote(slide, Inches(1.5), Inches(3.5), Inches(10.3), Inches(1.8), "A professional workflow includes testing, security review, deployment automation, and monitoring.")
                add_bullet_list(slide, Inches(1.6), Inches(5.6), Inches(10.0), Inches(1.0), [
                    "Use scripts for repeated deployment",
                    "Validate state transitions before production"
                ], font_size=16)

            elif slide_info["title"] == "Ganache":
                add_bullet_list(slide, Inches(1.0), Inches(1.7), Inches(5.8), Inches(3.8), [
                    "Local Ethereum blockchain for development",
                    "Mints test accounts with private keys",
                    "Lets you test transactions and gas without public network costs",
                    "Useful for DApp development and smart contract debugging",
                    "Builds quick feedback loops for iteration"
                ], font_size=20)
                add_card(slide, Inches(7.4), Inches(1.8), Inches(4.5), Inches(1.8), "Why it matters", "Faster local testing and easier debugging before deployment", ACCENT)
                add_card(slide, Inches(7.4), Inches(4.0), Inches(4.5), Inches(1.6), "Typical Use", "Smart contract testing, local frontend validation, wallet integration", ACCENT_2)

            elif slide_info["title"] == "DApp Architecture":
                add_card(slide, Inches(1.0), Inches(1.6), Inches(3.3), Inches(1.8), "Frontend", "UI, wallet, user flows", ACCENT)
                add_card(slide, Inches(4.9), Inches(1.6), Inches(3.3), Inches(1.8), "Web3 Layer", "RPC, providers, contract ABIs", ACCENT_2)
                add_card(slide, Inches(8.8), Inches(1.6), Inches(3.3), Inches(1.8), "Blockchain", "Smart contracts and state transitions", ACCENT_3)
                add_bullet_list(slide, Inches(1.2), Inches(4.1), Inches(11.0), Inches(2.1), [
                    "Users interact via browser or mobile interfaces",
                    "The wallet signs transactions using private keys",
                    "The blockchain stores logic, state, and canonical ownership"
                ], font_size=18)

            elif slide_info["title"] == "Wallets and Identity":
                add_bullet_list(slide, Inches(1.0), Inches(1.8), Inches(5.8), Inches(3.9), [
                    "Wallets hold private keys and provide access to accounts",
                    "MetaMask is a common browser wallet for Ethereum",
                    "Signatures authorize transactions without revealing secret keys",
                    "DApps use public addresses as digital identity anchors",
                    "Security depends on key management and credential hygiene"
                ], font_size=19)
                add_card(slide, Inches(7.5), Inches(1.8), Inches(4.3), Inches(1.6), "User Identity", "Public address + signed intent = identity on-chain", ACCENT)
                add_card(slide, Inches(7.5), Inches(4.0), Inches(4.3), Inches(1.6), "Best Practice", "Never share seed phrases or private keys", RED)

            elif slide_info["title"] == "Use Case: DeFi":
                add_blockquote(slide, Inches(1.1), Inches(1.7), Inches(5.3), Inches(2.3), "DeFi removes centralized intermediaries from lending, trading, and payments using smart contract automation.")
                add_bullet_list(slide, Inches(1.3), Inches(4.2), Inches(5.7), Inches(2.4), [
                    "Lending and borrowing",
                    "Automated market makers",
                    "Stablecoins and yield products",
                    "Cross-border settlement"
                ], font_size=19)
                add_card(slide, Inches(7.3), Inches(1.8), Inches(4.8), Inches(1.8), "Benefits", "Open access, composability, 24/7 operations", ACCENT)
                add_card(slide, Inches(7.3), Inches(4.0), Inches(4.8), Inches(1.6), "Risks", "Smart contract risk, liquidation events, volatility", ACCENT_3)

            elif slide_info["title"] == "Use Case: NFTs":
                add_bullet_list(slide, Inches(1.0), Inches(1.8), Inches(5.8), Inches(3.8), [
                    "Represent ownership of unique digital or physical assets",
                    "Powered by token standards such as ERC-721 and ERC-1155",
                    "Enable collectibles, art, gaming, and identity credentials",
                    "Help create provable scarcity and digital provenance",
                    "Connect creators to programmable monetization"
                ], font_size=19)
                add_card(slide, Inches(7.5), Inches(1.8), Inches(4.2), Inches(1.8), "Use Cases", "Art, gaming, tickets, memberships, digital collectibles", ACCENT)
                add_card(slide, Inches(7.5), Inches(4.0), Inches(4.2), Inches(1.5), "Core Value", "Authenticity and transferable ownership", ACCENT_2)

            elif slide_info["title"] == "Use Case: Supply Chain":
                add_bullet_list(slide, Inches(1.0), Inches(1.8), Inches(5.7), Inches(3.9), [
                    "Track products from source to customer",
                    "Verify origin and authenticity of goods",
                    "Reduce fraud and improve product accountability",
                    "Support audits and compliance documentation",
                    "Create transparent, shared operational records"
                ], font_size=19)
                add_card(slide, Inches(7.4), Inches(1.8), Inches(4.5), Inches(1.9), "Industry Example", "Food safety, luxury goods, pharmaceuticals, logistics", ACCENT)
                add_card(slide, Inches(7.4), Inches(4.0), Inches(4.5), Inches(1.5), "Impact", "Greater trust and operational visibility", ACCENT_3)

            elif slide_info["title"] == "Security Best Practices":
                add_bullet_list(slide, Inches(1.0), Inches(1.7), Inches(5.8), Inches(3.8), [
                    "Follow secure coding patterns and style guides",
                    "Use checks-effects-interactions discipline",
                    "Validate inputs and guard against reentrancy",
                    "Perform threat modeling and security reviews",
                    "Test extensively before mainnet deployment"
                ], font_size=19)
                add_card(slide, Inches(7.4), Inches(1.8), Inches(4.5), Inches(1.8), "Common Attacks", "Reentrancy, overflow, access control flaws", RED)
                add_card(slide, Inches(7.4), Inches(4.0), Inches(4.5), Inches(1.5), "Tooling", "Static analysis, fuzzing, audits, testnets", ACCENT_2)

            elif slide_info["title"] == "Challenges & Scalability":
                add_card(slide, Inches(1.0), Inches(1.8), Inches(3.4), Inches(2.0), "Scalability", "Blockchains must process more transactions efficiently.", ACCENT)
                add_card(slide, Inches(4.8), Inches(1.8), Inches(3.4), Inches(2.0), "Cost", "Gas fees can rise during network congestion.", ACCENT_2)
                add_card(slide, Inches(8.6), Inches(1.8), Inches(3.4), Inches(2.0), "UX", "Wallet setup and transaction flows still challenge mainstream adoption.", ACCENT_3)
                add_bullet_list(slide, Inches(1.4), Inches(4.5), Inches(10.6), Inches(2.0), [
                    "Layer 2 solutions help reduce burden on mainnet",
                    "Shared infrastructure and standards improve interoperability",
                    "Trade-offs continue between decentralization, security, and speed"
                ], font_size=18)

            elif slide_info["title"] == "Future of Ethereum":
                add_bullet_list(slide, Inches(1.0), Inches(1.7), Inches(5.8), Inches(3.8), [
                    "Layer 2 scaling is increasing throughput and reducing cost",
                    "Account abstraction may improve wallet and UX experiences",
                    "Ethereum remains central to DeFi, NFTs, identity, and governance",
                    "New tooling and developer ecosystems continue to expand",
                    "The next wave will prioritize usability and interoperability"
                ], font_size=19)
                add_card(slide, Inches(7.4), Inches(1.8), Inches(4.5), Inches(1.8), "Emerging Trends", "Layer 2s, modular chains, smarter tooling", ACCENT)
                add_card(slide, Inches(7.4), Inches(4.0), Inches(4.5), Inches(1.5), "Opportunity", "Builders can create open, user-owned digital systems", ACCENT_2)

            elif slide_info["title"] == "Thank You":
                add_textbox(slide, Inches(2.0), Inches(2.2), Inches(9.0), Inches(0.8), "Thank You", font_size=32, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
                add_textbox(slide, Inches(2.5), Inches(3.2), Inches(8.0), Inches(0.8), "Questions & Discussion", font_size=22, color=MUTED, align=PP_ALIGN.CENTER)
                add_card(slide, Inches(3.2), Inches(4.6), Inches(6.6), Inches(1.2), "Ethereum + Smart Contracts + DApps", "Building trustless systems for the next generation of digital experiences", ACCENT)

    prs.save(OUTPUT_FILE)


if __name__ == "__main__":
    build_deck()
    print(f"Created {OUTPUT_FILE}")
