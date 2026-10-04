"""Visual components for the bilingual petty cash workspace."""
from html import escape
from urllib.parse import quote

from i18n import tr

ICON_PATHS = {
    "receipt": '<path d="M6 3h12v18l-3-2-3 2-3-2-3 2V3Z"/><path d="M9 7h6M9 11h6M9 15h3"/>',
    "upload": '<path d="M12 16V3m-5 5 5-5 5 5M4 15v5h16v-5"/>',
    "edit": '<path d="m15 5 4 4M4 20l4-1L20 7a3 3 0 0 0-4-4L4 15v5Z"/>',
    "chart": '<path d="M4 3v17h17M8 16v-4m5 4V7m5 9V4"/>',
    "download": '<path d="M12 3v13m-5-5 5 5 5-5M4 17v4h16v-4"/>',
    "check": '<path d="m5 12 4 4L19 6"/>',
}


def icon(name: str) -> str:
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{ICON_PATHS[name]}</svg>'


def icon_url(name: str) -> str:
    return "data:image/svg+xml," + quote(icon(name).replace("currentColor", "black"), safe="")


def brand_header(language: str) -> str:
    return f'<div class="app-heading"><div class="brand-row"><div class="brand-mark">{icon("receipt")}</div><div><h1>{escape(tr(language,"app_title"))}</h1><p>{escape(tr(language,"header_subtitle"))}</p></div></div></div>'


def section_header(language: str, key: str, symbol: str) -> str:
    return f'<div class="section-heading"><span class="section-icon">{icon(symbol)}</span><h3>{escape(tr(language,key))}</h3></div>'


def workspace_intro(language: str) -> str:
    return f'<div class="workspace-intro"><div class="eyebrow">{escape(tr(language,"workspace_label"))}</div><h2>{escape(tr(language,"workspace_title"))} <span class="intro-accent">{escape(tr(language,"workspace_accent"))}</span></h2><p>{escape(tr(language,"workspace_intro"))}</p></div>'


def workspace_summary(language: str, count: int, amount: str) -> str:
    return f'<div class="workspace-summary"><div><small>{escape(tr(language,"workspace_count"))}</small><strong>{count:02d}</strong></div><div class="summary-total"><small>{escape(tr(language,"workspace_total"))}</small><strong>{escape(amount)} <span>Bs</span></strong></div></div>'


def workflow_guide(language: str) -> str:
    receipt = '<svg viewBox="0 0 190 115" fill="none" aria-hidden="true"><ellipse cx="96" cy="107" rx="63" ry="6" fill="#111B3C" opacity=".35"/><circle cx="24" cy="31" r="3" fill="#9E9AFD" opacity=".45"/><circle cx="166" cy="16" r="6" stroke="#9E9AFD" opacity=".3"/><rect x="47" y="6" width="98" height="97" rx="10" fill="#F9FAFF" transform="rotate(-6 96 55)"/><path d="M65 27h57M65 40h38M65 54h56M65 67h27" stroke="#CED4ED" stroke-width="5" stroke-linecap="round"/><path d="M65 83h34" stroke="#7770EF" stroke-width="7" stroke-linecap="round"/><circle cx="144" cy="78" r="20" fill="#4DD6B2" stroke="#243363" stroke-width="5"/><path d="m136 78 5 5 10-11" stroke="#163D38" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/></svg>'
    steps = ''.join(f'<div class="guide-step">{icon(symbol)}<span>{escape(tr(language,key))}</span></div>' for symbol,key in [('upload','guide_read'),('check','guide_review'),('download','guide_export')])
    return f'<aside class="workflow-guide"><div class="eyebrow">{escape(tr(language,"guide_label"))}</div><div class="receipt-visual">{receipt}</div><h3>{escape(tr(language,"guide_title"))}</h3>{steps}<p class="guide-note">{escape(tr(language,"guide_note"))}</p></aside>'


def app_css(dark_mode: bool) -> str:
    grid_filter = "invert(0.90) hue-rotate(180deg) brightness(1.08)" if dark_mode else "none"
    if dark_mode:
        colors = {
            "bg": "#0F172A", "surface": "#19243D", "sidebar": "#111827",
            "text": "#F1F5F9", "muted": "#B5C1D7", "border": "#34425E",
            "input": "#202D48", "primary": "#A5B4FC", "accent": "#22C55E",
            "scheme": "dark",
        }
    else:
        colors = {
            "bg": "#F5F6FB", "surface": "#FFFFFF", "sidebar": "#F8FAFC",
            "text": "#19213A", "muted": "#5B6680", "border": "#E3E7F1",
            "input": "#F8FAFC", "primary": "#5148E8", "accent": "#16A34A",
            "scheme": "light",
        }
    return f"""
    <style>
    :root {{
      --app-bg:{colors['bg']}; --app-surface:{colors['surface']}; --app-sidebar:{colors['sidebar']};
      --app-text:{colors['text']}; --app-muted:{colors['muted']}; --app-border:{colors['border']};
      --app-input:{colors['input']}; --app-primary:{colors['primary']}; --app-accent:{colors['accent']};
      --background-color:{colors['bg']}; --secondary-background-color:{colors['surface']};
      --text-color:{colors['text']}; --secondary-text-color:{colors['muted']};
      --primary-color:{colors['primary']}; --border-color:{colors['border']}; color-scheme:{colors['scheme']};
    }}
    #MainMenu, footer {{visibility:hidden;}}
    header[data-testid="stHeader"] {{display:none;}}
    .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {{background:var(--app-bg)!important;color:var(--app-text)!important;}}
    .block-container {{padding:1.65rem 2rem 1rem;max-width:1120px;}}
    [data-testid="stVerticalBlock"] {{gap:.85rem;}}
    .app-heading h1 {{font-size:1.65rem!important;font-weight:700;letter-spacing:-.04em;margin:0;padding:0;line-height:1.2;}}
    .app-heading p {{font-size:.86rem;margin:.35rem 0 0;color:var(--app-muted)!important;}}
    .st-key-app_header {{padding-bottom:.6rem;}}
    .st-key-app_header [data-testid="stHorizontalBlock"] {{align-items:center;}}
    .st-key-app_header [data-testid="stRadio"] {{margin:0;}}
    [data-testid="stTabs"] [role="tablist"] {{gap:1.5rem;border-bottom:1px solid var(--app-border);}}
    [data-testid="stTabs"] [role="tab"] {{padding:.65rem 0;height:auto;}}
    [data-testid="stTabs"] [data-baseweb="tab-panel"] {{padding-top:1.2rem;}}
    .stApp h3 {{font-size:1.15rem!important;letter-spacing:-.02em;padding-top:0;}}
    .st-key-receipt_panel, .st-key-manual_panel {{background:var(--app-surface);border:1px solid var(--app-border);border-radius:12px;padding:1.3rem;}}
    [data-testid="stForm"] {{border:0;padding:0;}}
    [data-testid="stFileUploaderDropzone"] {{min-height:125px;border:1px dashed var(--app-border);border-radius:10px;padding:1.2rem;}}
    [data-testid="stAlert"] {{padding:.6rem .8rem!important;border-radius:8px!important;}}
    [data-testid="stMetric"] {{padding:.75rem!important;border-radius:10px!important;}}
    [data-testid="stMetricValue"] {{font-size:1.65rem;}}
    .st-key-session_footer {{border-top:1px solid var(--app-border);padding-top:.55rem;margin-top:1rem;}}
    @media (max-width:640px) {{
      .block-container {{padding:1rem 1rem .8rem;}}
      .st-key-app_header [data-testid="stHorizontalBlock"] {{flex-wrap:wrap;gap:.5rem!important;}}
      .st-key-app_header [data-testid="stColumn"] {{flex:1 1 40%;min-width:0!important;width:auto!important;}}
      .st-key-app_header [data-testid="stColumn"]:first-child {{flex:1 1 100%;}}
      [data-testid="stTabs"] [role="tablist"] {{gap:1rem;}}
      [data-testid="stTabs"] [role="tab"] p {{font-size:.8rem;}}
      .st-key-receipt_panel, .st-key-manual_panel {{padding:1rem;}}
      [data-testid="stFileUploaderDropzone"] {{padding:.75rem;}}
    }}
    .stApp, .stApp p, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp [data-testid="stMarkdownContainer"] {{color:var(--app-text);}}
    .stApp [data-testid="stCaptionContainer"] {{color:var(--app-muted);}}
    [data-testid="stTooltipIcon"], [data-testid="stTooltipIcon"] svg {{color:var(--app-muted)!important;fill:var(--app-muted)!important;}}
    .stButton > button, .stDownloadButton > button {{border-radius:8px;font-weight:600;padding:.5rem .9rem;border:1px solid var(--app-border);background:var(--app-surface);color:var(--app-text);}}
    .stButton > button[kind="primary"], .stDownloadButton > button {{background:#5148E8;color:#FFFFFF;border-color:#5148E8;}}
    .stButton > button[kind="primary"] p, .stDownloadButton > button p {{color:#FFFFFF!important;}}
    .stButton > button:hover {{border-color:var(--app-primary);}}
    .stButton > button:disabled, .stDownloadButton > button:disabled {{background:var(--app-input)!important;color:var(--app-muted)!important;border:1px solid var(--app-border)!important;opacity:1!important;}}
    .stButton > button:disabled p, .stDownloadButton > button:disabled p {{color:var(--app-muted)!important;}}
    [data-testid="stMetric"], [data-testid="stVerticalBlockBorderWrapper"], [data-testid="stAlert"] {{background:var(--app-surface)!important;color:var(--app-text)!important;border:1px solid var(--app-border);border-radius:14px;padding:12px;}}
    [data-testid="stTextInputRootElement"], [data-testid="stNumberInputContainer"], [data-baseweb="input"] > div, [data-baseweb="select"] > div, textarea {{background:var(--app-input)!important;color:var(--app-text)!important;border-color:var(--app-border)!important;}}
    input, textarea, [data-baseweb="input"] input {{color:var(--app-text)!important;caret-color:var(--app-text);}}
    input::placeholder, textarea::placeholder {{color:var(--app-muted)!important;}}
    [data-testid="stFileUploaderDropzone"], [data-testid="stDataEditor"], [data-testid="stDataFrame"], [data-testid="stTable"] {{background:var(--app-surface)!important;color:var(--app-text)!important;border-color:var(--app-border)!important;}}
    [data-testid="stDataEditor"] *, [data-testid="stDataFrame"] *, [data-testid="stTable"] * {{color:var(--app-text);}}
    [data-testid="stFileUploaderDropzone"] button {{background:var(--app-input)!important;color:var(--app-text)!important;border:1px solid var(--app-border)!important;}}
    [role="tablist"] button {{color:var(--app-text)!important;}}
    [data-testid="stDataEditor"], [data-testid="stDataFrame"] {{--gdg-bg-cell:var(--app-surface);--gdg-bg-cell-medium:var(--app-input);--gdg-bg-header:var(--app-input);--gdg-bg-header-has-focus:var(--app-input);--gdg-bg-header-hovered:var(--app-input);--gdg-text-dark:var(--app-text);--gdg-text-medium:var(--app-muted);--gdg-text-light:var(--app-muted);--gdg-text-header:var(--app-text);--gdg-border-color:var(--app-border);--gdg-accent-color:var(--app-primary);--gdg-accent-light:color-mix(in srgb,var(--app-primary) 20%,transparent);}}
    .stDataFrameGlideDataEditor {{--gdg-bg-cell:var(--app-surface)!important;--gdg-bg-cell-medium:var(--app-input)!important;--gdg-bg-header:var(--app-input)!important;--gdg-bg-header-has-focus:var(--app-input)!important;--gdg-bg-header-hovered:var(--app-input)!important;--gdg-bg-group-header:var(--app-input)!important;--gdg-bg-header-top-left:var(--app-input)!important;--gdg-text-dark:var(--app-text)!important;--gdg-text-medium:var(--app-muted)!important;--gdg-text-light:var(--app-muted)!important;--gdg-text-bubble:var(--app-muted)!important;--gdg-text-header:var(--app-text)!important;--gdg-text-group-header:var(--app-text)!important;--gdg-bg-icon-header:var(--app-muted)!important;--gdg-fg-icon-header:var(--app-surface)!important;--gdg-border-color:var(--app-border)!important;--gdg-horizontal-border-color:var(--app-border)!important;--gdg-drilldown-border:var(--app-border)!important;--gdg-accent-color:var(--app-primary)!important;--gdg-accent-fg:#FFFFFF!important;--gdg-accent-light:color-mix(in srgb,var(--app-primary) 20%,transparent)!important;--gdg-link-color:var(--app-primary)!important;}}
    [data-testid="stDataEditor"] canvas, [data-testid="stDataFrame"] canvas {{color-scheme:{colors['scheme']};}}
    [data-testid="stDataFrameResizable"] {{border-color:var(--app-border)!important;}}
    .stDataFrameGlideDataEditor canvas {{filter:{grid_filter};}}
    [data-testid="stVegaLiteChart"] svg {{filter:none!important;background-color:var(--app-bg)!important;}}
    [data-testid="stVegaLiteChart"] svg text {{fill:var(--app-muted)!important;}}
    [data-testid="stVegaLiteChart"] svg .role-axis-grid line, [data-testid="stVegaLiteChart"] svg .role-axis-domain {{stroke:var(--app-border)!important;}}
    [data-testid="stElementToolbarButtonContainer"] {{background:var(--app-surface)!important;color:var(--app-text)!important;}}
    [data-testid="stElementToolbarButtonContainer"] button {{background:var(--app-input)!important;color:var(--app-text)!important;border:1px solid var(--app-border)!important;}}
    [data-testid="stExpander"] details > summary {{background:var(--app-surface)!important;color:var(--app-text)!important;}}
    [data-testid="stExpander"] summary [data-testid="stMarkdownContainer"], [data-testid="stExpander"] summary p {{color:var(--app-text)!important;}}
    [data-testid="stExpander"] [data-testid="stExpanderDetails"] {{background:var(--app-surface)!important;color:var(--app-text)!important;}}
    [data-testid="stDateInputField"] {{background:var(--app-input)!important;color:var(--app-text)!important;border:1px solid var(--app-border)!important;}}
    [data-testid="stDateInputField"] > div, [data-testid="stDateInputField"] [role="group"] {{background:var(--app-input)!important;color:var(--app-text)!important;}}
    [data-testid="stDateInputField"] span {{color:var(--app-text)!important;}}
    [data-testid="stBaseLinkButton-secondary"], [data-testid="stBaseLinkButton-primary"] {{background:var(--app-input)!important;color:var(--app-text)!important;border:1px solid var(--app-border)!important;}}
    [data-testid="stAlert"] [data-testid="stMarkdownContainer"], [data-testid="stAlert"] p {{color:var(--app-text)!important;}}
    [data-testid="stFileUploaderDropzone"] *, [data-testid="stFileUploaderDropzone"] button {{color:var(--app-text)!important;}}
    [data-testid="stFileChip"] {{background:var(--app-input)!important;color:var(--app-text)!important;border:1px solid var(--app-border)!important;}}
    [data-testid="stFileChip"] * {{color:var(--app-text)!important;}}
    [data-testid="stBaseButton-secondaryFormSubmit"] {{background:var(--app-input)!important;color:var(--app-text)!important;border:1px solid var(--app-border)!important;}}
    [data-testid="stBaseButton-secondaryFormSubmit"]:disabled {{background:var(--app-input)!important;color:var(--app-muted)!important;border:1px solid var(--app-border)!important;opacity:1!important;}}
    [data-baseweb="popover"] *, [data-baseweb="menu"] * {{color:var(--app-text);}}
    [data-testid="stTabs"] [data-baseweb="tab-highlight"] {{background-color:var(--app-primary)!important;}}

    :root {{--app-soft:color-mix(in srgb,var(--app-primary) 9%,var(--app-surface));--app-shadow:0 12px 38px rgba(30,38,77,.055);}}
    [data-testid="stMain"] {{background:radial-gradient(ellipse at 20% 0%,color-mix(in srgb,var(--app-primary) 5%,transparent),transparent 48%),var(--app-bg)!important;}}
    .block-container {{max-width:1180px;padding:1.35rem 2rem 1rem;}}
    .st-key-app_header {{padding-bottom:.8rem;}}
    .brand-row {{display:flex;align-items:center;gap:12px;}}
    .brand-mark {{width:43px;height:43px;flex:none;display:grid;place-items:center;background:linear-gradient(145deg,#625AF2,#4338CA);border-radius:13px;box-shadow:0 6px 16px rgba(79,70,229,.18);}}
    .brand-mark svg {{width:25px;height:25px;color:white;}}
    .app-heading h1 {{font-size:1.25rem!important;letter-spacing:-.035em;}}
    .app-heading p {{font-size:.73rem;margin:.22rem 0 0;}}
    .brand-badge {{font-size:9px;font-weight:700;letter-spacing:.09em;text-transform:uppercase;padding:4px 6px;border-radius:5px;color:var(--app-muted);background:var(--app-soft);margin-left:8px;vertical-align:middle;}}
    .st-key-app_header [data-testid="stRadioGroup"] {{display:flex;gap:2px;background:var(--app-surface);border:1px solid var(--app-border);border-radius:10px;padding:4px;}}
    .st-key-app_header [data-testid="stRadioOption"] {{padding:5px 8px!important;margin:0!important;border-radius:6px;}}
    .st-key-app_header [data-testid="stRadioOption"] p {{font-size:.75rem;font-weight:600;}}
    .st-key-app_header [data-testid="stRadioOption"]:has(input:checked) {{background:var(--app-soft);}}
    .st-key-app_header [data-testid="stRadioOption"] > div > div > div:first-child:not([data-testid]) {{display:none;}}
    .st-key-app_header [data-testid="stCheckbox"] {{background:var(--app-surface);border:1px solid var(--app-border);border-radius:10px;padding:10px 12px;}}
    .st-key-app_header [data-testid="stCheckbox"] p {{font-size:.75rem;white-space:nowrap;}}
    [data-testid="stTabs"] [role="tablist"] {{gap:7px;padding:6px;border:1px solid var(--app-border);border-radius:13px;background:color-mix(in srgb,var(--app-surface) 60%,var(--app-bg));}}
    [data-testid="stTabs"] [role="tab"] {{flex:1;justify-content:center;gap:10px;padding:12px 16px;border-radius:9px;transition:background .15s,box-shadow .15s;}}
    [data-testid="stTabs"] [role="tab"] p {{font-size:.82rem;font-weight:500;color:var(--app-muted)!important;}}
    [data-testid="stTabs"] [role="tab"][aria-selected="true"] {{background:var(--app-surface);box-shadow:0 2px 8px rgba(30,38,77,.075);}}
    [data-testid="stTabs"] [role="tab"][aria-selected="true"] p {{font-weight:650;color:var(--app-primary)!important;}}
    [data-testid="stTabs"] [data-baseweb="tab-highlight"], [data-testid="stTabs"] [data-baseweb="tab-border"] {{display:none;}}
    [data-testid="stTabs"] [role="tab"]::before {{content:"";width:16px;height:16px;display:block;flex:none;background:var(--app-muted);mask:url('{icon_url("upload")}') center/contain no-repeat;}}
    [data-testid="stTabs"] [role="tab"]:nth-of-type(2)::before {{mask-image:url('{icon_url("edit")}');}}
    [data-testid="stTabs"] [role="tab"]:nth-of-type(3)::before {{mask-image:url('{icon_url("chart")}');}}
    [data-testid="stTabs"] [role="tab"]:nth-of-type(4)::before {{mask-image:url('{icon_url("download")}');}}
    [data-testid="stTabs"] [role="tab"][aria-selected="true"]::before {{background:var(--app-primary);}}
    [data-testid="stTabs"] [data-baseweb="tab-panel"] {{padding-top:1.2rem;}}
    .workspace-intro {{padding:3px 0 0;}}
    .eyebrow {{font-size:.63rem;font-weight:750;letter-spacing:.14em;color:var(--app-muted);text-transform:uppercase;}}
    .workspace-intro h2 {{font-size:1.85rem!important;letter-spacing:-.05em;line-height:1.25;margin:7px 0 8px;padding:0;color:var(--app-text)!important;}}
    .workspace-intro h2 .intro-accent {{color:var(--app-primary);}}
    .workspace-intro p {{font-size:.84rem;color:var(--app-muted)!important;margin:0;line-height:1.55;}}
    .workspace-summary {{display:flex;align-items:center;justify-content:space-between;gap:16px;background:var(--app-surface);border:1px solid var(--app-border);border-radius:13px;padding:15px 18px;box-shadow:var(--app-shadow);}}
    .workspace-summary small {{display:block;color:var(--app-muted);font-size:.62rem;text-transform:uppercase;letter-spacing:.08em;}}
    .workspace-summary strong {{display:block;color:var(--app-text);font-size:1.3rem;font-weight:650;margin-top:6px;font-variant-numeric:tabular-nums;}}
    .workspace-summary .summary-total {{border-left:1px solid var(--app-border);padding-left:18px;text-align:right;}}
    .workspace-summary .summary-total strong {{font-size:1.05rem;}}
    .summary-total strong span {{font-size:.65rem;font-weight:500;color:var(--app-muted);}}
    .st-key-input_switch {{padding:6px 0 4px;}}
    .st-key-input_switch [data-testid="stRadioGroup"] {{width:fit-content;gap:3px;padding:4px;background:color-mix(in srgb,var(--app-border) 55%,var(--app-bg));border-radius:11px;}}
    .st-key-input_switch [data-testid="stRadioOption"] {{border-radius:8px;padding:8px 16px!important;margin:0!important;transition:background .15s;}}
    .st-key-input_switch [data-testid="stRadioOption"] p {{font-size:.8rem;font-weight:550;color:var(--app-muted)!important;}}
    .st-key-input_switch [data-testid="stRadioOption"]:has(input:checked) {{background:var(--app-surface);box-shadow:0 2px 5px rgba(30,38,77,.08);}}
    .st-key-input_switch [data-testid="stRadioOption"]:has(input:checked) p {{color:var(--app-primary)!important;}}
    .st-key-input_switch [data-testid="stRadioOption"] > div > div > div:first-child:not([data-testid]) {{display:none;}}
    [data-testid="stRadioOption"]:focus-within {{outline:2px solid var(--app-primary);outline-offset:2px;}}
    .st-key-receipt_panel, .st-key-manual_panel, .st-key-report_panel {{border-radius:17px;padding:22px 24px;background:var(--app-surface);border:1px solid var(--app-border);box-shadow:var(--app-shadow);}}
    .section-heading {{display:flex;align-items:center;gap:10px;}}
    .section-heading .section-icon {{width:34px;height:34px;display:grid;place-items:center;border-radius:10px;background:var(--app-soft);color:var(--app-primary);flex:none;}}
    .section-icon svg {{width:19px;height:19px;}}
    .section-heading h3 {{margin:0;font-size:1.1rem!important;letter-spacing:-.025em;padding:0;}}
    [data-testid="stFileUploaderDropzone"] {{min-height:142px;padding:20px 16px;border-radius:12px;background:linear-gradient(135deg,var(--app-soft),var(--app-input))!important;border:1px dashed color-mix(in srgb,var(--app-primary) 30%,var(--app-border));}}
    [data-testid="stFileUploaderDropzone"] button {{border-radius:8px!important;background:var(--app-surface)!important;font-weight:600;box-shadow:0 2px 4px rgba(30,38,77,.035);}}
    [data-testid="stFileUploaderDropzone"] small {{font-size:.7rem;color:var(--app-muted)!important;}}
    .workflow-guide {{border-radius:17px;padding:24px;background:radial-gradient(circle at 95% 0%,rgba(145,139,255,.25),transparent 54%),linear-gradient(145deg,#24245B,#1C2D52);color:#F5F6FF;min-height:324px;box-sizing:border-box;overflow:hidden;}}
    .workflow-guide .eyebrow {{color:#B6BDEA;font-size:.59rem;}}
    .receipt-visual {{height:94px;display:flex;justify-content:center;margin:4px 0 8px;}}
    .receipt-visual svg {{width:172px;height:104px;}}
    .workflow-guide h3 {{color:#F5F6FF!important;font-size:1.12rem!important;letter-spacing:-.025em;margin:0 0 14px;padding:0;}}
    .guide-step {{display:flex;align-items:center;gap:10px;margin:11px 0;font-size:.76rem;color:#DDE2F5;}}
    .guide-step svg {{width:16px;height:16px;color:#AEBBFF;flex:none;}}
    .workflow-guide .guide-note {{font-size:.68rem;color:#B8C2DF!important;margin:15px 0 0;line-height:1.55;border-top:1px solid rgba(255,255,255,.1);padding-top:12px;}}
    .stButton > button, [data-testid="stDownloadButton"] button {{border-radius:9px;min-height:41px;font-size:.82rem;}}
    .stButton > button[kind="primary"], [data-testid="stDownloadButton"] button {{background:#5148E8!important;color:white!important;border:1px solid #5148E8!important;box-shadow:0 4px 10px rgba(79,70,229,.13);}}
    .stButton > button[kind="primary"] p, [data-testid="stDownloadButton"] button p {{color:white!important;}}
    .stButton > button[kind="primary"]:hover, [data-testid="stDownloadButton"] button:hover {{background:#4338CA!important;border-color:#4338CA!important;}}
    .stButton > button:disabled {{box-shadow:none!important;background:var(--app-input)!important;border-color:var(--app-border)!important;}}
    [data-testid="stTextInput"] label p {{font-size:.75rem!important;font-weight:550;}}
    [data-testid="stTextInputRootElement"] {{border-radius:9px!important;}}
    [data-testid="stExpander"] {{border-radius:10px;}}
    [data-testid="stMetric"] {{background:var(--app-surface)!important;border:1px solid var(--app-border);box-shadow:var(--app-shadow);border-radius:13px!important;padding:15px 17px!important;}}
    [data-testid="stMetricLabel"] p {{font-size:.65rem!important;text-transform:uppercase;letter-spacing:.075em;color:var(--app-muted)!important;}}
    [data-testid="stMetricValue"] {{font-size:1.7rem;letter-spacing:-.04em;font-weight:650;font-variant-numeric:tabular-nums;}}
    [data-testid="stDataFrame"] {{border-radius:13px!important;overflow:hidden;box-shadow:var(--app-shadow);}}
    [data-testid="stVegaLiteChart"] {{border:1px solid var(--app-border);border-radius:15px;padding:16px;background:var(--app-surface)!important;}}
    [data-testid="stVegaLiteChart"] svg {{background:var(--app-surface)!important;}}
    .st-key-session_footer {{margin-top:.6rem;padding-top:.6rem;}}
    .st-key-session_footer p {{font-size:.69rem!important;}}
    @media(max-width:760px) {{
      .block-container {{padding:1rem;}}
      .workspace-intro h2 {{font-size:1.5rem!important;}}
      .workspace-summary {{padding:12px 15px;}}
      [data-testid="stTabs"] [role="tablist"] {{gap:3px;padding:4px;}}
      [data-testid="stTabs"] [role="tab"] {{padding:10px 5px;gap:5px;}}
      [data-testid="stTabs"] [role="tab"]::before {{display:none;}}
      [data-testid="stTabs"] [role="tab"] p {{font-size:.72rem;}}
      .st-key-receipt_panel, .st-key-manual_panel, .st-key-report_panel {{padding:18px;border-radius:14px;}}
      .workflow-guide {{min-height:0;padding:20px;}}
      .workflow-guide {{display:none;}}
      .receipt-visual {{display:none;}}
      .workflow-guide .guide-note {{margin-top:10px;}}
      .st-key-input_switch [data-testid="stRadioOption"] {{padding:8px 12px!important;}}
    }}

    [data-testid="stTabs"] .react-aria-SelectionIndicator {{display:none;}}
    .stButton > button[kind="primary"]:disabled p {{color:var(--app-muted)!important;}}
    @media(min-width:761px) {{
      .st-key-app_header [data-testid="stColumn"]:first-child {{flex:1 1 auto!important;width:auto!important;}}
      .st-key-app_header [data-testid="stColumn"]:nth-child(2) {{flex:0 0 90px!important;width:90px!important;}}
      .st-key-app_header [data-testid="stColumn"]:nth-child(3) {{flex:0 0 145px!important;width:145px!important;}}
    }}
    </style>
    """

