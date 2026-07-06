MAIN_CSS = """
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body, [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #f5f7fa 0%, #e8ecf1 100%);
        font-family: 'Segoe UI', 'Arial', sans-serif;
    }
    [data-testid="stSidebar"] {
        background: #FFFFFF;
        color: #333333;
        border-right: 1px solid #E0E0E0;
    }
    [data-testid="stSidebar"] > div:first-child { background: transparent !important; }
    [data-testid="stSidebar"] .stMarkdown { color: #333333; }
    [data-testid="stSidebar"] h3 {
        color: #1565C0;
        font-weight: 700;
        letter-spacing: 0.05em;
        margin-top: 16px;
        margin-bottom: 10px;
    }
    [data-testid="stSidebar"] .stSelectbox > div > div,
    [data-testid="stSidebar"] .stNumberInput input {
        background: #F5F5F5 !important;
        color: #333333 !important;
        border: 1px solid #CCCCCC !important;
        border-radius: 6px !important;
    }
    .header-bar {
        background: linear-gradient(90deg, #0D2340 0%, #1565C0 100%);
        border-radius: 10px;
        padding: 22px 28px;
        margin-bottom: 22px;
        box-shadow: 0 4px 15px rgba(13, 35, 64, 0.2);
    }
    .header-bar h1 {
        color: #FFFFFF;
        font-size: 1.5rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.02em;
    }
    .header-bar p {
        color: #90CAF9;
        font-size: 0.85rem;
        margin: 6px 0 0 0;
    }
    .metric-card {
        background: linear-gradient(135deg, #FFFFFF 0%, #F5F9FC 100%);
        border: 1px solid rgba(200, 210, 225, 0.5);
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        margin-bottom: 10px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
        transition: all 0.3s ease;
    }
    .metric-card:hover {
        box-shadow: 0 4px 16px rgba(13, 35, 64, 0.1);
        transform: translateY(-2px);
    }
    .metric-card .label {
        font-size: 0.70rem;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        font-weight: 700;
        margin-bottom: 6px;
    }
    .metric-card .value {
        font-size: 1.65rem;
        font-weight: 700;
        color: #0D2340;
        line-height: 1.1;
    }
    .metric-card .sub {
        font-size: 0.72rem;
        color: #94A3B8;
        margin-top: 4px;
    }
    .info-box {
        background: linear-gradient(135deg, #EFF6FF 0%, #E3F2FD 100%);
        border-left: 4px solid #1565C0;
        border-radius: 6px;
        padding: 12px 16px;
        font-size: 0.82rem;
        color: #1E3A5F;
        margin-bottom: 12px;
    }
    .section-divider {
        border: none;
        border-top: 1px solid rgba(13, 35, 64, 0.2);
        margin: 16px 0;
    }
    .footer-note {
        font-size: 0.70rem;
        color: #94A3B8;
        font-style: italic;
        text-align: center;
        margin-top: 28px;
        padding: 16px;
        border-top: 1px solid rgba(13, 35, 64, 0.1);
    }
    .stButton > button {
        background: linear-gradient(135deg, #1565C0 0%, #0D47A1 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 700;
        font-size: 0.85rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 12px rgba(21, 101, 192, 0.3);
        letter-spacing: 0.05em;
    }
    .stButton > button:hover {
        box-shadow: 0 6px 20px rgba(21, 101, 192, 0.5);
        transform: translateY(-2px);
    }
</style>
"""
