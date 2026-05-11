
import os
import uuid
import re
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from docx import Document
from docx.shared import Inches as DocxInches
from fpdf import FPDF
 
# ---------------- THEMES ---------------- #
THEMES = {
    "Emerald":       {"p": (33, 115, 70),  "bg": (220, 252, 231), "d": (20, 83, 45)},
    "Light Purple":  {"p": (109, 40, 217), "bg": (245, 243, 255), "d": (76, 29, 149)},
    "Dark Blue":     {"p": (30, 58, 138),  "bg": (219, 234, 254), "d": (23, 37, 84)},
    "Gradient Blue": {"p": (2, 132, 199),  "bg": (224, 242, 254), "d": (12, 74, 110)},
    "Midnight":      {"p": (15, 23, 42),   "bg": (226, 232, 240), "d": (2, 6, 23)},
    "Sunset":        {"p": (194, 65, 12),  "bg": (255, 247, 237), "d": (124, 45, 18)}
}
 
# ---------------- UTILS ---------------- #
def strip_emojis(text):
    return re.sub(r'[^\x00-\x7F]+', '', str(text)).strip()
 
def safe_filename(name):
    clean = re.sub(r'[\\/*?:"<>|]', "", str(name))
    return clean[:50] if clean else "data"
 
def load_file(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == '.csv': return {"Sheet1": pd.read_csv(path)}
    return pd.read_excel(path, sheet_name=None)
 
def clean_data(df):
    if df is None or df.empty: return pd.DataFrame()
    df = df.dropna(how='all').dropna(axis=1, how='all').reset_index(drop=True)
    
    # Simple header detection
    new_cols = []
    seen = set()
    for i, col in enumerate(df.columns):
        base = str(col).strip()
        if "Unnamed" in base or not base: base = f"Column_{i}"
        
        final = base
        c = 1
        while final in seen:
            final = f"{base}_{c}"
            c += 1
        seen.add(final)
        new_cols.append(final)
    df.columns = new_cols
    return df
 
def get_column_types(df):
    num, text = [], []
    for col in df.columns:
        # Exclude common non-value IDs from numeric analysis
        if any(k in str(col).lower() for k in ["date", "year", "month", "no", "id", "code", "sl"]):
            text.append(col)
            continue
            
        # Try converting to numeric if it's not already
        if not pd.api.types.is_numeric_dtype(df[col]):
            try:
                # Remove common currency symbols and commas
                clean_s = df[col].astype(str).str.replace(r'[$,₹,£,€,%, ]', '', regex=True)
                temp_num = pd.to_numeric(clean_s, errors='coerce')
                if temp_num.notna().mean() > 0.5:
                    df[col] = temp_num
            except: pass

        if pd.api.types.is_numeric_dtype(df[col]):
            num.append(col)
        else:
            text.append(col)
    return num, text
 
# ---------------- ANALYSIS ---------------- #
def generate_kpis(df, num_cols, txt_cols):
    kpis = {}
    if df.empty: return kpis
    
    # Most Repeated logic
    def _most_repeated(col):
        clean_s = df[col].astype(str).str.strip()
        clean_s = clean_s[~clean_s.str.lower().isin(["nan", "none", ""])]
        if clean_s.empty: return None, 0
        
        def get_pattern(x):
            s = re.sub(r'[^a-zA-Z0-9\s]', ' ', str(x))
            words = s.split()
            return " ".join(words[:3]).lower() if words else ""
            
        patterns = clean_s.apply(get_pattern)
        patterns = patterns[patterns != ""]
        if patterns.empty: return None, 0
        
        top_counts = patterns.value_counts()
        if top_counts.empty: return None, 0
        
        top = top_counts.index[0]
        cnt = int(top_counts.iloc[0])
        
        matches = clean_s[patterns == top].value_counts()
        if matches.empty: return None, 0
        
        orig = matches.index[0]
        return orig, cnt
 
    # Pick a good column for Most Repeated
    target = next((c for c in txt_cols if "name" in c.lower() or "client" in c.lower()), txt_cols[0] if txt_cols else None)
    if target:
        orig, count = _most_repeated(target)
        if orig: kpis[f"🔁 Most Repeated ({target})"] = {"Current": f"{orig} ({count} times)", "Trend": "Frequency"}
 
    # Calculate Grand Totals, Max, and Min for financial columns
    for col in num_cols:
        c_low = str(col).lower()
        if any(k in c_low for k in ["invoice", "value", "amount", "total", "price", "revenue", "cost", "amt"]):
            try:
                # Force numeric for sum/max/min to avoid mixed type errors
                v_numeric = pd.to_numeric(df[col], errors='coerce')
                v_total = v_numeric.sum()
                v_max = v_numeric.max()
                v_min = v_numeric.min()
                
                if pd.notna(v_total): kpis[f"💰 Total {col}"] = {"Current": f"{v_total:,.2f}", "Trend": f"Sum of {col}"}
                if pd.notna(v_max):   kpis[f"📈 Max: {col}"] = {"Current": f"{v_max:,.2f}", "Trend": f"Highest in {col}"}
                if pd.notna(v_min):   kpis[f"📉 Min: {col}"] = {"Current": f"{v_min:,.2f}", "Trend": f"Lowest in {col}"}
            except: pass

    for col in num_cols[:5]:
        if f"💰 Total {col}" not in kpis:
            try:
                v_max = df[col].max()
                v_avg = df[col].mean()
                if pd.notna(v_max): kpis[f"🔝 Max: {col}"] = {"Current": v_max, "Trend": f"Highest in {col}"}
                if pd.notna(v_avg): kpis[f"📊 Avg: {col}"] = {"Current": f"{v_avg:.2f}", "Trend": f"Mean of {col}"}
            except: pass
    return kpis

def get_overall_stats(df, num_cols):
    stats = {"Total Value": 0.0, "Average": 0.0, "Max Value": 0.0, "Min Value": 0.0, "Records": len(df)}
    if df.empty or not num_cols:
        return stats
    
    # Try to find a 'main' numeric column (Amount, Total, etc.)
    target = None
    for col in num_cols:
        c_low = str(col).lower()
        if any(k in c_low for k in ["total", "amount", "revenue", "price", "value", "amt"]):
            target = col
            break
    
    if not target:
        target = num_cols[0]
        
    try:
        v_numeric = pd.to_numeric(df[target], errors='coerce')
        stats["Total Value"] = float(v_numeric.sum())
        stats["Average"] = float(v_numeric.mean()) if not v_numeric.empty else 0.0
        stats["Max Value"] = float(v_numeric.max()) if not v_numeric.empty else 0.0
        stats["Min Value"] = float(v_numeric.min()) if not v_numeric.empty else 0.0
    except:
        pass
        
    return stats
 
def operational_summary(df, text_cols):
    if df.empty or len(text_cols) < 2: return None
    cat = text_cols[0]; act = text_cols[1]
    summary = df.groupby([cat, act]).size().reset_index(name='Frequency')
    return summary.sort_values(by='Frequency', ascending=False).head(30).to_dict('records')
 
def charts(df, num_cols, txt_cols, rid):
    path = f"static/{rid}"; os.makedirs(path, exist_ok=True)
    res = {}
    if df.empty: return res
    
    # 1. Bar Chart & Pie Chart (Category Distribution)
    if txt_cols:
        target_col = txt_cols[0]
        counts = df[target_col].value_counts().head(5)
        if not counts.empty:
            # Bar Chart
            plt.figure(figsize=(8, 5)); counts.plot(kind='bar', color='#217346')
            plt.title(f"Top 5: {target_col}"); plt.tight_layout()
            f_bar = os.path.join(path, "bar.png"); plt.savefig(f_bar); plt.close()
            res["Categorical Distribution (Bar)"] = {"path": f_bar, "conclusion": "Top categories identified."}
            
            # Pie Chart
            plt.figure(figsize=(8, 5)); counts.plot(kind='pie', autopct='%1.1f%%', colors=['#217346', '#4CAF50', '#8BC34A', '#C8E6C9', '#E8F5E9'])
            plt.title(f"Market Share: {target_col}"); plt.ylabel(""); plt.tight_layout()
            f_pie = os.path.join(path, "pie.png"); plt.savefig(f_pie); plt.close()
            res["Percentage Breakdown (Pie)"] = {"path": f_pie, "conclusion": "Proportional distribution visualized."}

    # 2. Line Chart (Numeric Trend)
    if num_cols:
        target_num = num_cols[0]
        # Use first 20 records for a clean trend view
        subset = df[target_num].head(20)
        if not subset.empty:
            plt.figure(figsize=(8, 5)); subset.plot(kind='line', marker='o', color='#217346')
            plt.title(f"Trend: {target_num}"); plt.tight_layout()
            f_line = os.path.join(path, "line.png"); plt.savefig(f_line); plt.close()
            res["Performance Trend (Line)"] = {"path": f_line, "conclusion": "Value fluctuations tracked."}
            
    return res
 
# ---------------- REPORTS (Simplified) ---------------- #
def create_ppt(kpis, charts_data, theme_name, client_data, top_table=None, insights=None, font_name="Arial"):
    prs = Presentation(); t = THEMES.get(theme_name, THEMES["Emerald"])
    PRI = RGBColor(*t["p"]); BG = RGBColor(*t["bg"]); WHITE = RGBColor(255,255,255)
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    box.fill.solid(); box.fill.fore_color.rgb = PRI
    
    if client_data and len(client_data) > 0:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        display_data = client_data[:12]
        headers = list(display_data[0].keys())
        rows, cols = len(display_data) + 1, len(headers)
        tbl = slide.shapes.add_table(rows, cols, Inches(0.2), Inches(1.0), Inches(9.6), Inches(5.5)).table
        
        # Set headers
        for i, h in enumerate(headers):
            tbl.cell(0, i).text = str(h)
            
        # Set rows
        for r_idx, row_data in enumerate(display_data):
            for c_idx, h in enumerate(headers):
                tbl.cell(r_idx + 1, c_idx).text = str(row_data.get(h, ""))
        
    if kpis:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        for i, (l, d) in enumerate(list(kpis.items())[:8]):
            x, y = Inches(0.3 + (i%2)*4.8), Inches(1 + (i//2)*1.5)
            b = slide.shapes.add_shape(1, x, y, Inches(4.5), Inches(1.2)); b.fill.solid(); b.fill.fore_color.rgb = BG
            tx_frame = slide.shapes.add_textbox(x, y, Inches(4.5), Inches(1.2)).text_frame
            p = tx_frame.paragraphs[0]
            p.text = f"{strip_emojis(l)}: {d['Current']}"
            p.font.size = Pt(14); p.font.name = font_name
 
    for k, d in charts_data.items():
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(d['path'], Inches(1), Inches(1), width=Inches(11))
        
    f = f"reports/report_{uuid.uuid4()}.pptx"; prs.save(f); return f
 
def create_docx(kpis, charts_data, theme_name, client_data, top_table=None, insights=None, font_name="Arial"):
    doc = Document(); doc.add_heading('Analysis Report', 0)
    for l, d in kpis.items(): doc.add_paragraph(f"{strip_emojis(l)}: {d['Current']}")
    for k, d in charts_data.items(): doc.add_heading(k, 1); doc.add_picture(d['path'], width=DocxInches(5))
    f = f"reports/report_{uuid.uuid4()}.docx"; doc.save(f); return f
 
def create_pdf(kpis, charts_data, theme_name, client_data, top_table=None, insights=None, font_name="Arial"):
    pdf = FPDF(); pdf.add_page(); pdf.set_font("Helvetica", 'B', 16)
    pdf.cell(190, 10, text="Analysis Report", align='C')
    pdf.ln(10)
    pdf.set_font("Helvetica", '', 10)
    for l, d in kpis.items(): pdf.multi_cell(190, 7, text=f"{strip_emojis(l)}: {d['Current']}", border=1)
    for k, d in charts_data.items(): pdf.add_page(); pdf.image(d['path'], x=10, w=180)
    f = f"reports/report_{uuid.uuid4()}.pdf"; pdf.output(f); return f
 
 # ---------------- DATA QUALITY ---------------- #
def data_quality_report(df):
    report = {}

    # Missing Values
    missing = df.isnull().sum().sum()

    # Duplicate Rows
    duplicates = df.duplicated().sum()

    # Empty Columns
    empty_cols = df.columns[df.isnull().all()].tolist()

    # Total Cells
    total_cells = df.shape[0] * df.shape[1]

    # Filled Cells
    filled_cells = total_cells - missing

    # Quality Score
    quality_score = (filled_cells / total_cells) * 100 if total_cells > 0 else 0

    report["Missing Values"] = int(missing)
    report["Duplicate Rows"] = int(duplicates)
    report["Empty Columns"] = len(empty_cols)
    report["Quality Score"] = f"{quality_score:.2f}/100"

    return report
# ---------------- ANOMALY DETECTION ---------------- #
def detect_anomalies(df, num_cols):
    anomalies = {}

    for col in num_cols:
        try:
            mean = df[col].mean()
            std = df[col].std()

            upper_limit = mean + (2 * std)
            lower_limit = mean - (2 * std)

            outliers = df[
                (df[col] > upper_limit) |
                (df[col] < lower_limit)
            ]

            if not outliers.empty:
                anomalies[col] = {
                    "count": len(outliers),
                    "max_value": outliers[col].max(),
                    "min_value": outliers[col].min()
                }

        except:
            pass

    return anomalies
# ---------------- TREND ANALYSIS ---------------- #
def trend_analysis(df, num_cols):
    trends = {}

    for col in num_cols:
        try:
            series = df[col].dropna()

            if len(series) < 2:
                continue

            first = series.iloc[0]
            last = series.iloc[-1]

            growth = ((last - first) / first) * 100 if first != 0 else 0

            if growth > 5:
                trend = "Increasing 📈"
            elif growth < -5:
                trend = "Decreasing 📉"
            else:
                trend = "Stable ➖"

            trends[col] = {
                "trend": trend,
                "growth_percent": f"{growth:.2f}%"
            }

        except:
            pass

    return trends

# ---------------- PROCESS ---------------- #
def process_file(path, theme, font=None, options=None):
    try:
        sheets = load_file(path); cleaned = {}
        for name, df in sheets.items():
            dfc = clean_data(df)
            if not dfc.empty: 
                dfc["Sheet Source"] = name
                cleaned[name] = dfc
        
        if not cleaned: return {"status": "error", "message": "Empty file."}
        
        order = {}
        if len(cleaned) > 1: order["Overall Summary"] = pd.concat(cleaned.values(), ignore_index=True, sort=False)
        order.update(cleaned)
        
        results = {}
        for name, df in order.items():
            num, txt = get_column_types(df)
            kpis = generate_kpis(df, num, txt)
            #data quality
            quality_report=data_quality_report(df)
            #anomaly detection
            anomalies=detect_anomalies(df,num)
            #trend analysis
            trends=trend_analysis(df,num)
            
            # Overall Stats for DB
            overall_stats = get_overall_stats(df, num)
            
            # Special KPIs for Overall Summary
            if name == "Overall Summary":
                kpis["📂 Total Sheets Merged"] = {"Current": len(cleaned), "Trend": "Source Composition"}
            
            ops = operational_summary(df, txt)
            ch = charts(df, num, txt, uuid.uuid4())
            
            # Mention sheet breakdown in insights
            sheet_info = []
            if name == "Overall Summary":
                for s_name, s_df in cleaned.items():
                    sheet_info.append(f"{s_name}: {len(s_df)} records")
            
            results[name] = {
                "kpis": kpis,
                "quality_report":quality_report,
                "anomalies":anomalies,
                "trends":trends,
                "ppt_path": create_ppt(kpis, ch, theme, ops, font_name=font),
                "docx_path": create_docx(kpis, ch, theme, ops, font_name=font),
                "pdf_path": create_pdf(kpis, ch, theme, ops, font_name=font),
                "overall_metrics": overall_stats,
                "numeric_insights": [f"{k}: {v['Current']}" for k,v in kpis.items()],
                "text_insights": sheet_info, 
                "charts": ch, 
                "full_data": df.head(500).to_dict('records')
            }
        return {"status": "success", "sheets": results}
    except Exception as e: return {"status": "error", "message": str(e)}