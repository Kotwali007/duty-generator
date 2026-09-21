import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import streamlit as st

st.set_page_config(page_title="थाना कोतवाली - ड्यूटी चार्ट जनरेटर", layout="wide")

# --- अद्यतन स्टाफ मास्टर डेटा (Master Staff List) ---
DEFAULT_STAFF = [
    {"name": "व0उ0नि0 श्री हरिओम सिंह", "phone": "8979768608", "gender": "M"},
    {"name": "प्र0नि0 श्री माघो सिंह विष्ट", "phone": "9454402942", "gender": "M"},
    {"name": "म0उ0नि0 श्रीमती कंचन लता", "phone": "6392118831", "gender": "F"},
    {"name": "उ0नि0 श्री जगमेन्दर सिंह", "phone": "9868944988", "gender": "M"},
    {"name": "उ0नि0 श्री संजय कुमार", "phone": "9997568632", "gender": "M"},
    {"name": "उ0नि0 श्री अवधेश कुमार", "phone": "7007573939", "gender": "M"},
    {"name": "उ0नि0 श्री आयुष आर्य", "phone": "8171984502", "gender": "M"},
    {"name": "उ0नि0 श्री रवेन्द्र सिंह", "phone": "8077662578", "gender": "M"},
    {"name": "उ0नि0 श्री आलोक कुमार", "phone": "9412826890", "gender": "M"},
    {"name": "उ0नि0 श्री पवन कुमार", "phone": "7457844908", "gender": "M"},
    {"name": "हे0का0 237 जसवीर", "phone": "8077629450", "gender": "M"},
    {"name": "का0 1723 रोमिल", "phone": "9808368421", "gender": "M"},
    {"name": "आ0चा0 हे0का0 योगेन्द्र सिंह", "phone": "", "gender": "M"},
    {"name": "का0 217 काली चरन", "phone": "9927158077", "gender": "M"},
    {"name": "हे0का0 332 गजेन्द्र सिंह", "phone": "7983656757", "gender": "M"},
    {"name": "का0 1767 सन्नी देव", "phone": "7017184899", "gender": "M"},
    {"name": "का0 1294 लवीन कुमार", "phone": "8218281737", "gender": "M"},
    {"name": "हे0का0 465 पवन कुमार", "phone": "8077855873", "gender": "M"},
    {"name": "का0 655 सुशील कुमार", "phone": "8057564076", "gender": "M"},
    {"name": "का0 1633 नीरज कुमार", "phone": "7668536100", "gender": "M"},
    {"name": "का0 858 ललित कुमार", "phone": "8384887402", "gender": "M"},
    {"name": "कम्प्यूटर आपरेटर शाहिद", "phone": "8171176064", "gender": "M"},
    {"name": "का0 1604 अंकित कुमार", "phone": "", "gender": "M"},
    {"name": "का0 302 रोहित कुमार", "phone": "7906423638", "gender": "M"},
    {"name": "म0का0 03 आंचल (CCTNS)", "phone": "7418024740", "gender": "F"},
    {"name": "का0 1989 आशीष चौहान", "phone": "8218308650", "gender": "M"},
    {"name": "का0 2067 प्रमोद कुमार", "phone": "9650734852", "gender": "M"},
    {"name": "हे0का0 254 धर्मेन्द्र सिंह", "phone": "8859682510", "gender": "M"},
    {"name": "का0 1207 चन्द्रपाल सिंह", "phone": "", "gender": "M"},
    {"name": "का0 2215 हरकेश कुमार", "phone": "9627847384", "gender": "M"},
    {"name": "का0 275 राकेश सिंह", "phone": "8630692893", "gender": "M"},
    {"name": "का0 1630 अविष्कार उजवल", "phone": "9528068044", "gender": "M"},
    {"name": "का0 2222 हिमांशु सागर", "phone": "8700586325", "gender": "M"},
    {"name": "का0 107 कुलदीप मावी", "phone": "9720085817", "gender": "M"},
    {"name": "हे0का0 665 वसीम", "phone": "9456656746", "gender": "M"},
    {"name": "का0 416 अंकुश पँवार", "phone": "8273736772", "gender": "M"},
    {"name": "का0 1314 जितेन्द्र कुमार", "phone": "8865930263", "gender": "M"},
    {"name": "का0 357 तेजपाल", "phone": "9992779903", "gender": "M"},
    {"name": "का0 1945 संजीत कुमार", "phone": "6398417500", "gender": "M"},
    {"name": "का0 1980 सुभम कालहर", "phone": "8191941817", "gender": "M"},
    {"name": "का0 1405 विनित कुमार", "phone": "7668416732", "gender": "M"},
    {"name": "का0 1972 नितीश कुमार", "phone": "9389609101", "gender": "M"},
    {"name": "का0 347 सुशील कुमार", "phone": "7505593425", "gender": "M"},
    {"name": "का0 1792 शाहिद", "phone": "7505083692", "gender": "M"},
    {"name": "का0 909 अजय कुमार", "phone": "9897589254", "gender": "M"},
    {"name": "हे0का0 359 दुर्वेश कुमार", "phone": "9759638980", "gender": "M"},
    {"name": "का0 1407 साजिद", "phone": "7060304602", "gender": "M"},
    {"name": "का0 89 अनित कुमार", "phone": "9759113594", "gender": "M"},
    {"name": "का0 523 ओम प्रकाश", "phone": "6398912086", "gender": "M"},
    {"name": "का0 34 पातीराम", "phone": "9412471320", "gender": "M"},
    {"name": "का0 809 दीपक यादव", "phone": "6397492328", "gender": "M"},
    {"name": "का0 213 विशाल त्रिवेदी", "phone": "9643979614", "gender": "M"},
    {"name": "हे0का0 490 राजीव यादव", "phone": "7668717368", "gender": "M"},
    {"name": "का0 1131 ओमप्रकाश सिंह", "phone": "7905752096", "gender": "M"},
    {"name": "का0 322 हिमांशु साहनी", "phone": "7248029286", "gender": "M"},
    {"name": "हे0का0 538 दीनदयाल", "phone": "062531610", "gender": "M"},
    {"name": "का0 1855 अंकुश कुमार", "phone": "8057563321", "gender": "M"},
    {"name": "का0 598 सुवेन्द्र कुमार", "phone": "9897452117", "gender": "M"},
    {"name": "हे0का0 243 अरूण कुमार", "phone": "8077908595", "gender": "M"},
    {"name": "का0 1272 नमित नागर", "phone": "7838973308", "gender": "M"},
    {"name": "का0 526 पुष्पेन्द्र कुमार", "phone": "9761630264", "gender": "M"},
    {"name": "का0 1894 विनीत कुमार", "phone": "9759507288", "gender": "M"},
    {"name": "का0 1368 अंकित कुमार", "phone": "9045450702", "gender": "M"},
    {"name": "का0 262 मौहम्मद रिहान", "phone": "8923437941", "gender": "M"},
    {"name": "रि0आ0 3344 लोकेन्द्र", "phone": "9368352026", "gender": "M"},
    {"name": "रि0आ0 2556 अमित कुमार", "phone": "9758503589", "gender": "M"},
    {"name": "रि0आ0 3019 अंकित कुमार", "phone": "9027914858", "gender": "M"},
    {"name": "रि0आ0 2523 अर्जुन राणा", "phone": "8126202274", "gender": "M"},
    {"name": "रि0आ0 3345 मोहित शर्मा", "phone": "7248643729", "gender": "M"},
    {"name": "रि0आ0 3340 मनीष", "phone": "6397606268", "gender": "M"},
    {"name": "रि0आ0 2785 खुशहाल सिंह", "phone": "8791565923", "gender": "M"},
    {"name": "रि0का0 2799 नितिन कुमार", "phone": "8057182201", "gender": "M"},
    {"name": "रि0आ0 2899 मनीष मावी", "phone": "9818376308", "gender": "M"},
    {"name": "रि0आ0 3154 अक्षय कुमार", "phone": "9105472609", "gender": "M"},
    {"name": "रि0आ0 3156 यश कुमार", "phone": "9520610285", "gender": "M"},
    {"name": "रि0आ0 3270 आदिल सैफी", "phone": "9897972232", "gender": "M"},
    {"name": "रि0आ0 2885 वरूण कुमार", "phone": "8449181881", "gender": "M"},
    {"name": "रि0आ0 3358 तुषार मावी", "phone": "7037931324", "gender": "M"},
    {"name": "रि0आ0 3308 निखिल कटारिया", "phone": "7452871017", "gender": "M"},
    {"name": "रि0आ0 2925 लविश", "phone": "8392838822", "gender": "M"},
    {"name": "रि0आ0 2932 नीशू पुनिया", "phone": "9368006066", "gender": "M"},
    {"name": "रि0आ0 2794 शुभम कुमार", "phone": "8171536355", "gender": "M"},
    {"name": "रि0आ0 2943 शिवा", "phone": "6397170451", "gender": "M"},
    {"name": "रि0आ0 3286 अजय कुमार", "phone": "7906376813", "gender": "M"},
    {"name": "रि0आ0 3162 कपिल कुमार", "phone": "8171791130", "gender": "M"},
    {"name": "रि0आ0 2715 नितिन कुमार", "phone": "9027015472", "gender": "M"},
    {"name": "रि0आ0 2699 आकाश कुमार चौधरी", "phone": "6396178361", "gender": "M"},
    {"name": "रि0आ0 1824 निशान्त मलिक", "phone": "9027832612", "gender": "M"},
    {"name": "रि0आ0 3031 नितिन कुमार", "phone": "9700000000", "gender": "M"},
    {"name": "रि0आ0 3277 आकाश मलिक", "phone": "7668589144", "gender": "M"},
    {"name": "रि0आ0 2662 रवि कुमार", "phone": "8273427380", "gender": "M"},
    {"name": "रि0आ0 2784 तेजप्रताप", "phone": "8448596773", "gender": "M"},
    {"name": "रि0आ0 3311 सागर कुमार", "phone": "8273749652", "gender": "M"},
    {"name": "रि0आ0 3259 मंयक कुमार", "phone": "6396272951", "gender": "M"},
    {"name": "रि0आ0 3218 अनिकेत गिरि", "phone": "8006832077", "gender": "M"},
    {"name": "रि0आ0 3353 रोहित पवार", "phone": "8979117348", "gender": "M"},
    {"name": "रि0आ0 2726 अमित आर्या", "phone": "9557109763", "gender": "M"},
    {"name": "रि0आ0 2907 गुड्डू", "phone": "9068718138", "gender": "M"},
    {"name": "रि0आ0 2512 शोभित स्वामी", "phone": "9897096479", "gender": "M"},
    {"name": "रि0आ0 407 फिरोज खान", "phone": "8175499660", "gender": "M"},
    {"name": "रि0आ0 2213 अरविन्द कुमार", "phone": "7889267459", "gender": "M"},
    {"name": "रि0आ0 2857 अमित कुमार", "phone": "9528877069", "gender": "M"},
    {"name": "रि0आ0 346 ललित पंवार", "phone": "7895854938", "gender": "M"},
    # --- महिला पुलिसकर्मी ---
    {"name": "म0का0 292 अनामिका", "phone": "8726601039", "gender": "F"},
    {"name": "म0का0 295 कंचन", "phone": "8859680883", "gender": "F"},
    {"name": "म0का0 2224 सुमन", "phone": "8006883723", "gender": "F"},
    {"name": "म0का0 1562 शिवानी", "phone": "6395213878", "gender": "F"},
    {"name": "म0का0 2206 आरती वर्मा", "phone": "8810943958", "gender": "F"},
    {"name": "म0का0 79 शीतू देवी", "phone": "9548035396", "gender": "F"},
    {"name": "म0का0 44 संजली साहू", "phone": "9956289489", "gender": "F"},
    {"name": "म0का0 1553 अंजली", "phone": "9027002064", "gender": "F"},
    {"name": "म0का0 146 अंशु चौधरी", "phone": "7906457344", "gender": "F"},
    {"name": "का0 403 अर्चना राजपूत", "phone": "6307856339", "gender": "F"},
    {"name": "म0का0 704 मनीषा", "phone": "7668872569", "gender": "F"},
    {"name": "रि0म0आ0 2387 पिंकी मावी", "phone": "9953452093", "gender": "F"},
    {"name": "रि0म0आ0 2338 प्रियंका यादव", "phone": "7819052785", "gender": "F"},
    {"name": "रि0म0आ0 2264 पूजापाल", "phone": "7505627244", "gender": "F"},
    {"name": "रि0म0आ0 2361 इन्दू पंवार", "phone": "7017842115", "gender": "F"},
    {"name": "रि0म0आ0 2375 वंशिका", "phone": "8630425483", "gender": "F"},
    {"name": "रि0म0आ0 2290 तनु पीएनओ", "phone": "7983753884", "gender": "F"},
    {"name": "रि0म0आ0 2249 तनु", "phone": "9548584448", "gender": "F"},
    {"name": "रि0म0आ0 2432 आरती", "phone": "8864839874", "gender": "F"},
    {"name": "रि0म0आ0 2379 अनीषा यादव", "phone": "8979860636", "gender": "F"},
    {"name": "रि0म0आ0 2440 शिवानी", "phone": "9084916609", "gender": "F"}
]

def get_phone(name):
    return next((s['phone'] for s in DEFAULT_STAFF if s['name'] == name), '')

def generate_kotwali_duty_excel(date_str, duty_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Duty Chart"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=14, bold=True)
    font_sec_header = Font(name="Calibri", size=11, bold=True)
    font_tbl_header = Font(name="Calibri", size=10, bold=True)
    font_data = Font(name="Calibri", size=10)

    fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_light_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin = Side(border_style="thin", color="000000")
    border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title Banner
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = f"ड्यूटी चार्ट दिनांक {date_str} थाना कोतवाली जनपद बदायूँ"
    title_cell.font = font_title
    title_cell.alignment = align_center
    title_cell.fill = fill_yellow

    for col in range(1, 8):
        ws.cell(row=1, column=col).border = border_thin

    # Table Headers
    ws.merge_cells("A2:A3")
    ws.merge_cells("B2:B3")
    ws.merge_cells("C2:C3")
    ws.merge_cells("D2:E2")
    ws.merge_cells("F2:G2")

    ws["A2"] = "क्र0सं0"
    ws["B2"] = "थाना"
    ws["C2"] = "ड्यूटी का प्रकार / स्थान"
    ws["D2"] = f"दिनांक {date_str} ड्यूटी दिन समय\n( 08.00 से 20.00 बजे तक )"
    ws["F2"] = f"दिनांक {date_str} रात्रि\n( 20.00 से सुबह 08.00 बजे तक )"

    ws["D3"] = "नाम अधि0/ कर्म0गण"
    ws["E3"] = "मोबाइल नं0"
    ws["F3"] = "नाम अधि0/ कर्म0"
    ws["G3"] = "मो0नं0"

    for r in [2, 3]:
        for c in range(1, 8):
            cell = ws.cell(row=r, column=c)
            cell.font = font_tbl_header
            cell.alignment = align_center
            cell.border = border_thin
            cell.fill = fill_light_gray

    def add_section_banner(row_idx, text):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
        cell = ws.cell(row=row_idx, column=1)
        cell.value = text
        cell.font = font_sec_header
        cell.alignment = align_center
        cell.fill = fill_light_gray
        for c in range(1, 8):
            ws.cell(row=row_idx, column=c).border = border_thin

    def write_block(start_row, sr_no, thana, location, day_staff, night_staff):
        num_rows = max(len(day_staff), len(night_staff), 1)
        end_row = start_row + num_rows - 1

        if num_rows > 1:
            ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
            ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
            ws.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)

        ws.cell(row=start_row, column=1, value=sr_no).alignment = align_center
        ws.cell(row=start_row, column=2, value=thana).alignment = align_center
        ws.cell(row=start_row, column=3, value=location).alignment = align_center

        for i in range(num_rows):
            r = start_row + i
            if i < len(day_staff):
                ws.cell(row=r, column=4, value=day_staff[i][0]).alignment = align_left
                ws.cell(row=r, column=5, value=str(day_staff[i][1])).alignment = align_center
            if i < len(night_staff):
                ws.cell(row=r, column=6, value=night_staff[i][0]).alignment = align_left
                ws.cell(row=r, column=7, value=str(night_staff[i][1])).alignment = align_center

            for c in range(1, 8):
                cell = ws.cell(row=r, column=c)
                cell.font = font_data
                cell.border = border_thin

        return end_row + 1

    cur_row = 4

    for block in duty_data:
        if block["type"] == "banner":
            add_section_banner(cur_row, block["title"])
            cur_row += 1
        elif block["type"] == "row":
            cur_row = write_block(
                cur_row,
                block["sr_no"],
                "कोतवाली",
                block["location"],
                block["day_staff"],
                block["night_staff"]
            )

    col_widths = {"A": 8, "B": 12, "C": 35, "D": 32, "E": 16, "F": 32, "G": 16}
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- UI Interface ---
st.title("🚓 थाना कोतवाली - ऑटोमैटिक व डायनेमिक ड्यूटी चार्ट जनरेटर")

st.sidebar.header("⚙️ सेटिंग्स व फ़िल्टर")

selected_date = st.sidebar.date_input("📅 ड्यूटी दिनांक चुनें")
date_str = selected_date.strftime("%d-%m-%Y")

st.sidebar.subheader("🏖️ अवकाश / गैरहाजिर कर्मचारी")
all_staff_names = [s["name"] for s in DEFAULT_STAFF]
women_staff_names = [s["name"] for s in DEFAULT_STAFF if s["gender"] == "F"]
men_staff_names = [s["name"] for s in DEFAULT_STAFF if s["gender"] == "M"]

absent_staff = st.sidebar.multiselect(
    "अवकाश/रवाना/गैरहाजिर कर्मचारी जिन्हें सूची से हटाना है:",
    options=all_staff_names,
    default=[]
)

def filter_staff(staff_tuples):
    return [s for s in staff_tuples if s[0] not in absent_staff]

st.subheader("📋 दैनिक ड्यूटी आवंटन")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1. अधिकारी व जी०डी०", 
    "2. बैंक व महिला हेल्प डेस्क", 
    "3. हमराह व पहरा", 
    "4. पिकेट व अन्य ड्यूटियाँ",
    "5. चीता मोबाइल (01-07)"
])

with tab1:
    st.markdown("### अधिकारी व जी०डी० ड्यूटी")
    day_officer = st.selectbox("दिवस/रात्रि अधिकारी", options=all_staff_names, index=0)
    gd_day = st.multiselect("जी०डी० कार्यलेख (दिवस)", options=all_staff_names, default=["रि0आ0 3270 आदिल सैफी"])
    gd_night = st.multiselect("जी०डी० कार्यलेख (रात्रि)", options=all_staff_names, default=["का0 1723 रोमिल"])

with tab2:
    st.markdown("### 🏦 बैंक ड्यूटी एवं महिला हेल्प डेस्क (केवल दिवस शिफ्ट)")
    col_b1, col_b2 = st.columns(2)
    with col_b1:
        sbi_jogipura = st.multiselect("SBI बैंक जोगीपुरा", options=women_staff_names, default=["म0का0 292 अनामिका"])
        sbi_ticketganj = st.multiselect("SBI बैंक टिकटगंज", options=women_staff_names, default=["म0का0 295 कंचन"])
        bob_jogipura = st.multiselect("बैंक ऑफ बड़ौदा जोगीपुरा", options=women_staff_names, default=["म0का0 2224 सुमन"])
    
    with col_b2:
        st.markdown("*महिला हेल्प डेस्क (Shift-wise)*")
        w_help_1 = st.selectbox("शिफ्ट 1 (08:00 से 14:00 - 1 महिला)", options=women_staff_names, index=3)
        w_help_2 = st.selectbox("शिफ्ट 2 (14:00 से 22:00 - 1 महिला)", options=women_staff_names, index=4)
        w_help_3 = st.selectbox("शिफ्ट 3 (22:00 से 04:00 - 1 महिला)", options=women_staff_names, index=5)

with tab3:
    st.markdown("### हमराह व पहरा ड्यूटी (पहरा: दिन में महिला, रात में पुरुष)")
    pehra_day = st.multiselect("पहरा ड्यूटी (दिवस - 2 महिलाएँ)", options=women_staff_names, default=["म0का0 292 अनामिका", "म0का0 295 कंचन"])
    pehra_night = st.multiselect("पहरा ड्यूटी (रात्रि - 2 पुरुष)", options=men_staff_names, default=["रि0आ0 407 फिरोज खान", "रि0आ0 2523 अर्जुन राणा"])

with tab4:
    st.markdown("### 🛡️ पिकेट एवं अन्य महत्वपूर्ण ड्यूटियाँ (दिन एवं रात्रि)")
    
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        khrati_day = st.multiselect("खराती चौक पिकेट (दिवस)", options=all_staff_names, default=[])
        khrati_night = st.multiselect("खराती चौक पिकेट (रात्रि)", options=all_staff_names, default=[])

        halwai_day = st.multiselect("हलवाई चौक पिकेट (दिवस)", options=all_staff_names, default=[])
        halwai_night = st.multiselect("हलवाई चौक पिकेट (रात्रि)", options=all_staff_names, default=[])

        chhe_sadka_day = st.multiselect("छह सड़का पिकेट (दिवस)", options=all_staff_names, default=[])
        chhe_sadka_night = st.multiselect("छह सड़का पिकेट (रात्रि)", options=all_staff_names, default=[])

        gopi_day = st.multiselect("गोपी चौक पिकेट (दिवस)", options=all_staff_names, default=[])
        gopi_night = st.multiselect("गोपी चौक पिकेट (रात्रि)", options=all_staff_names, default=[])

        ticketganj_piket_day = st.multiselect("टिकटगंज पिकेट (दिवस)", options=all_staff_names, default=[])
        ticketganj_piket_night = st.multiselect("टिकटगंज पिकेट (रात्रि)", options=all_staff_names, default=[])

        bob_piket_day = st.multiselect("बैंक ऑफ बड़ौदा पिकेट (दिवस)", options=all_staff_names, default=[])
        bob_piket_night = st.multiselect("बैंक ऑफ बड़ौदा पिकेट (रात्रि)", options=all_staff_names, default=[])

    with col_p2:
        bob_ticket_day = st.multiselect("बैंक ऑफ बड़ौदा टिकटगंज (दिवस)", options=all_staff_names, default=[])
        bob_ticket_night = st.multiselect("बैंक ऑफ बड़ौदा टिकटगंज (रात्रि)", options=all_staff_names, default=[])

        nehru_day = st.multiselect("नेहरू चौक पिकेट (दिवस)", options=all_staff_names, default=[])
        nehru_night = st.multiselect("नेहरू चौक पिकेट (रात्रि)", options=all_staff_names, default=[])

        malpul_day = st.multiselect("मालपुल तिराहा पिकेट (दिवस)", options=all_staff_names, default=[])
        malpul_night = st.multiselect("मालपुल तिराहा पिकेट (रात्रि)", options=all_staff_names, default=[])

        hdfc_day = st.multiselect("HDFC बैंक जोगीपुरा (दिवस)", options=all_staff_names, default=[])
        hdfc_night = st.multiselect("HDFC बैंक जोगीपुरा (रात्रि)", options=all_staff_names, default=[])

        st.markdown("---")
        st.markdown("*विशेष कार्यालय ड्यूटियाँ (केवल दिवस कालीन)*")
        cyber_desk = st.selectbox("साइबर हेल्प डेस्क", options=all_staff_names, index=all_staff_names.index("का0 302 रोहित कुमार") if "का0 302 रोहित कुमार" in all_staff_names else 0)
        igrs_portal = st.selectbox("IGRS पोर्टल", options=all_staff_names, index=all_staff_names.index("रि0का0 2932 नीशू पुनिया") if "रि0का0 2932 नीशू पुनिया" in all_staff_names else 0)
        
        court_pero_1 = st.selectbox("कोर्ट पेरोकार (1)", options=all_staff_names, index=0)
        court_pero_2 = st.selectbox("कोर्ट पेरोकार (2)", options=all_staff_names, index=1)
        dak_runner = st.selectbox("डाक रनर", options=all_staff_names, index=2)

with tab5:
    st.markdown("### 🐆 चीता मोबाइल ड्यूटी आवंटन (चीता 01 से 07)")
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        c1_day = st.multiselect("चीता - 01 चौकी लालपुल (दिवस)", options=all_staff_names, default=["उ0नि0 श्री जगमेन्दर सिंह", "का0 217 काली चरन", "का0 1767 सन्नी देव"])
        c1_night = st.multiselect("चीता - 01 चौकी लालपुल (रात्रि)", options=all_staff_names, default=["उ0नि0 श्री जगमेन्दर सिंह", "का0 275 राकेश सिंह", "का0 2222 हिमांशु सागर"])

        c2_day = st.multiselect("चीता - 02 चौकी मीराजी (दिवस)", options=all_staff_names, default=["उ0नि0 श्री संजय कुमार", "का0 1314 जितेन्द्र कुमार", "का0 107 कुलदीप मावी"])
        c2_night = st.multiselect("चीता - 02 चौकी मीराजी (रात्रि)", options=all_staff_names, default=["उ0नि0 श्री संजय कुमार", "का0 357 तेजपाल", "का0 416 अंकुश पँवार"])

        c3_day = st.multiselect("चीता - 03 चौकी सरकारीगंज (दिवस)", options=all_staff_names, default=["उ0नि0 श्री अवधेश कुमार", "का0 1980 सुभम कालहर", "रि0आ0 3277 आकाश मलिक"])
        c3_night = st.multiselect("चीता - 03 चौकी सरकारीगंज (रात्रि)", options=all_staff_names, default=["उ0नि0 श्री अवधेश कुमार", "का0 1972 नितीश कुमार", "रि0आ0 3019 अंकित कुमार"])

        c4_day = st.multiselect("चीता - 04 चौकी मालवीयगंज (दिवस)", options=all_staff_names, default=["उ0नि0 श्री आयुष आर्य", "का0 1792 शाहिद", "का0 347 सुशील कुमार"])
        c4_night = st.multiselect("चीता - 04 चौकी मालवीयगंज (रात्रि)", options=all_staff_names, default=["उ0नि0 श्री आयुष आर्य", "हे0का0 359 दुर्वेश कुमार", "का0 909 अजय कुमार"])

    with col_c2:
        c5_day = st.multiselect("चीता - 05 चौकी नई सराय (दिवस)", options=all_staff_names, default=["उ0नि0 श्री रवेन्द्र सिंह", "का0 34 पातीराम", "का0 89 अनित कुमार"])
        c5_night = st.multiselect("चीता - 05 चौकी नई सराय (रात्रि)", options=all_staff_names, default=["उ0नि0 श्री रवेन्द्र सिंह", "का0 213 विशाल त्रिवेदी", "का0 809 दीपक यादव"])

        c6_day = st.multiselect("चीता - 06 चौकी छ सड़का (दिवस)", options=all_staff_names, default=["उ0नि0 श्री आलोक कुमार", "का0 1131 ओमप्रकाश सिंह", "का0 262 मौहम्मद रिहान"])
        c6_night = st.multiselect("चीता - 06 चौकी छ सड़का (रात्रि)", options=all_staff_names, default=["उ0नि0 श्री आलोक कुमार", "हे0का0 538 दीनदयाल", "का0 322 हिमांशु साहनी"])

        c7_day = st.multiselect("चीता - 07 चौकी सोथा (दिवस)", options=all_staff_names, default=["उ0नि0 श्री पवन कुमार", "का0 1368 अंकित कुमार", "का0 1855 अंकुश कुमार"])
        c7_night = st.multiselect("चीता - 07 चौकी सोथा (रात्रि)", options=all_staff_names, default=["उ0नि0 श्री पवन कुमार", "का0 598 सुवेन्द्र कुमार", "का0 1272 नमित नागर"])

# --- स्वचालित रिजर्व स्टाफ गणना ---
assigned_set = set(absent_staff)
assigned_list = [
    day_officer,
    *gd_day, *gd_night,
    *sbi_jogipura, *sbi_ticketganj, *bob_jogipura,
    w_help_1, w_help_2, w_help_3,
    *pehra_day, *pehra_night,
    "प्र0नि0 श्री माघो सिंह विष्ट", "का0 1405 विनित कुमार", "आ0चा0 हे0का0 योगेन्द्र सिंह",
    "का0 1207 चन्द्रपाल सिंह", "का0 2215 हरकेश कुमार", "का0 1407 साजिद",
    "म0उ0नि0 श्रीमती कंचन लता", "हे0का0 243 अरूण कुमार", "का0 1630 अविष्कार उजवल", "का0 1894 विनीत कुमार",
    *khrati_day, *khrati_night, *halwai_day, *halwai_night,
    *chhe_sadka_day, *chhe_sadka_night, *gopi_day, *gopi_night,
    *ticketganj_piket_day, *ticketganj_piket_night, *bob_piket_day, *bob_piket_night,
    *bob_ticket_day, *bob_ticket_night, *nehru_day, *nehru_night,
    *malpul_day, *malpul_night, *hdfc_day, *hdfc_night,
    cyber_desk, igrs_portal, court_pero_1, court_pero_2, dak_runner,
    *c1_day, *c1_night, *c2_day, *c2_night, *c3_day, *c3_night,
    *c4_day, *c4_night, *c5_day, *c5_night, *c6_day, *c6_night,
    *c7_day, *c7_night
]

assigned_set.update(assigned_list)

reserve_staff_tuples = [
    (s["name"], s["phone"]) for s in DEFAULT_STAFF if s["name"] not in assigned_set
]

# ड्यूटी चार्ट डेटा स्ट्रक्चर
duty_structure = [
    {
        "type": "row", "sr_no": 1,
        "location": "रात्रिअधिकारी व दिवसअधिकारी",
        "day_staff": filter_staff([(day_officer, get_phone(day_officer))]),
        "night_staff": []
    },
    {
        "type": "row", "sr_no": 2,
        "location": "क्राइम/ सूचना / जी0डी0 कार्यलेख",
        "day_staff": filter_staff([(name, get_phone(name)) for name in gd_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in gd_night])
    },
    {"type": "banner", "title": "बैंक एवं महिला सुरक्षा ड्यूटी (केवल दिवस कालीन)"},
    {
        "type": "row", "sr_no": 1, "location": "SBI बैंक जोगीपुरा",
        "day_staff": filter_staff([(name, get_phone(name)) for name in sbi_jogipura]), "night_staff": []
    },
    {
        "type": "row", "sr_no": 2, "location": "SBI बैंक टिकटगंज",
        "day_staff": filter_staff([(name, get_phone(name)) for name in sbi_ticketganj]), "night_staff": []
    },
    {
        "type": "row", "sr_no": 3, "location": "बैंक ऑफ बड़ौदा जोगीपुरा",
        "day_staff": filter_staff([(name, get_phone(name)) for name in bob_jogipura]), "night_staff": []
    },
    {
        "type": "row", "sr_no": 4, "location": "महिला हेल्प डेस्क (08:00 से 14:00)",
        "day_staff": filter_staff([(w_help_1, get_phone(w_help_1))]), "night_staff": []
    },
    {
        "type": "row", "sr_no": 5, "location": "महिला हेल्प डेस्क (14:00 से 22:00)",
        "day_staff": filter_staff([(w_help_2, get_phone(w_help_2))]), "night_staff": []
    },
    {
        "type": "row", "sr_no": 6, "location": "महिला हेल्प डेस्क (22:00 से 04:00)",
        "day_staff": filter_staff([(w_help_3, get_phone(w_help_3))]), "night_staff": []
    },
    {"type": "banner", "title": "प्रभारी निरीक्षक हमराह/ द्वितीय मोबाइल हमराह ड्यूटी का विवरण"},
    {
        "type": "row", "sr_no": 1,
        "location": "प्र0नि0 महोदय हमराह प्रातः 7:00 बजे से",
        "day_staff": filter_staff([("प्र0नि0 श्री माघो सिंह विष्ट", get_phone("प्र0नि0 श्री माघो सिंह विष्ट")), ("का0 1405 विनित कुमार", get_phone("का0 1405 विनित कुमार")), ("आ0चा0 हे0का0 योगेन्द्र सिंह", get_phone("आ0चा0 हे0का0 योगेन्द्र सिंह"))]),
        "night_staff": filter_staff([("प्र0नि0 श्री माघो सिंह विष्ट", get_phone("प्र0नि0 श्री माघो सिंह विष्ट")), ("का0 1207 चन्द्रपाल सिंह", get_phone("का0 1207 चन्द्रपाल सिंह")), ("का0 2215 हरकेश कुमार", get_phone("का0 2215 हरकेश कुमार")), ("का0 1407 साजिद", get_phone("का0 1407 साजिद"))])
    },
    {
        "type": "row", "sr_no": 2,
        "location": "द्वितीय मोबाइल बैंक चैकिंग/मिशन शक्ति अभियान",
        "day_staff": filter_staff([("म0उ0नि0 श्रीमती कंचन लता", get_phone("म0उ0नि0 श्रीमती कंचन लता")), ("हे0का0 243 अरूण कुमार", get_phone("हे0का0 243 अरूण कुमार")), ("का0 1630 अविष्कार उजवल", get_phone("का0 1630 अविष्कार उजवल"))]),
        "night_staff": filter_staff([("का0 1723 रोमिल", get_phone("का0 1723 रोमिल")), ("का0 1894 विनीत कुमार", get_phone("का0 1894 विनीत कुमार"))])
    },
    {"type": "banner", "title": "पहरा ड्यूटी का विवरण"},
    {
        "type": "row", "sr_no": 1,
        "location": "पहरा ड्यूटी",
        "day_staff": filter_staff([(name, get_phone(name)) for name in pehra_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in pehra_night])
    },
    {"type": "banner", "title": "पिकेट एवं अन्य महत्वपूर्ण ड्यूटियाँ"},
    {
        "type": "row", "sr_no": 1, "location": "खराती चौक पिकेट",
        "day_staff": filter_staff([(name, get_phone(name)) for name in khrati_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in khrati_night])
    },
    {
        "type": "row", "sr_no": 2, "location": "हलवाई चौक पिकेट",
        "day_staff": filter_staff([(name, get_phone(name)) for name in halwai_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in halwai_night])
    },
    {
        "type": "row", "sr_no": 3, "location": "छह सड़का पिकेट",
        "day_staff": filter_staff([(name, get_phone(name)) for name in chhe_sadka_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in chhe_sadka_night])
    },
    {
        "type": "row", "sr_no": 4, "location": "गोपी चौक पिकेट",
        "day_staff": filter_staff([(name, get_phone(name)) for name in gopi_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in gopi_night])
    },
    {
        "type": "row", "sr_no": 5, "location": "टिकटगंज पिकेट",
        "day_staff": filter_staff([(name, get_phone(name)) for name in ticketganj_piket_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in ticketganj_piket_night])
    },
    {
        "type": "row", "sr_no": 6, "location": "बैंक ऑफ बड़ौदा पिकेट",
        "day_staff": filter_staff([(name, get_phone(name)) for name in bob_piket_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in bob_piket_night])
    },
    {
        "type": "row", "sr_no": 7, "location": "बैंक ऑफ बड़ौदा टिकटगंज",
        "day_staff": filter_staff([(name, get_phone(name)) for name in bob_ticket_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in bob_ticket_night])
    },
    {
        "type": "row", "sr_no": 8, "location": "नेहरू चौक पिकेट",
        "day_staff": filter_staff([(name, get_phone(name)) for name in nehru_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in nehru_night])
    },
    {
        "type": "row", "sr_no": 9, "location": "मालपुल तिराहा पिकेट",
        "day_staff": filter_staff([(name, get_phone(name)) for name in malpul_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in malpul_night])
    },
    {
        "type": "row", "sr_no": 10, "location": "HDFC बैंक जोगीपुरा",
        "day_staff": filter_staff([(name, get_phone(name)) for name in hdfc_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in hdfc_night])
    },
    {
        "type": "row", "sr_no": 11, "location": "साइबर हेल्प डेस्क",
        "day_staff": filter_staff([(cyber_desk, get_phone(cyber_desk))]),
        "night_staff": []
    },
    {
        "type": "row", "sr_no": 12, "location": "IGRS पोर्टल",
        "day_staff": filter_staff([(igrs_portal, get_phone(igrs_portal))]),
        "night_staff": []
    },
    {
        "type": "row", "sr_no": 13, "location": "कोर्ट पेरोकार",
        "day_staff": filter_staff([(court_pero_1, get_phone(court_pero_1)), (court_pero_2, get_phone(court_pero_2))]),
        "night_staff": []
    },
    {
        "type": "row", "sr_no": 14, "location": "डाक रनर",
        "day_staff": filter_staff([(dak_runner, get_phone(dak_runner))]),
        "night_staff": []
    },
    {"type": "banner", "title": "चीता मोबाइल ड्यूटी"},
    {
        "type": "row", "sr_no": 1, "location": "चीता - 01 चौकी लालपुल",
        "day_staff": filter_staff([(name, get_phone(name)) for name in c1_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in c1_night])
    },
    {
        "type": "row", "sr_no": 2, "location": "चीता - 02 चौकी मीराजी",
        "day_staff": filter_staff([(name, get_phone(name)) for name in c2_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in c2_night])
    },
    {
        "type": "row", "sr_no": 3, "location": "चीता - 03 चौकी सरकारीगंज",
        "day_staff": filter_staff([(name, get_phone(name)) for name in c3_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in c3_night])
    },
    {
        "type": "row", "sr_no": 4, "location": "चीता - 04 चौकी मालवीयगंज",
        "day_staff": filter_staff([(name, get_phone(name)) for name in c4_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in c4_night])
    },
    {
        "type": "row", "sr_no": 5, "location": "चीता - 05 चौकी नई सराय",
        "day_staff": filter_staff([(name, get_phone(name)) for name in c5_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in c5_night])
    },
    {
        "type": "row", "sr_no": 6, "location": "चीता - 06 चौकी छ सड़का",
        "day_staff": filter_staff([(name, get_phone(name)) for name in c6_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in c6_night])
    },
    {
        "type": "row", "sr_no": 7, "location": "चीता - 07 चौकी सोथा",
        "day_staff": filter_staff([(name, get_phone(name)) for name in c7_day]),
        "night_staff": filter_staff([(name, get_phone(name)) for name in c7_night])
    },
    {"type": "banner", "title": "रिजर्व थाना कोतवाली (आवश्यकता पड़ने पर ड्यूटी हेतु)"},
    {
        "type": "row", "sr_no": 1,
        "location": "रिजर्व स्टाफ (महिला व पुरुष)",
        "day_staff": filter_staff(reserve_staff_tuples),
        "night_staff": []
    }
]

st.divider()

if st.button("📊 आज की Excel फाइल जनरेट करें", type="primary"):
    excel_file = generate_kotwali_duty_excel(date_str, duty_structure)
    st.success(f"✅ दिनांक {date_str} का ड्यूटी चार्ट तैयार है!")
    st.download_button(
        label=f"📥 Kotwali_Duty_Chart_{date_str}.xlsx डाउनलोड करें",
        data=excel_file,
        file_name=f"Kotwali_Duty_Chart_{date_str}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )