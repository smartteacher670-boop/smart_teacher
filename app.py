import streamlit as st
import base64
import os
import io
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# إعدادات الصفحة الأساسية
st.set_page_config(page_title="منصة الأستاذ الذكي - تضامن بيداغوجي واحترافي", page_icon="📚", layout="centered")

# --- دالة لتوليد ملف Word احترافي للأسبوع الأول أو الموسم الكامل ---
def generate_word_document(level_name, teacher_name, is_full_season=False):
    doc = Document()
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Traditional Arabic'
    font.size = Pt(14)
    font.color.rgb = RGBColor(0, 0, 0)
    
    p_header = doc.add_paragraph()
    p_header.paragraph_format.rtl = True
    p_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_h = p_header.add_run("الجمهورية الجزائرية الديمقراطية الشعبية\nوزارة التربية الوطنية\nمديرية التربية لولاية: ....................\nالمؤسسة الابتدائية: ....................\n")
    run_h.font.name = 'Traditional Arabic'
    run_h.font.size = Pt(12)
    
    title_text = f'كراس اليومية - الموسم الدراسي الكامل ({level_name})' if is_full_season else f'كراس اليومية - الأسبوع الأول (النسخة المجانية التجريبية) ({level_name})'
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.rtl = True
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(title_text)
    run_title.font.name = 'Traditional Arabic'
    run_title.font.size = Pt(18)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(46, 125, 50)
    
    p_info = doc.add_paragraph()
    p_info.paragraph_format.rtl = True
    p_info.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_info = p_info.add_run(f'الموسم الدراسي: 2026/2025  |  الأستاذ(ة): {teacher_name}  |  المستوى: {level_name}')
    run_info.font.name = 'Traditional Arabic'
    run_info.font.size = Pt(14)
    run_info.font.bold = True
    
    doc.add_paragraph('----------------------------------------------------------------------------------')
    
    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    headers = ['الملاحظات', 'النشاط والمحتوى البيداغوجي', 'الحصة / المادة', 'اليوم والتاريخ']
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        p = hdr_cells[i].paragraphs[0]
        p.paragraph_format.rtl = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.name = 'Traditional Arabic'
            run.font.bold = True
            run.font.size = Pt(14)

    rows_data = []
    
    if "الأولى" in level_name:
        rows_data = [
            ["الأحد (الدخول المدرسي)", "استقبال الأطفال، التعارف، تنظيم المقاعد، وتوجيهات أولية للدخول المدرسي.", "استقبال التلاميذ", "تنظيم الجو العام للقسم"],
            ["الأحد", "البسملة والتعوذ، قواعد الانضباط والنظافة داخل القسم المدرسي.", "تربية إسلامية / مدنية", "تحسيس التلاميذ"],
            ["الإثنين", "تواصل وتحية، التعبير الشفوي من خلال مشاهد دالة (صورة الدخول المدرسي).", "لغة عربية (شفوي)", "ملاحظة التعبير"],
            ["الإثنين", "التموضع في الفضاء والمقارنة (فوق، تحت، أمام، خلف، يمين، يسار).", "رياضيات", "تطبيق تدريبي"],
            ["الثلاثاء", "ألعاب تربوية تواصلية لكسر الجليد + تلوين وتخطيط أولي.", "تربية بدنية وفنية", "حركة ونشاط"],
            ["الأربعاء", "حفظ وتثبيت سورة الفاتحة (الاستماع والتلاوة النموذجية).", "تربية إسلامية", "متابعة الحفظ"],
            ["الخميس", "تقويم تشخيصي أولي للقدرات Hركية والبصرية والانتباه.", "تقويم تشخيصي", "تشخيص المكتسبات"]
        ]
        if is_full_season:
            rows_data.append(["الأسبوع 2 - الأحد", "تقديم الحروف الأساسية (أ، ب) + قراءة وتمارين مكتوبة.", "لغة عربية", "تثبيت الحروف"])
            rows_data.append(["الأسبوع 2 - الإثنين", "الأعداد من 1 إلى 5: التعرف، الكتابة والتمثيل.", "رياضيات", "حساب وعمليات"])
    else:
        rows_data = [
            ["الأحد (الدخول المدرسي)", "استقبال التلاميذ، توزيع التوقيت السنوي، وإجراء تقويم تشخيصي.", "استقبال وتنظيم", "ضبط النظام الداخلي"],
            ["الأحد", "فهم المنطوق (الأسبوع 1: العودة إلى المدرسة) + استخراج الأفكار الأساسية.", "لغة عربية", "إبداء الرأي"],
            ["الإثنين", "الأعداد إلى 999: التمثيل، المقارنة، والربط بين الكتابة الألفية والرمزية.", "رياضيات", "تمرينات فردية"]
        ]
        if is_full_season:
            rows_data.append(["الأسبوع 2 - الأحد", "الصيغ النحوية والتركيبية + التعبير الكتابي.", "لغة عربية", "قواعد لغوية"])

    for row_data in rows_data:
        row_cells = table.add_row().cells
        for i, text in enumerate(row_data):
            row_cells[i].text = text
            p = row_cells[i].paragraphs[0]
            p.paragraph_format.rtl = True
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for run in p.runs:
                run.font.name = 'Traditional Arabic'
                run.font.size = Pt(14)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def get_base64_of_bin_file(bin_file):
    if not os.path.exists(bin_file):
        return ""
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def set_png_as_page_bg(png_file):
    bin_str = get_base64_of_bin_file(png_file)
    page_bg_img = f'''
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Amiri:wght@400;700&family=Cairo:wght@400;600;700&display=swap');

    .stApp {{
        background-image: url("data:image/png;base64,{bin_str}");
        background-size: cover;
        background-attachment: fixed;
    }}
    
    .stApp, html, body, [class*="css"], label, input, select, button, p, span {{
        font-family: 'Cairo', 'Traditional Arabic', serif !important;
    }}
    
    label p, .stTextInput label, .stSelectbox label {{
        color: #111111 !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
    }}
    
    .greeting-text {{
        font-family: 'Amiri', 'Traditional Arabic', serif;
        font-size: 2.5rem;
        font-weight: 700;
        color: #FFFFFF;
        text-align: center;
        text-shadow: 2px 2px 8px rgba(0,0,0,0.9);
        padding: 20px 0;
        line-height: 1.4;
    }}
    
    div[data-testid="stForm"] {{
        background-color: #FFFFFF;
        border-radius: 16px;
        padding: 30px;
        box-shadow: 0 8px 25px rgba(0,0,0,0.3);
        border: 2px solid #2e7d32;
    }}
    
    div.stButton > button:first-child {{
        background-color: #2e7d32;
        color: white;
        font-size: 1.2rem;
        font-weight: bold;
        padding: 12px 24px;
        border-radius: 8px;
        border: none;
        width: 100%;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        transition: background-color 0.3s;
    }}
    div.stButton > button:first-child:hover {{
        background-color: #1b5e20;
    }}
    
    .emotional-hook {{
        background: linear-gradient(135deg, #fff8e1 0%, #ffe0b2 100%);
        border: 2px dashed #f57f17;
        border-radius: 14px;
        padding: 22px;
        text-align: right;
        margin-top: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        color: #4e342e;
    }}
    
    .pricing-card-pro {{
        background: #ffffff;
        border: 2px solid #2e7d32;
        border-radius: 14px;
        padding: 24px;
        text-align: center;
        margin-top: 20px;
        box-shadow: 0 6px 20px rgba(46,125,50,0.15);
    }}
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

background_image_path = 'background.jpg'
if os.path.exists(background_image_path):
    set_png_as_page_bg(background_image_path)

# --- لوحة تحكم المشرف لتفعيل الاشتراكات ---
with st.sidebar:
    st.markdown("### 🔐 لوحة تحكم المشرف")
    admin_password = st.text_input("كلمة مرور المشرف:", type="password")
    
    if admin_password == "admin123":
        st.success("مرحباً بك أستاذ (المشرف)!")
        
        if os.path.exists("leads_data.txt"):
            with open("leads_data.txt", "r", encoding="utf-8") as f:
                lines = f.readlines()
            
            if lines:
                st.info(f"إجمالي المسجلين: {len(lines)}")
                
                table_data = []
                for idx, line in enumerate(lines, 1):
                    parts = line.strip().split(" | ")
                    name_val = parts[0].replace("الاسم: ", "") if len(parts) > 0 else "---"
                    phone_val = parts[1].replace("الهاتف: ", "") if len(parts) > 1 else "---"
                    status_val = parts[4].replace("الحالة: ", "") if len(parts) > 4 else "مجاني"
                    
                    table_data.append({
                        "الرقم": idx,
                        "الاسم": name_val,
                        "الهاتف": phone_val,
                        "الحالة": status_val
                    })
                
                st.dataframe(table_data, use_container_width=True)
                
                st.write("---")
                st.markdown("#### تفعيل اشتراك مدفوع:")
                phone_to_activate = st.text_input("رقم هاتف الأستاذ:")
                if st.button("تفعيل الموسم الكامل ⭐"):
                    if phone_to_activate:
                        updated_lines = []
                        activated = False
                        for line in lines:
                            if phone_to_activate in line:
                                parts = line.strip().split(" | ")
                                parts[4] = "الحالة: مشترك مفعل"
                                updated_lines.append(" | ".join(parts) + "\n")
                                activated = True
                            else:
                                updated_lines.append(line)
                        
                        if activated:
                            with open("leads_data.txt", "w", encoding="utf-8") as f:
                                f.writelines(updated_lines)
                            st.success(f"تم تفعيل اشتراك الرقم {phone_to_activate} بنجاح!")
                        else:
                            st.error("رقم الهاتف غير موجود.")
            else:
                st.warning("لا توجد سجلات بعد.")
    elif admin_password != "":
        st.error("كلمة المرور غير صحيحة.")

# --- المحتوى الرئيسي ---
st.markdown('<p class="greeting-text">أستاذي، أستاذتي.. العمل لا يرهق بالمساعدة</p>', unsafe_allow_html=True)

st.markdown("""
<div style="background-color: #FFFFFF; padding: 22px; border-radius: 12px; color: #111; text-align: center; font-size: 1.15rem; box-shadow: 0 4px 15px rgba(0,0,0,0.2); border-right: 6px solid #2e7d32;">
<strong>عودة ميمونة أخي الأستاذ..</strong><br>
هديتي لكم هو <strong>الأسبوع الأول</strong> من كراس اليومية.<br>
⚠️ <em>تنبيه شفاف: باقي أسابيع الموسم الدراسي كاملاً متاحة حصرياً عبر الباقة المهنية المدفوعة.</em>
</div>
""", unsafe_allow_html=True)

st.divider()

if "submitted_successfully" not in st.session_state:
    st.session_state.submitted_successfully = False
    st.session_state.teacher_name = ""
    st.session_state.teacher_level = ""
    st.session_state.teacher_status = "مجاني"

tab1, tab2 = st.tabs(["📝 تسجيل جديد", "🔑 تسجيل الدخول"])

if not st.session_state.submitted_successfully:
    
    with tab1:
        st.write("")
        with st.form("teacher_form"):
            st.markdown("### <div style='color: #2e7d32;'>📩 سجل معلوماتك لاستلام هدية الأسبوع الأول:</div>", unsafe_allow_html=True)
            
            name_input = st.text_input("الاسم واللقب الكريم (إجباري):")
            phone_input = st.text_input("رقم الهاتف (مثال: 0612345678):")
            password_input = st.text_input("أنشئ كلمة سر شخصية خاصة بك:", type="password")
            level_input = st.selectbox(
                "اختر مستواك التدريسي لهذا الموسم:",
                ["السنة الأولى ابتدائي (1AP)", "السنة الثانية ابتدائي (2AP)", "السنة الثالثة ابتدائي (3AP)"]
            )
            
            st.write("")
            submitted = st.form_submit_button("إتمام التسجيل والوصول 🚀")
            
            if submitted:
                name_input = name_input.strip()
                phone_input = phone_input.strip()
                password_input = password_input.strip()
                
                if name_input == "" or phone_input == "" or password_input == "":
                    st.warning("⚠️ يرجى تعبئة جميع الحقول الإجبارية.")
                elif not re.match(r'^(05|06|07)[0-9]{8}$', phone_input):
                    st.error("⚠️ رقم الهاتف غير صحيح! يجب أن يتكون من 10 أرقام ويبدأ بـ 05، 06 أو 07.")
                else:
                    phone_exists = False
                    user_status_val = "مجاني"
                    if os.path.exists("leads_data.txt"):
                        with open("leads_data.txt", "r", encoding="utf-8") as f:
                            for line in f:
                                if phone_input in line:
                                    phone_exists = True
                                    if "مشترك مفعل" in line:
                                        user_status_val = "مشترك مفعل"
                                    break
                    
                    if phone_exists:
                        st.error("⚠️ رقم الهاتف مسجل مسبقاً! يرجى تسجيل الدخول من التبويب المجاور.")
                    else:
                        st.session_state.submitted_successfully = True
                        st.session_state.teacher_name = name_input
                        st.session_state.teacher_level = level_input
                        st.session_state.teacher_status = user_status_val
                        
                        try:
                            with open("leads_data.txt", "a", encoding="utf-8") as f:
                                f.write(f"الاسم: {name_input} | الهاتف: {phone_input} | المستوى: {level_input} | كلمة السر: {password_input} | الحالة: مجاني\n")
                        except Exception as e:
                            st.error(f"حدث خطأ أثناء الحفظ: {e}")
                        
                        st.rerun()

    with tab2:
        st.write("")
        with st.form("login_form"):
            st.markdown("### <div style='color: #1976d2;'>🔑 تسجيل الدخول للحساب:</div>", unsafe_allow_html=True)
            
            login_phone = st.text_input("أدخل رقم هاتفك المسجل:")
            login_pass = st.text_input("أدخل كلمة السر (اتركها فارغة لاسترجاعها):", type="password")
            
            st.write("")
            login_submitted = st.form_submit_button("دخول / استرجاع 🔍")
            
            if login_submitted:
                login_phone = login_phone.strip()
                login_pass = login_pass.strip()
                
                if login_phone == "":
                    st.warning("⚠️ يرجى إدخال رقم الهاتف.")
                else:
                    found_user = False
                    user_details = None
                    
                    if os.path.exists("leads_data.txt"):
                        with open("leads_data.txt", "r", encoding="utf-8") as f:
                            for line in f:
                                if login_phone in line:
                                    found_user = True
                                    user_details = line.strip()
                                    break
                    
                    if not found_user:
                        st.error("❌ رقم الهاتف غير مسجل. يرجى التسجيل أولاً!")
                    else:
                        parts = user_details.split(" | ")
                        saved_name = parts[0].replace("الاسم: ", "")
                        saved_level = parts[2].replace("المستوى: ", "")
                        saved_pass = parts[3].replace("كلمة السر: ", "") if len(parts) > 3 else ""
                        saved_status = parts[4].replace("الحالة: ", "") if len(parts) > 4 else "مجاني"
                        
                        if login_pass == "":
                            st.info(f"💡 أهلاً أستاذ **{saved_name}**! كلمة السر الخاصة بك هي: **{saved_pass}**")
                        elif login_pass == saved_pass:
                            st.session_state.submitted_successfully = True
                            st.session_state.teacher_name = saved_name
                            st.session_state.teacher_level = saved_level
                            st.session_state.teacher_status = saved_status
                            st.success(f"مرحباً بك مجدداً يا أستاذ {saved_name}!")
                            st.rerun()
                        else:
                            st.error("❌ كلمة السر غير صحيحة! اتركها فارغة لاسترجاعها.")

if st.session_state.submitted_successfully:
    selected_level = st.session_state.teacher_level
    teacher_name = st.session_state.teacher_name
    is_active_subscriber = (st.session_state.get("teacher_status", "مجاني") == "مشترك مفعل")
    
    st.markdown(f"""
    <div style="background: #ffffff; border: 2px solid #2e7d32; border-radius: 14px; padding: 22px; text-align: center; margin-top: 15px;">
        <h2 style="color: #2e7d32; margin-bottom: 8px;">👋 أهلاً بك يا أستاذ {teacher_name}</h2>
        <p style="font-size: 1.1rem; color: #333;">
            المستوى المسجل: <strong>{selected_level}</strong> | نوع العضوية: <span style="color: {'#2e7d32' if is_active_subscriber else '#d32f2f'}; font-weight: bold;">{'مشترك في الموسم الكامل ⭐' : 'عضو مجاني (الأسبوع الأول فقط 🎁)'}</span>
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.write("")
    
    # توليد وتحميل الملف المناسب لحالة الأستاذ
    word_file = generate_word_document(selected_level, teacher_name, is_full_season=is_active_subscriber)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        btn_label = f"📥 تحميل كراس اليومية (الموسم الدراسي الكامل) - {selected_level}" if is_active_subscriber else f"📥 تحميل هدية الأسبوع الأول (مجاناً) - {selected_level}"
        st.download_button(
            label=btn_label,
            data=word_file,
            file_name=f"كراس_اليومية_{selected_level.split()[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    
    if not is_active_subscriber:
        st.markdown("""
        <div class="emotional-hook">
            <h3 style="color: #d84315; margin-bottom: 12px; font-weight: 700;">☕ أستاذي الفاضل.. هل ستقضي كل عطلة نهاية أسبوع في إرهاق الكتابة والتحضير اليدوي؟</h3>
            <p style="font-size: 1.1rem; line-height: 1.8; color: #3e2723;">
                نعلم جيداً كم تستهلك تدوينات كراس اليومية من وقتك الثمين على حساب راحتك، عائلتك، وصحتك النفسية والجسدية. 
                أسبوع واحد مجاني هو مجرد <strong>نافذة صغيرة</strong> لترى بنفسك جودة التنظيم ودقة التخطيط... لكن ماذا عن الأسابيع الـ 34 القادمة؟ 
                هل تستحق راحة بالك واستقرارك الذهني طوال الموسم الدراسي مبلغاً رمزياً لا يوازي سعر كوب قهوة؟
            </p>
        </div>
        
        <div class="pricing-card-pro">
            <h3 style="color: #2e7d32; margin-bottom: 12px;">⭐ حرّر وقتك واضمن راحة بالاك طوال السنة</h3>
            <p style="font-size: 1.05rem; color: #444; line-height: 1.6; margin-bottom: 15px;">
                احصل على كراس اليومية مكتملاً، منسقاً، وجاهزاً لكل أسابيع الموسم الدراسي دون عناء التفكير أو البحث الأسبوعي المتكرر.
            </p>
            <div style="background: #e8f5e9; padding: 14px; border-radius: 8px; border: 1px solid #c8e6c9; margin-bottom: 15px;">
                <strong style="color: #1b5e20; font-size: 1.1rem;">💳 استثمار بسيط لراحة ممتدة:</strong><br>
                <span style="color: #333; font-size: 1.05rem;">فعّل حسابك الآن عبر (BaridiMob / CCP) وتواصل مع المشرف لتحميل الموسم كاملاً بضغطة زر واحدة.</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    st.write("")
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        if st.button("تسجيل الخروج 🚪"):
            st.session_state.submitted_successfully = False
            st.rerun()

st.divider()
st.caption("💡 مبادرة تضامنية بيداغوجية بين الأساتذة | منصة الأستاذ الذكي © 2026")
