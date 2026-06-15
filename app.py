import streamlit as st
import os
from PIL import Image
import io
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# إعدادات الصفحة
st.set_page_config(page_title="أرشيف وصانع ألبومات الوورد", page_icon="📁", layout="wide")

st.title("📁 ألبوم صور المشاريع وصانع ملفات الوورد للطباعة")
st.markdown("---")

# المسارات الأساسية للتخزين
BASE_DIR = "photos"
BC_DIR = os.path.join(BASE_DIR, "Bons_de_Commande")
MARCHE_DIR = os.path.join(BASE_DIR, "Marches")

for folder in [BC_DIR, MARCHE_DIR]:
    os.makedirs(folder, exist_ok=True)

# القائمة الجانبية
st.sidebar.header("🔍 تصفح وأرشِف")
choice = st.sidebar.radio("اختر نوع المشروع:", ["سندات الطلب (Bons de Commande)", "الصفقات (Les Marchés)"])
target_dir = BC_DIR if "Bons de Commande" in choice else MARCHE_DIR

# قراءة المشاريع الحالية
projects = [d for d in os.listdir(target_dir) if os.path.isdir(os.path.join(target_dir, d))]

# إمكانية إضافة مشروع جديد ديريكت
new_project = st.sidebar.text_input("➕ إنشاء مشروع جديد (إسم أو رقم):")
if st.sidebar.button("إنشاء الدوسي"):
    if new_project:
        os.makedirs(os.path.join(target_dir, new_project), exist_ok=True)
        st.sidebar.success(f"✅ تم إنشاء دوسي: {new_project}")
        st.rerun()

# تصفح وعرض الألبوم
if projects:
    selected_project = st.sidebar.selectbox("اختر المشروع الحالي:", projects)
    project_path = os.path.join(target_dir, selected_project)
    
    st.header(f"🏗️ المشروع الحالي: {selected_project}")
    
    # رفع صور جديدة للدوسي
    uploaded_files = st.file_uploader("📸 إرفع الصور هنا لحفظها وتوليد ملف الوورد (يمكنك اختيار عدة صور معاً):", accept_multiple_files=True, type=["png", "jpg", "jpeg"])
    
    if uploaded_files:
        for uploaded_file in uploaded_files:
            save_path = os.path.join(project_path, uploaded_file.name)
            with open(save_path, "wb") as f:
                f.write(uploaded_file.get_buffer())
        st.success("✅ تم حفظ الصور في الدوسي بنجاح!")
    
    # قراءة الصور الموجودة في الدوسي لعرضها وصنع الوورد منها
    valid_extensions = (".png", ".jpg", ".jpeg", ".webp")
    images_in_folder = [img for img in os.listdir(project_path) if img.lower().endswith(valid_extensions)]
    
    if images_in_folder:
        st.write(f"📊 عدد الصور في هذا المشروع: {len(images_in_folder)}")
        
        # 📥 قسم توليد ملف الوورد للطباعة
        st.markdown("### 🖨️ تحميل الألبوم على شكل ملف Word جاهز")
        
        # خيارات التنسيق
        cols_per_page = st.radio("كيف بغيتي تستاف ديال التصاور فالصفحة د الوورد؟", ["صورة واحدة كبيرة في كل صفحة", "جوج صور متناسقين في كل صفحة"], index=1)
        
        if st.button("🪄 صاوب ليا ملف الوورد دابا"):
            with st.spinner("جاري إعداد ملف الوورد وترتيب الصور..."):
                doc = Document()
                
                # إعدادات الهوامش لتكبير مساحة العرض
                sections = doc.sections
                for section in sections:
                    section.top_margin = Inches(0.5)
                    section.bottom_margin = Inches(0.5)
                    section.left_margin = Inches(0.5)
                    section.right_margin = Inches(0.5)
                
                # عنوان الملف الرئيسي داخل الوورد
                title = doc.add_paragraph()
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = title.add_run(f"ألبوم صور مشروع: {selected_project}")
                run.font.name = 'Arial'
                run.font.size = Pt(22)
                run.bold = True
                
                doc.add_paragraph(f"نوع المشروع: {choice}").alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph("--------------------------------------------------").alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # إضافة الصور في الوورد
                if cols_per_page == "صورة واحدة كبيرة في كل صفحة":
                    for img_name in images_in_folder:
                        img_full_path = os.path.join(project_path, img_name)
                        
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run(f"📸 {img_name}\n").bold = True
                        
                        doc.add_picture(img_full_path, width=Inches(6.5)) # حجم كبير ملاءم للصفحة
                        doc.add_page_break() # صفحة جديدة للصورة الموالية
                else:
                    # صورتين في الصفحة
                    for idx, img_name in enumerate(images_in_folder):
                        img_full_path = os.path.join(project_path, img_name)
                        
                        p = doc.add_paragraph()
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p.add_run(f"📸 {img_name}").bold = True
                        
                        doc.add_picture(img_full_path, width=Inches(5.0)) # حجم متناسق باش يجيو جوج
                        
                        # يلا كملنا جوج تصاور نديرو صفحة جديدة، وإلا غير فراغ
                        if (idx + 1) % 2 == 0:
                            doc.add_page_break()
                        else:
                            doc.add_paragraph("\n")
                
                # حفظ الملف في الذاكرة باش يتشارجا ديريكت
                bio = io.BytesIO()
                doc.save(bio)
                bio.seek(0)
                
                st.download_button(
                    label="📥 اضغط هنا لتحميل ملف الوورد (.docx)",
                    data=bio,
                    file_name=f"Album_{selected_project.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.balloons()
        
        st.markdown("---")
        st.markdown("### 👁️ معاينة الصور على التطبيق:")
        # عرض شبكة الصور في التطبيق للعرض فقط
        cols = st.columns(3)
        for idx, img_name in enumerate(images_in_folder):
            img_path = os.path.join(project_path, img_name)
            with cols[idx % 3]:
                st.image(Image.open(img_path), caption=img_name, use_column_width=True)
                
    else:
        st.info("ℹ️ هاد الدوسي مازال ما فيه حتا شي تصويرة. إرفع أول تصويرة الفوق!")
else:
    st.info("ℹ️ ابدأ بإنشاء أول دوسي مشروع من القائمة الجانبية (مثلاً: BC_01_2026).")
