import os
import google.generativeai as genai

# 1. إعداد مفتاح الـ API
API_KEY = os.getenv("GEMINI_API_KEY", "ضع_مفتاحك_هنا")
genai.configure(api_key=API_KEY)

# استخدام أحدث نموذج مدعوم لتجنب أخطاء 404
MODEL_NAME = "gemini-2.5-flash"

# 2. القوائم والتفاصيل الأساسية
GRADES = [
    "الصف الأول", "الصف الثاني", "الصف الثالث", "الصف الرابع", 
    "الصف الخامس", "الصف السادس", "الصف السابع", "الصف الثامن", "الصف التاسع"
]

EDUCATIONAL_SYSTEMS = [
    "وطني (National)", 
    "دولي (International)"
]

JORDAN_GOVERNORATES = [
    "العاصمة (عمان)", "إربد", "الزرقاء", "البلقاء", "المفرق", 
    "الكرك", "مادبا", "جرش", "عجلون", "معان", "الطفيلة", "العقبة"
]

# 3. التصنيفات الشاملة والدقيقة للحالات الخاصة والتربية الخاصة بناءً على طلبك
SPECIAL_CONDITIONS_CATEGORIES = {
    "1. الإعاقات الحسية والجسدية": [
        "الإعاقة البصرية (كف تام أو ضعف بصر شديد / بريل ومطبوعات كبيرة)",
        "الإعاقة السمعية (صمم تام أو ضعف سمعي بحاجة لمعينات/إشارة)",
        "الإعاقة الحركية أو الجسدية (شلل، ضمور عضلات، بتر أطراف، تشوهات)",
        "الإعاقة الحسية المزدوجة (الصم-المكفوفين)"
    ],
    "2. الاضطرابات النمائية وصعوبات التعلم": [
        "صعوبات التعلم المحددة (ديسليكسيا، عسر كتابة، صعوبة حساب)",
        "اضطراب طيف التوحد (ASD)",
        "اضطراب فرط الحركة ونقص الانتباه (ADHD)",
        "اضطرابات النطق واللغة والتواصل (تأتأة، عيوب نطق)"
    ],
    "3. الإعاقات الذهنية والسلوكية": [
        "الإعاقة الذهنية / العقلية (بسيطة، متوسطة، شديدة)",
        "الاضطرابات الانفعالية والسلوكية (قلق شديد، اكتئاب، مخاوف مدرسية)",
        "الإعاقات المتعددة (أكثر من إعاقة معاً)"
    ],
    "4. الإعاقات والحالات الصحية المزمنة": [
        "الأمراض المزمنة المحتاجة لمتابعة (سكري، ربو شديد، صرع، أمراض قلب)",
        "مرضى السرطان (برامج استكمال وعلاجات مستمرة)",
        "حالات الفشل الكلوي (غسيل دوري)"
    ],
    "5. فئة الموهبة والتفوق": [
        "الطلبة الموهوبون والمتفوقون (برامج إثراء معرفي وتسريع أكاديمي)"
    ]
}

def adapt_worksheet(original_content, grade, educational_system, governorate, condition_category, specific_condition):
    """
    دالة تكييف ورقة العمل الشاملة مع التصنيفات الدقيقة للحالات الخاصة.
    """
    if not original_content or not original_content.strip():
        return "خطأ: يرجى إدخال محتوى ورقة العمل الأصلية أولاً."

    prompt = f"""
    أنت خبير تربوي ومختص في مناهج التربية الخاصة والدمج. يرجى تكييف وتطوير ورقة العمل التالية بدقة عالية:
    - الصف الدراسي: {grade}
    - النظام التعليمي: {educational_system}
    - موقع المدرسة (المحافظة): {governorate} - الأردن
    - تصنيف الحالة: {condition_category} -> الحالة المحددة: {specific_condition}
    
    محتوى ورقة العمل الأصلية:
    {original_content}
    
    يرجى إعادة صياغة ورقة العمل وتنظيمها بطريقة تربوية احترافية تراعي الفروق الفردية والخصائص الطبية أو النمائية المذكورة بدقة تامة.
    """

    try:
        model = genai.GenerativeModel(MODEL_NAME)
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"حدث خطأ أثناء الاتصال بالذكاء الاصطناعي: {str(e)}"

# --- مثال تجريبي للاستخدام الفعلي ---
if __name__ == "__main__":
    selected_grade = GRADES[3]             # الصف الرابع
    selected_system = EDUCATIONAL_SYSTEMS[0]  # وطني
    selected_gov = JORDAN_GOVERNORATES[0]     # العاصمة
    
    # اختيار فئة وحالة من التصنيفات الجديدة
    category_key = "2. الاضطرابات النمائية وصعوبات التعلم"
    condition_val = SPECIAL_CONDITIONS_CATEGORIES[category_key][0] # صعوبات التعلم المحددة
    
    user_original_text = """
    أجب عن الأسئلة الآتية:
    1. استخرج من النص: حرف جر، وفعلاً ماضياً.
    2. صنف الكلمات حسب نوع اللام (شمسية أو قمرية): (السيارة - المعلم - الورد - الشباك).
    """
    
    print(f"جاري معالجة ورقة العمل لـ [{selected_grade}]، مع مراعاة حالة: [{condition_val}]...")
    
    final_output = adapt_worksheet(
        original_content=user_original_text,
        grade=selected_grade,
        educational_system=selected_system,
        governorate=selected_gov,
        condition_category=category_key,
        specific_condition=condition_val
    )
    
    print("\n--- ورقة العمل المطورة للتربية الخاصة ---")
    print(final_output)
