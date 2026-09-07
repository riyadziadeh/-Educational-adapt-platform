import os
import google.generativeai as genai

# 1. إعداد مفتاح الـ API (استبدل النص بمفتاحك أو استخدم متغيرات البيئة)
API_KEY = os.getenv("GEMINI_API_KEY", "ضع_مفتاحك_هنا")
genai.configure(api_key=API_KEY)

# استخدام أحدث نموذج مدعوم لتجنب أخطاء 404
MODEL_NAME = "gemini-2.5-flash"

# الخيارات والقوائم المطلوبة
EDUCATIONAL_SYSTEMS = [
    "وطني (National)", 
    "دولي (International)"
]

JORDAN_GOVERNORATES = [
    "العاصمة (عمان)",
    "إربد",
    "الزرقاء",
    "البلقاء",
    "المفرق",
    "الكرك",
    "مادبا",
    "جرش",
    "عجلون",
    "معان",
    "الطفيلة",
    "العقبة"
]

def adapt_worksheet(original_content, educational_system, governorate):
    """
    دالة لتكييف ورقة العمل مع دمج التفاصيل القديمة والجديدة:
    - محتوى ورقة العمل الأصلية
    - النظام التعليمي (وطني / دولي)
    - محافظة المدرسة في الأردن
    """
    # التحقق من إدخال المحتوى
    if not original_content or not original_content.strip():
        return "خطأ: يرجى إدخال محتوى ورقة العمل الأصلية أولاً."

    prompt = f"""
    أنت خبير تربوي ومصمم مناهج متخصص في المناهج الأردنية.
    قم بتكييف وتطوير ورقة العمل التالية لتتوافق بدقة مع المعطيات الآتية:
    - النظام التعليمي: {educational_system}
    - موقع المدرسة (المحافظة): {governorate} - الأردن
    
    محتوى ورقة العمل الأصلية:
    {original_content}
    
    يرجى إعادة صياغة ورقة العمل وتنظيمها بأسلوب تربوي احترافي ومناسب لطلبة الصف المستهدف.
    """

    try:
        # تهيئة واستدعاء النموذج المحدث
        model = genai.GenerativeModel(MODEL_NAME)
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"حدث خطأ أثناء الاتصال بالذكاء الاصطناعي: {str(e)}"

# --- مثال تجريبي للاستخدام الفعلي ---
if __name__ == "__main__":
    # محاكاة المدخلات (النص القديم + الخيارات الجديدة)
    user_original_text = """
    أجب عن الأسئلة الآتية:
    1. استخرج من النص: حرف جر، فعلاً ماضياً.
    2. صنف الكلمات حسب نوع اللام (شمسية أو قمرية): (السيارة - المعلم - الورد - الشباك).
    """
    
    chosen_system = EDUCATIONAL_SYSTEMS[0]      # نظام وطني
    chosen_governorate = JORDAN_GOVERNORATES[0]  # محافظة العاصمة (عمان)
    
    print(f"جاري معالجة ورقة العمل لنظام [{chosen_system}] في [{chosen_governorate}]...")
    
    # تنفيذ التكييف
    final_output = adapt_worksheet(user_original_text, chosen_system, chosen_governorate)
    
    print("\n--- ورقة العمل المطورة ---")
    print(final_output)
