import os
import io
import time
import streamlit as st
import streamlit.components.v1 as components
from google import genai
from google.genai import types
import pypdf
from PIL import Image

# محاولة استيراد مكتبات Word و PowerPoint بأمان تامة لضمان عدم انهيار السيرفر
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

# محاولة استيراد مكتبة FPDF لتوليد ملفات الـ PDF بأمان
try:
    from fpdf import FPDF
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# إعداد صفحة ستريمليت مع العنوان الرسمي الأنيق والأيقونة
st.set_page_config(
    page_title="نظام تكييف أوراق العمل بالذكاء الاصطناعي | Educational Worksheet Adaptation Platform", 
    page_icon="https://cdn.jsdelivr.net/gh/riyadziadeh/-Educational-adapt-platform@main/store_icon.png", 
    layout="centered"
)

# حقن الشعار الصحيح ضمن meta tags حتى تظهر الصورة الصحيحة عند مشاركة الرابط (واتساب/تيليجر