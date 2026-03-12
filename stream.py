# ==========================================================
# RegulaAI – AI Powered Regulatory Compliance System
# ==========================================================

import streamlit as st
import PyPDF2
import os
import io
import re
import smtplib
import matplotlib.pyplot as plt
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

from groq import Groq
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch
from dotenv import load_dotenv

# Try to import docx, install if not available
try:
    import docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# Load environment
load_dotenv()

# Get API keys
GROQ_KEY = os.getenv("GROQ_API_KEY")
EMAIL_SENDER = os.getenv("SENDER_EMAIL")
EMAIL_PASSWORD = os.getenv("SENDER_PASSWORD")

# Initialize Groq client
if GROQ_KEY:
    try:
        client = Groq(api_key=GROQ_KEY)
    except:
        import httpx
        client = Groq(api_key=GROQ_KEY, http_client=httpx.Client())
else:
    client = None

# ==========================================================
# CLAUSE DEFINITIONS
# ==========================================================

CLAUSE_PATTERNS = {
    "Confidentiality": {
        "patterns": [
            r'confidential(?:ity)?\s+(?:agreement|clause|obligation)',
            r'non-?disclosure',
            r'proprietary\s+information',
            r'secret',
            r'private\s+information'
        ],
        "risk": "medium",
        "description": "Controls how confidential information is handled"
    },
    "Termination": {
        "patterns": [
            r'terminat(?:ion|e|es)\s+(?:clause|condition|right)',
            r'end\s+of\s+(?:agreement|contract)',
            r'cancel(?:lation)?',
            r'breach\s+of\s+contract'
        ],
        "risk": "high",
        "description": "Defines conditions under which contract can be ended"
    },
    "Payment Terms": {
        "patterns": [
            r'payment\s+(?:terms|condition|schedule)',
            r'fee\s+(?:structure|schedule)',
            r'invoice',
            r'compensation',
            r'price\s+and\s+payment'
        ],
        "risk": "high",
        "description": "Specifies payment amounts, timing, and methods"
    },
    "Liability": {
        "patterns": [
            r'liability\s+(?:clause|limitation|limit)',
            r'limit\s+of\s+liability',
            r'damages',
            r'indemnif',
            r'liquidated\s+damages'
        ],
        "risk": "high",
        "description": "Limits or defines financial responsibilities"
    },
    "Intellectual Property": {
        "patterns": [
            r'intellectual\s+property',
            r'i\.?p\s+rights',
            r'patent\s+and\s+copyright',
            r'trade\s+secret',
            r'ownership\s+of\s+(?:rights|property)'
        ],
        "risk": "medium",
        "description": "Defines ownership of creative works and inventions"
    },
    "Dispute Resolution": {
        "patterns": [
            r'dispute\s+resolution',
            r'arbitration',
            r'governing\s+law',
            r'jurisdiction',
            r'legal\s+proceedings'
        ],
        "risk": "medium",
        "description": "Defines how conflicts will be resolved"
    },
    "Force Majeure": {
        "patterns": [
            r'force\s+majeure',
            r'act\s+of\s+god',
            r'unforeseeable\s+circumstances',
            r'natural\s+disaster'
        ],
        "risk": "low",
        "description": "Excuses performance due to extraordinary events"
    },
    "Warranty": {
        "patterns": [
            r'warrant(?:y|ies)?\s+(?:clause|statement)',
            r'guarantee',
            r'representation',
            r'declaration\s+of'
        ],
        "risk": "medium",
        "description": "Promises about quality or performance"
    },
    "Assignment": {
        "patterns": [
            r'assign(?:ment)?\s+(?:clause|right|restriction)',
            r'transfer\s+of\s+rights',
            r'successor\s+and\s+assigns'
        ],
        "risk": "low",
        "description": "Controls transfer of contractual rights"
    },
    "Indemnification": {
        "patterns": [
            r'indemnif(?:y|ication)\s+(?:clause|obligation)',
            r'hold\s+harmless',
            r'defense\s+and\s+indemnification'
        ],
        "risk": "high",
        "description": "Requires one party to cover losses of another"
    },
    "Notices": {
        "patterns": [
            r'notice\s+(?:clause|requirement|provision)',
            r'communication\s+(?:clause|method)',
            r'written\s+notice'
        ],
        "risk": "low",
        "description": "Defines how formal communications are sent"
    },
    "Entire Agreement": {
        "patterns": [
            r'entire\s+agreement\s+clause',
            r'whole\s+agreement',
            r'merger\s+clause',
            r'supersedes\s+all'
        ],
        "risk": "low",
        "description": "States contract contains all terms"
    }
}

# Risk level colors and icons
RISK_CONFIG = {
    "high": {"color": "🔴", "level": "High Risk", "score": 75},
    "medium": {"color": "🟠", "level": "Medium Risk", "score": 50},
    "low": {"color": "🟢", "level": "Low Risk", "score": 25},
    "missing": {"color": "⚪", "level": "Missing Clause", "score": 0}
}

# ==========================================================
# SESSION STATE INITIALIZATION
# ==========================================================

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "contract_text" not in st.session_state:
    st.session_state.contract_text = ""

if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

if "improved_contract" not in st.session_state:
    st.session_state.improved_contract = ""

if "analysis_results" not in st.session_state:
    st.session_state.analysis_results = None

# ==========================================================
# LOGIN/LOGOUT FUNCTIONS
# ==========================================================

def login_page():
    # Custom CSS for login with Google Fonts
    st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <style>
    .login-container {
        max-width: 400px;
        margin: 50px auto;
        padding: 30px;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 20px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.3);
    }
    .login-title {
        /* Using Poppins font - imported via Google Fonts */
        font-family: 'Poppins', sans-serif;
        text-align: center;
        font-size: 42px;
        font-weight: 800;
        margin-bottom: 8px;
        /* Gradient: Purple → Blue → Pink */
        background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        /* Glow effect */
        text-shadow: 0px 4px 20px rgba(102, 126, 234, 0.5);
        letter-spacing: 2px;
    }
    
    .login-subtitle {
        color: #888;
        text-align: center;
        font-size: 14px;
        margin-bottom: 30px;
        font-family: 'Poppins', sans-serif;
        font-weight: 400;
        letter-spacing: 0.5px;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Center the login form
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        # Modern Brand Title with Icon, Gradient, and Glow
        st.markdown("""
        <div style="text-align: center; margin-bottom: 10px;">
            <!-- Legal/AI themed icon -->
            <span style="font-size: 36px; vertical-align: middle;">⚖️</span>
        </div>
        <!-- Modern Gradient Title with Glow Effect -->
        <h1 style="
            text-align: center; 
            font-family: 'Poppins', sans-serif;
            font-size: 48px; 
            font-weight: 800; 
            margin: 10px 0;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 50%, #f093fb 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
            text-shadow: 0px 4px 30px rgba(102, 126, 234, 0.6);
            letter-spacing: 3px;
        ">RegulaAI</h1>
        <!-- Clean subtitle with light gray -->
        <p style="
            text-align: center; 
            color: #888888; 
            font-family: 'Poppins', sans-serif;
            font-size: 14px; 
            font-weight: 400;
            margin-top: 5px;
            margin-bottom: 25px;
            letter-spacing: 1px;
        ">AI-Powered Regulatory Compliance System</p>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        with st.form("login_form"):
            email = st.text_input("📧 Email", placeholder="Enter your email address")
            password = st.text_input("🔑 Password", type="password", placeholder="Enter password (min 6 characters)")
            
            st.markdown("---")
            
            submit = st.form_submit_button("🚀 Login", use_container_width=True)
            
            if submit:
                if "@" in email and len(password) >= 6:
                    st.session_state.logged_in = True
                    st.session_state.user_email = email
                    st.success("✅ Login successful!")
                    st.rerun()
                else:
                    st.error("❌ Invalid email or password (min 6 characters)")
        
        st.markdown("---")
        st.caption("🔒 Secure Login | RegulaAI v1.0")
        
        # Social Login Divider
        st.markdown("""
        <div style="text-align: center; margin: 20px 0; position: relative;">
            <span style="background: white; padding: 0 15px; color: #888; font-size: 13px; position: relative;">or continue with</span>
        </div>
        """, unsafe_allow_html=True)
        
        # Social Login Buttons (Google, Facebook, Apple)
        col_google, col_fb, col_apple = st.columns([1, 1, 1])
        
        with col_google:
            st.markdown("""
            <button style="
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 8px;
                background: white;
                border: 2px solid #e0e0e0;
                border-radius: 10px;
                padding: 12px;
                font-size: 14px;
                font-weight: 500;
                color: #333;
                cursor: pointer;
                transition: all 0.3s ease;
                width: 100%;
            " disabled>
                <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
                    <path d="M17.5 9.25c0-.81-.07-1.59-.19-2.33H9v4.42h4.72c-.2 1.08-.82 2-1.73 2.61v2.18h2.81c1.63-1.5 2.81-3.71 2.81-6.38z" fill="#4285F4"/>
                    <path d="M9 16.5c2.24 0 4.11-.74 5.47-2.01l-2.81-2.18c-.74.5-1.69.79-2.66.79-2.05 0-3.79-1.39-4.41-3.26H.91v2.25C2.45 14.92 5.41 16.5 9 16.5z" fill="#34A853"/>
                    <path d="M4.59 10.67c-.22-.66-.35-1.36-.35-2.17s.13-1.51.35-2.17V5.93H.91C.32 7.08 0 8.5 0 10s.32 2.92.91 4.07l2.68-2.4z" fill="#FBBC05"/>
                    <path d="M9 3.53c1.17 0 2.22.4 3.04 1.2l2.73-2.73C13.11.89 11.24 0 9 0 5.41 0 2.45 1.58.91 3.93l2.68 2.4C5.21 4.92 6.95 3.53 9 3.53z" fill="#EA4335"/>
                </svg>
                Google
            </button>
            """, unsafe_allow_html=True)
        
        with col_fb:
            st.markdown("""
            <button style="
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 8px;
                background: white;
                border: 2px solid #e0e0e0;
                border-radius: 10px;
                padding: 12px;
                font-size: 14px;
                font-weight: 500;
                color: #333;
                cursor: pointer;
                transition: all 0.3s ease;
                width: 100%;
            " disabled>
                <svg width="18" height="18" viewBox="0 0 18 18" fill="#1877F2">
                    <path d="M18 9.17c0-.82-.07-1.64-.19-2.43h-8.52v4.63h4.87c-.21 1.13-.85 2.09-1.82 2.73v2.27h2.95c1.72-1.58 2.71-3.91 2.71-6.2z"/>
                    <path d="M9 18c2.43 0 4.47-.8 5.96-2.18l-2.95-2.27c-.8.54-1.83.86-3.01.86-2.32 0-4.28-1.57-4.98-3.68H.92v2.35C2.36 15.67 5.4 18 9 18z"/>
                    <path d="M4.02 10.73c-.24-.71-.38-1.46-.38-2.23s.14-1.52.38-2.23V4.92H.92C.25 6.23 0 7.72 0 9.5s.25 3.27.92 4.58l2.1-2.35z"/>
                    <path d="M9 3.58c1.34 0 2.54.46 3.49 1.37l2.7-2.7C13.46.88 11.39 0 9 0 5.4 0 2.36 2.33.92 4.92l2.1 2.1c.7-2.11 2.66-3.44 4.98-3.44z"/>
                </svg>
                Facebook
            </button>
            """, unsafe_allow_html=True)
        
        with col_apple:
            st.markdown("""
            <button style="
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 8px;
                background: white;
                border: 2px solid #e0e0e0;
                border-radius: 10px;
                padding: 12px;
                font-size: 14px;
                font-weight: 500;
                color: #333;
                cursor: pointer;
                transition: all 0.3s ease;
                width: 100%;
            " disabled>
                <svg width="18" height="18" viewBox="0 0 18 18" fill="#000000">
                    <path d="M13.67 9.52c0-1.12.32-2.19.87-3.07.52-.84 1.23-1.56 2.08-2.11.57-.37 1.21-.66 1.89-.85.68-.19 1.39-.29 2.12-.29.18 0 .36.01.54.02-.17-.56-.26-1.15-.26-1.76 0-2.91 1.17-5.46 3.08-6.83A8.44 8.44 0 0018 1.06 8.558 8.558 0 0013.67.23C10.4.23 7.57 1.67 5.55 4.11 3.53 6.55 2.32 9.77 2.32 13.23c0 .68.06 1.35.17 2a8.196 8.196 0 003.42-.76c-.93-.47-1.7-1.13-2.27-1.93-.57-.8-.87-1.75-.87-2.77 0-.47.09-.93.25-1.36.16-.44.4-.85.71-1.22.31-.37.69-.69 1.12-.95.43-.26.91-.45 1.41-.57.5-.12 1.02-.17 1.54-.15.52.02 1.03.1 1.52.25.49.14.95.35 1.37.62.42.27.79.6 1.1.98.31.38.56.81.75 1.27.19.46.3.95.32 1.45v.28a4.28 4.28 0 01-2.24.74 4.42 4.42 0 01-1.58-.3 4.19 4.19 0 01-1.24-.8 4.38 4.38 0 01-.84-1.13 4.28 4.28 0 01-.32-1.4c0-.58.09-1.15.27-1.69.18-.54.44-1.04.77-1.48.33-.44.73-.81 1.19-1.1a4.5 4.5 0 011.52-.7 4.55 4.55 0 011.65-.24c.43.03.85.1 1.26.21z"/>
                </svg>
                Apple
            </button>
            """, unsafe_allow_html=True)
        
        # Footer: Don't have an account?
        st.markdown("""
        <div style="text-align: center; margin-top: 25px; color: #666; font-size: 14px;">
            Don't have an account? <a href="#" style="color: #667eea; font-weight: 600; text-decoration: none;">Create Account</a>
        </div>
        """, unsafe_allow_html=True)

def logout():
    st.session_state.logged_in = False
    st.session_state.contract_text = ""
    st.session_state.chat_history = []
    st.session_state.improved_contract = ""
    st.session_state.analysis_results = None
    st.rerun()

# ==========================================================
# FILE EXTRACTION FUNCTIONS
# ==========================================================

def extract_text_from_pdf(file):
    try:
        pdf_reader = PyPDF2.PdfReader(file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text() or ""
        return text
    except Exception as e:
        return f"Error: {str(e)}"

def extract_text_from_docx(file):
    if not DOCX_AVAILABLE:
        return "DOCX support not installed. Please use PDF or TXT."
    try:
        doc = docx.Document(file)
        text = ""
        for para in doc.paragraphs:
            text += para.text + "\n"
        return text
    except Exception as e:
        return f"Error: {str(e)}"

def extract_text_from_txt(file):
    try:
        return file.getvalue().decode('utf-8')
    except:
        return file.getvalue().decode('latin-1')

# ==========================================================
# CLAUSE ANALYSIS FUNCTIONS
# ==========================================================

def analyze_contract_clauses(text):
    results = {}
    text_lower = text.lower()
    
    for clause_name, config in CLAUSE_PATTERNS.items():
        found = False
        preview = ""
        
        for pattern in config["patterns"]:
            match = re.search(pattern, text_lower, re.IGNORECASE)
            if match:
                found = True
                # Get context around the match
                start = max(0, match.start() - 50)
                end = min(len(text), match.end() + 100)
                preview = text[start:end].strip()
                break
        
        results[clause_name] = {
            "found": found,
            "risk_level": config["risk"] if found else "missing",
            "description": config["description"],
            "preview": preview
        }
    
    return results

def calculate_risk_score(results):
    scores = {
        "high": 75,
        "medium": 50,
        "low": 25,
        "missing": 0
    }
    
    total_score = 0
    for clause, data in results.items():
        total_score += scores[data["risk_level"]]
    
    # Average score
    avg_score = total_score / len(results) if results else 0
    return min(100, avg_score)

# ==========================================================
# AI FUNCTIONS
# ==========================================================

def improve_contract_with_ai(contract_text):
    if not client:
        return "⚠️ AI not available. Please configure GROQ_API_KEY."
    
    try:
        prompt = f"""You are a legal expert. Analyze and improve this contract to make it more balanced and legally safer.

Contract:
{contract_text[:8000]}

Instructions:
1. Identify risky clauses (payment terms, liability limits, termination, confidentiality)
2. Rewrite these clauses to be more balanced
3. Keep all legitimate business terms
4. Maintain the contract structure

Provide the improved contract:"""

        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"

def ask_chatbot(question, contract_text):
    if not client:
        return "⚠️ AI not available. Please configure GROQ_API_KEY."
    
    if not contract_text:
        return "Please upload a contract first."
    
    try:
        prompt = f"""You are a legal contract assistant. Answer questions about the contract below.

Contract:
{contract_text[:6000]}

Question: {question}

Provide a clear, helpful answer:"""

        response = client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        
        return response.choices[0].message.content
    except Exception as e:
        return f"Error: {str(e)}"

# ==========================================================
# EMAIL FUNCTION
# ==========================================================

def send_contract_email(recipient, subject, body, attachment_text=None):
    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        return False, "Email not configured. Please set SENDER_EMAIL and SENDER_PASSWORD in .env file."
    
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = recipient
        msg['Subject'] = subject
        
        # Create HTML body
        html_body = f"""
        <html>
        <body style="font-family: Arial, sans-serif; padding: 20px;">
            <h2 style="color: #667eea;">📋 RegulaAI Contract Delivery</h2>
            <p>Please find your contract below.</p>
            <hr>
            <pre style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; white-space: pre-wrap; word-wrap: break-word;">{body[:5000]}</pre>
            <hr>
            <p style="color: #666; font-size: 12px;">
                This email was sent from RegulaAI - AI-Powered Regulatory Compliance Platform.<br>
                Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            </p>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        if attachment_text:
            attachment = MIMEText(attachment_text, 'plain', 'utf-8')
            attachment.add_header('Content-Disposition', 'attachment', filename='contract.txt')
            msg.attach(attachment)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.sendmail(EMAIL_SENDER, recipient, msg.as_string())
        server.quit()
        
        return True, "Email sent successfully!"
    except smtplib.SMTPAuthenticationError:
        return False, "Email authentication failed. Please check SENDER_EMAIL and SENDER_PASSWORD in .env file. For Gmail, use App Password."
    except Exception as e:
        return False, f"Email error: {str(e)}"

# ==========================================================
# PDF GENERATION
# ==========================================================

def create_contract_pdf(contract_text, title="Contract"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=18, 
                                  spaceAfter=20, alignment=1)
    story.append(Paragraph(title, title_style))
    story.append(Spacer(1, 20))
    
    # Content
    content_style = ParagraphStyle('CustomContent', parent=styles['Normal'], fontSize=10,
                                   spaceAfter=12, leading=14)
    
    paragraphs = contract_text.split('\n\n')
    for para in paragraphs:
        if para.strip():
            story.append(Paragraph(para.strip(), content_style))
            story.append(Spacer(1, 10))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# ==========================================================
# VISUALIZATION FUNCTIONS
# ==========================================================

def create_risk_chart(results):
    risk_counts = {"High Risk": 0, "Medium Risk": 0, "Low Risk": 0, "Missing": 0}
    
    for clause, data in results.items():
        if data["risk_level"] == "high":
            risk_counts["High Risk"] += 1
        elif data["risk_level"] == "medium":
            risk_counts["Medium Risk"] += 1
        elif data["risk_level"] == "low":
            risk_counts["Low Risk"] += 1
        else:
            risk_counts["Missing"] += 1
    
    df = pd.DataFrame(list(risk_counts.items()), columns=['Category', 'Count'])
    
    colors = {'High Risk': '#ff4444', 'Medium Risk': '#ffaa00', 
              'Low Risk': '#44bb44', 'Missing': '#cccccc'}
    
    fig = px.bar(df, x='Category', y='Count', color='Category',
                 color_discrete_map=colors,
                 title="📊 Risk Distribution Across Clauses")
    
    fig.update_layout(showlegend=False, height=400)
    return fig

def create_risk_gauge(score):
    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = score,
        title = {'text': f"Overall Risk Score: {score}%"},
        gauge = {
            'axis': {'range': [0, 100], 'tickwidth': 1},
            'bar': {'color': "darkblue"},
            'bgcolor': "white",
            'steps': [
                {'range': [0, 30], 'color': '#d4edda'},
                {'range': [30, 60], 'color': '#fff3cd'},
                {'range': [60, 100], 'color': '#f8d7da'}
            ],
            'threshold': {
                'line': {'color': "red", 'width': 4},
                'thickness': 0.75,
                'value': score
            }
        }
    ))
    
    fig.update_layout(height=300, margin=dict(l=50, r=50, t=50, b=50))
    return fig

def create_clause_coverage_chart(results):
    found = sum(1 for r in results.values() if r["found"])
    missing = sum(1 for r in results.values() if not r["found"])
    
    fig = go.Figure(go.Pie(
        labels=['Clauses Found', 'Missing Clauses'],
        values=[found, missing],
        hole=0.4,
        marker=dict(colors=['#667eea', '#f0f0f0']
    )))
    
    fig.update_layout(title="📋 Clause Coverage", height=300)
    return fig

# ==========================================================
# MAIN APPLICATION
# ==========================================================

def main():
    # Page config
    st.set_page_config(
        page_title="RegulaAI - AI Contract Compliance",
        page_icon="📋",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Check login
    if not st.session_state.logged_in:
        login_page()
        return
    
    # Logout button in sidebar
    st.sidebar.markdown("---")
    st.sidebar.button("🚪 Logout", on_click=logout)
    st.sidebar.markdown(f"👤 Logged in as: {st.session_state.get('user_email', 'User')}")
    
    # Main title
    st.title("📋 RegulaAI - AI Contract Compliance System")
    st.caption("🤖 AI-Powered Contract Intelligence Platform")
    
    # Sidebar navigation
    st.sidebar.markdown("## 📌 Navigation")
    page = st.sidebar.radio("Go to:", [
        "📤 Upload Contract",
        "📊 Risk Analysis",
        "📈 Charts & Visualization",
        "✏️ Improve Contract",
        "📧 Email Delivery",
        "💬 AI Chat Assistant"
    ])
    
    st.markdown("---")
    
    # ==========================================================
    # UPLOAD CONTRACT PAGE
    # ==========================================================
    
    if page == "📤 Upload Contract":
        st.header("📤 Upload Contract")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader("Upload Methods")
            upload_method = st.radio("Choose:", ["📄 PDF File", "📝 DOCX File", "📃 TXT File", "✍️ Paste Text"],
                                     horizontal=True)
        
        with col2:
            st.subheader("Quick Stats")
            if st.session_state.contract_text:
                word_count = len(st.session_state.contract_text.split())
                char_count = len(st.session_state.contract_text)
                st.success(f"✅ Contract loaded: {word_count} words, {char_count} characters")
            else:
                st.info("ℹ️ No contract loaded yet")
        
        st.markdown("---")
        
        # File upload
        if upload_method == "📄 PDF File":
            file = st.file_uploader("Upload PDF", type=['pdf'])
            if file:
                st.session_state.contract_text = extract_text_from_pdf(file)
                st.success("✅ PDF uploaded successfully!")
                st.rerun()
                
        elif upload_method == "📝 DOCX File":
            if not DOCX_AVAILABLE:
                st.warning("⚠️ DOCX support not installed. Please use pip install python-docx")
            else:
                file = st.file_uploader("Upload DOCX", type=['docx'])
                if file:
                    st.session_state.contract_text = extract_text_from_docx(file)
                    st.success("✅ DOCX uploaded successfully!")
                    st.rerun()
                    
        elif upload_method == "📃 TXT File":
            file = st.file_uploader("Upload TXT", type=['txt'])
            if file:
                st.session_state.contract_text = extract_text_from_txt(file)
                st.success("✅ TXT uploaded successfully!")
                st.rerun()
                
        elif upload_method == "✍️ Paste Text":
            text = st.text_area("Paste contract text here:", height=300)
            if text:
                st.session_state.contract_text = text
                st.success("✅ Contract text saved!")
                st.rerun()
        
        # Show preview
        if st.session_state.contract_text:
            with st.expander("📄 Contract Preview"):
                st.text_area("", st.session_state.contract_text[:2000] + "..." if len(st.session_state.contract_text) > 2000 else st.session_state.contract_text,
                            height=300, key="preview")
    
    # ==========================================================
    # RISK ANALYSIS PAGE
    # ==========================================================
    
    elif page == "📊 Risk Analysis":
        st.header("📊 Contract Risk Analysis")
        
        if not st.session_state.contract_text:
            st.warning("⚠️ Please upload a contract first!")
            return
        
        # Analyze button
        if st.button("🔍 Analyze Contract", type="primary"):
            with st.spinner("Analyzing contract clauses..."):
                results = analyze_contract_clauses(st.session_state.contract_text)
                risk_score = calculate_risk_score(results)
                
                st.session_state.analysis_results = {
                    "results": results,
                    "risk_score": risk_score
                }
        
        # Display results
        if st.session_state.analysis_results:
            results = st.session_state.analysis_results["results"]
            risk_score = st.session_state.analysis_results["risk_score"]
            
            st.markdown("---")
            
            # Overall risk score
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Overall Risk Score", f"{risk_score}%")
            with col2:
                found_count = sum(1 for r in results.values() if r["found"])
                st.metric("Clauses Found", f"{found_count}/{len(results)}")
            with col3:
                missing_count = sum(1 for r in results.values() if not r["found"])
                st.metric("Missing Clauses", missing_count)
            
            st.markdown("---")
            
            # Detailed clause analysis
            st.subheader("📋 Clause Analysis")
            
            for clause_name, data in results.items():
                risk = data["risk_level"]
                config = RISK_CONFIG[risk]
                
                with st.expander(f"{config['color']} {clause_name} - {config['level']}"):
                    col1, col2 = st.columns([1, 2])
                    
                    with col1:
                        st.write(f"**Status:** {'✅ Found' if data['found'] else '❌ Missing'}")
                        st.write(f"**Risk Level:** {config['color']} {config['level']}")
                        st.write(f"**Description:** {data['description']}")
                    
                    with col2:
                        if data["preview"]:
                            st.write("**Preview:**")
                            st.info(f"...{data['preview']}...")
    
    # ==========================================================
    # CHARTS PAGE
    # ==========================================================
    
    elif page == "📈 Charts & Visualization":
        st.header("📈 Risk Visualization Dashboard")
        
        if not st.session_state.analysis_results:
            st.warning("⚠️ Please analyze a contract first!")
            return
        
        results = st.session_state.analysis_results["results"]
        risk_score = st.session_state.analysis_results["risk_score"]
        
        # Risk gauge
        st.subheader("🎯 Overall Risk Assessment")
        gauge = create_risk_gauge(risk_score)
        st.plotly_chart(gauge, use_container_width=True)
        
        st.markdown("---")
        
        # Two columns of charts
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📊 Risk Distribution")
            bar_chart = create_risk_chart(results)
            st.plotly_chart(bar_chart, use_container_width=True)
            
        with col2:
            st.subheader("📋 Clause Coverage")
            pie_chart = create_clause_coverage_chart(results)
            st.plotly_chart(pie_chart, use_container_width=True)
        
        st.markdown("---")
        
        # Risk matrix
        st.subheader("⚠️ Risk Matrix (Likelihood vs Impact)")
        
        # Create sample risk matrix data
        risk_data = []
        for clause, data in results.items():
            risk_data.append({
                "Clause": clause,
                "Severity": {"high": 3, "medium": 2, "low": 1, "missing": 0}[data["risk_level"]],
                "Found": 1 if data["found"] else 0
            })
        
        df_risk = pd.DataFrame(risk_data)
        
        if not df_risk.empty:
            fig_matrix = px.scatter(
                df_risk, x="Found", y="Severity", size="Severity", color="Clause",
                title="Clause Risk Matrix", hover_name="Clause"
            )
            fig_matrix.update_layout(
                xaxis_title="Clause Found",
                yaxis_title="Severity",
                xaxis=dict(ticktext=["Missing", "Found"], tickvals=[0, 1])
            )
            st.plotly_chart(fig_matrix, use_container_width=True)
    
    # ==========================================================
    # IMPROVE CONTRACT PAGE
    # ==========================================================
    
    elif page == "✏️ Improve Contract":
        st.header("✏️ AI Contract Amendment")
        
        if not st.session_state.contract_text:
            st.warning("⚠️ Please upload a contract first!")
            return
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.subheader("📄 Original Contract")
            st.text_area("", st.session_state.contract_text, height=300, key="original_display")
            
        with col2:
            st.subheader("⚡ Actions")
            if st.button("🤖 Generate Improved Contract", type="primary", use_container_width=True):
                with st.spinner("🤖 AI is improving your contract..."):
                    improved = improve_contract_with_ai(st.session_state.contract_text)
                    st.session_state.improved_contract = improved
            
            if st.session_state.improved_contract:
                st.success("✅ Improved contract generated!")
                
                # Download options
                st.markdown("### 💾 Download")
                
                # Text download
                st.download_button(
                    "📥 Download as TXT",
                    st.session_state.improved_contract,
                    "improved_contract.txt",
                    "text/plain",
                    use_container_width=True
                )
                
                # PDF download
                pdf_buffer = create_contract_pdf(st.session_state.improved_contract, "Improved Contract")
                st.download_button(
                    "📥 Download as PDF",
                    pdf_buffer,
                    "improved_contract.pdf",
                    "application/pdf",
                    use_container_width=True
                )
        
        # Show improved contract
        if st.session_state.improved_contract:
            st.markdown("---")
            st.subheader("✨ Improved Contract")
            
            # Edit option
            edited = st.text_area("Edit improved contract:", st.session_state.improved_contract, height=400)
            if edited != st.session_state.improved_contract:
                st.session_state.improved_contract = edited
    
    # ==========================================================
    # EMAIL DELIVERY PAGE
    # ==========================================================
    
    elif page == "📧 Email Delivery":
        st.header("📧 Email Contract Delivery")
        
        col1, col2 = st.columns([1, 1])
        
        with col1:
            st.subheader("📝 Email Details")
            recipient = st.text_input("Recipient Email", placeholder="recipient@example.com")
            subject = st.text_input("Subject", "Your Contract from RegulaAI")
            
            # Choose which contract to send
            contract_choice = st.radio("Contract to send:", 
                                       ["Original Contract", "Improved Contract"],
                                       horizontal=True)
            
            if contract_choice == "Original Contract":
                body_text = st.session_state.contract_text
            else:
                body_text = st.session_state.improved_contract or "No improved contract available"
        
        with col2:
            st.subheader("📤 Preview")
            if body_text:
                st.text_area("Email Body Preview:", body_text[:1000] + "..." if len(body_text) > 1000 else body_text,
                            height=200)
        
        st.markdown("---")
        
        if st.button("📧 Send Email", type="primary"):
            if not recipient:
                st.error("❌ Please enter a recipient email")
            elif "@" not in recipient:
                st.error("❌ Invalid email address")
            else:
                with st.spinner("Sending email..."):
                    attachment = body_text if contract_choice == "Improved Contract" and st.session_state.improved_contract else None
                    success, message = send_contract_email(recipient, subject, body_text, attachment)
                    
                    if success:
                        st.success(f"✅ {message}")
                    else:
                        st.error(f"❌ {message}")
    
    # ==========================================================
    # AI CHAT ASSISTANT PAGE
    # ==========================================================
    
    elif page == "💬 AI Chat Assistant":
        st.header("💬 AI Contract Chat Assistant")
        
        if not st.session_state.contract_text:
            st.warning("⚠️ Please upload a contract first!")
            return
        
        # Sample questions
        st.subheader("💡 Sample Questions")
        sample_questions = [
            "What are the main risks in this contract?",
            "Who are the parties involved?",
            "What are the payment terms?",
            "What are the termination conditions?",
            "What does the confidentiality clause say?",
            "What is the liability limitation?",
            "Are there any concerning clauses?",
            "What happens in case of breach?"
        ]
        
        cols = st.columns(4)
        for i, q in enumerate(sample_questions):
            if cols[i % 4].button(q, key=f"sample_{i}"):
                # Add to chat
                with st.spinner("Thinking..."):
                    answer = ask_chatbot(q, st.session_state.contract_text)
                    st.session_state.chat_history.append({"question": q, "answer": answer})
        
        st.markdown("---")
        
        # Chat input
        st.subheader("💬 Ask Your Question")
        
        with st.form("chat_form"):
            question = st.text_input("Ask about your contract:", placeholder="Type your question here...")
            submit = st.form_submit_button("Ask", type="primary")
        
        if submit and question:
            with st.spinner("Thinking..."):
                answer = ask_chatbot(question, st.session_state.contract_text)
                st.session_state.chat_history.append({"question": question, "answer": answer})
        
        st.markdown("---")
        
        # Chat history
        st.subheader("💭 Chat History")
        
        if st.button("🧹 Clear Chat History"):
            st.session_state.chat_history = []
            st.rerun()
        
        # Display messages
        for i, msg in enumerate(st.session_state.chat_history):
            with st.container():
                st.markdown(f"**👤 You:** {msg['question']}")
                st.markdown(f"**🤖 AI:** {msg['answer']}")
                st.markdown("---")

# Run the app
if __name__ == "__main__":
    main()

