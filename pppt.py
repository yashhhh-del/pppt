import streamlit as st
import requests
import base64
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image
import time
import json
import pandas as pd
import zipfile
import matplotlib.pyplot as plt
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

# Page configuration
st.set_page_config(
    page_title="AI PowerPoint Generator Pro",
    page_icon="📊",
    layout="wide"
)

# Initialize session state
if 'generation_count' not in st.session_state:
    st.session_state.generation_count = 0
    st.session_state.total_slides = 0
    st.session_state.slides_content = None
    st.session_state.edited_slides = None
    st.session_state.final_pptx = None

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        font-weight: bold;
        padding: 0.5rem;
        border-radius: 5px;
    }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .download-section {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin: 2rem 0;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">📊 AI PowerPoint Generator Pro</div>', unsafe_allow_html=True)
st.markdown("---")

# Sidebar
with st.sidebar:
    st.header("⚙️ API Configuration")
    
    claude_api_key = st.text_input("OpenRouter API Key *", type="password", 
                                    help="Required: For generating presentation content")
    
    model_choice = st.selectbox(
        "AI Model",
        [
            "Free Model (Google Gemini Flash)",
            "Free Model (Meta Llama 3.2)",
            "Free Model (Mistral 7B)",
            "Claude 3.5 Sonnet (Paid)"
        ],
        help="Try different free models if one is rate-limited"
    )
    
    if "Free" in model_choice:
        st.info("💡 Free models share rate limits. Switch models if limited.")
    
    st.info("💡 Using OpenRouter API")
    
    st.markdown("---")
    
    # Image API Configuration
    st.subheader("🖼️ Image Configuration")
    
    # Google API Key (pre-filled)
    google_api_key = st.text_input(
        "Google API Key *", 
        value="AIzaSyB8BKP0m1r6_cuB3byyxfUwsSiGtrRPMFI",
        type="password",
        help="Google Custom Search API Key"
    )
    
    # Custom Search Engine ID
    google_cx = st.text_input(
        "Google Search Engine ID *",
        type="password",
        help="Get it from: https://programmablesearchengine.google.com/"
    )
    
    if google_api_key and google_cx:
        st.success("✅ Google Image Search configured!")
        st.info("💡 Using Google Custom Search API for images")
    else:
        st.warning("⚠️ Need Search Engine ID for images")
    
    # Image fallback options
    st.markdown("### 🔄 Fallback Image Sources")
    use_unsplash_fallback = st.checkbox("Use Unsplash as fallback", value=True)
    use_pexels_fallback = st.checkbox("Use Pexels as fallback", value=False)
    
    if use_pexels_fallback:
        pexels_api_key = st.text_input(
            "Pexels API Key (Optional)", 
            type="password",
            help="FREE! Get it at: https://www.pexels.com/api/"
        )
    else:
        pexels_api_key = None
    
    st.markdown("---")
    
    # Logo Upload
    st.subheader("🏢 Branding")
    logo_file = st.file_uploader("Upload Company Logo", type=["png", "jpg", "jpeg"])
    logo_data = None
    if logo_file:
        logo_data = logo_file.read()
        st.success("✅ Logo uploaded!")
    
    st.markdown("---")
    
    # Usage Analytics
    st.markdown("### 📊 Your Stats")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Presentations", st.session_state.generation_count)
    with col2:
        st.metric("Total Slides", st.session_state.total_slides)
    
    st.markdown("---")
    st.markdown("### 📖 How to Use")
    st.markdown("""
    1. Enter OpenRouter API key
    2. Add Google Search Engine ID
    3. Enter your presentation topic
    4. Click Generate!
    5. Download immediately!
    """)
    st.markdown("---")
    st.markdown("### 🔗 Get API Keys")
    st.markdown("🔑 [Google Custom Search](https://programmablesearchengine.google.com/)")
    st.markdown("🆓 [Pexels API (FREE)](https://www.pexels.com/api/)")
    st.markdown("[OpenRouter API](https://openrouter.ai/keys)")

# ============ IMAGE FUNCTIONS ============

def generate_topic_search_terms(main_topic, slide_title, image_prompt):
    """Generate search terms prioritizing topic relevance"""
    search_terms = []
    
    # 1. AI's specific image prompt
    if image_prompt and image_prompt.strip():
        search_terms.append(image_prompt.strip())
    
    # 2. Topic + slide title combined
    if main_topic and slide_title:
        search_terms.append(f"{main_topic} {slide_title}")
    
    # 3. Just slide title
    if slide_title:
        search_terms.append(slide_title)
    
    # 4. Just main topic
    if main_topic:
        search_terms.append(main_topic)
    
    # Remove duplicates
    seen = set()
    unique = []
    for term in search_terms:
        lower = term.lower().strip()
        if lower and lower not in seen:
            seen.add(lower)
            unique.append(term)
    
    return unique

# ============ GOOGLE IMAGE SEARCH ============

def get_google_image(query, api_key, cx):
    """Get image using Google Custom Search API"""
    try:
        url = "https://www.googleapis.com/customsearch/v1"
        
        params = {
            'key': api_key,
            'cx': cx,
            'q': query,
            'searchType': 'image',
            'num': 3,
            'imgSize': 'large',
            'imgType': 'photo',
            'safe': 'active',
            'fileType': 'jpg,png'
        }
        
        response = requests.get(url, params=params, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            
            if 'items' in data and len(data['items']) > 0:
                # Try to get image from first result
                for item in data['items'][:3]:  # Try first 3 results
                    try:
                        image_url = item['link']
                        img_response = requests.get(image_url, timeout=10, headers={
                            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                        })
                        
                        if img_response.status_code == 200 and len(img_response.content) > 5000:
                            # Validate it's actually an image
                            img = Image.open(io.BytesIO(img_response.content))
                            if img.size[0] > 300 and img.size[1] > 200:
                                return img_response.content
                    except:
                        continue
        
        return None
    except Exception as e:
        st.write(f"      ⚠️ Google Search error: {str(e)}")
        return None

# ============ FALLBACK IMAGE SOURCES ============

def get_unsplash_image(query, width=800, height=600):
    """Get image from Unsplash Direct (Free)"""
    try:
        clean_query = query.strip().replace(' ', ',')
        url = f"https://source.unsplash.com/{width}x{height}/?{clean_query}"
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(url, timeout=15, allow_redirects=True, headers=headers)
        
        if response.status_code == 200 and len(response.content) > 5000:
            try:
                img = Image.open(io.BytesIO(response.content))
                if img.size[0] > 400 and img.size[1] > 300:
                    return response.content
            except:
                pass
        return None
    except:
        return None

def get_pexels_image(query, api_key):
    """Get image from Pexels Direct API"""
    if not api_key:
        return None
    
    try:
        url = "https://api.pexels.com/v1/search"
        headers = {"Authorization": api_key}
        params = {
            "query": query,
            "per_page": 3,
            "orientation": "landscape"
        }
        
        response = requests.get(url, headers=headers, params=params, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            if data.get("photos"):
                photo = data["photos"][0]
                img_url = photo["src"]["large"]
                
                img_response = requests.get(img_url, timeout=10)
                if img_response.status_code == 200:
                    return img_response.content
        return None
    except:
        return None

# ============ UNIFIED IMAGE RETRIEVAL ============

def get_topic_relevant_image(main_topic, slide_title, image_prompt, google_api_key, google_cx, use_unsplash, use_pexels, pexels_key):
    """Get highly relevant image using Google + fallbacks"""
    
    st.write(f"   🎯 Topic: {main_topic}")
    st.write(f"   📄 Slide: {slide_title}")
    
    # Generate search terms
    search_terms = generate_topic_search_terms(main_topic, slide_title, image_prompt)
    st.write(f"   🔍 Will try {len(search_terms)} search variations")
    
    # Try each search term
    for i, term in enumerate(search_terms, 1):
        st.write(f"      → Search {i}: '{term}'")
        
        # Try Google Image Search first
        if google_api_key and google_cx:
            st.write(f"         🔍 Searching Google...")
            image_data = get_google_image(term, google_api_key, google_cx)
            if image_data:
                st.write(f"      ✅ Found on Google!")
                return image_data
        
        # Try Pexels fallback
        if use_pexels and pexels_key:
            st.write(f"         🔍 Trying Pexels fallback...")
            image_data = get_pexels_image(term, pexels_key)
            if image_data:
                st.write(f"      ✅ Found on Pexels!")
                return image_data
        
        # Try Unsplash fallback
        if use_unsplash:
            st.write(f"         🔍 Trying Unsplash fallback...")
            image_data = get_unsplash_image(term)
            if image_data:
                st.write(f"      ✅ Found on Unsplash!")
                return image_data
    
    # Final fallback to generic topic
    st.write(f"   🆘 Trying generic fallback...")
    fallback = main_topic.split()[0] if main_topic else "business"
    
    if google_api_key and google_cx:
        image_data = get_google_image(fallback, google_api_key, google_cx)
        if image_data:
            st.write(f"      ✅ Got fallback from Google")
            return image_data
    
    if use_unsplash:
        image_data = get_unsplash_image(fallback)
        if image_data:
            st.write(f"      ✅ Got fallback from Unsplash")
            return image_data
    
    return None

# ============ CONTENT GENERATION ============

def generate_content_with_claude(api_key, topic, category, slide_count, tone, audience, key_points, model_choice, language):
    """Generate presentation content using AI"""
    try:
        # Model selection logic
        if "Gemini" in model_choice:
            model = "google/gemini-2.0-flash-exp:free"
        elif "Llama" in model_choice:
            model = "meta-llama/llama-3.2-3b-instruct:free"
        elif "Mistral" in model_choice:
            model = "mistralai/mistral-7b-instruct:free"
        else:
            model = "anthropic/claude-3.5-sonnet"
        
        calculated_tokens = min(slide_count * 200 + 300, 2000)
        
        language_instruction = f"Generate ALL content in {language} language." if language != "English" else ""
        
        prompt = f"""{language_instruction}
You are an expert presentation creator. Generate a PowerPoint structure about: {topic}

Category: {category}
Slides: {slide_count}
Tone: {tone}
Audience: {audience}
Key Points: {key_points if key_points else "None"}

Return ONLY valid JSON in this format:
{{
  "slides": [
    {{
      "title": "Presentation Title",
      "bullets": [],
      "image_prompt": "{topic} title image",
      "speaker_notes": "Opening remarks and introduction"
    }},
    {{
      "title": "Slide Title",
      "bullets": ["Point 1", "Point 2", "Point 3"],
      "image_prompt": "specific image description related to {topic}",
      "speaker_notes": "What to explain during this slide"
    }}
  ]
}}

CRITICAL REQUIREMENTS:
1. image_prompt must be specific to {topic}
2. Include detailed speaker_notes for each slide
3. Make content appropriate for {audience}
4. Total slides: exactly {slide_count}
5. ALL text content must be in {language}

Return ONLY JSON, no markdown."""

        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key.strip()}",
                "Content-Type": "application/json",
            },
            json={
                "model": model,
                "max_tokens": calculated_tokens,
                "messages": [{"role": "user", "content": prompt}]
            },
            timeout=30
        )
        
        if response.status_code == 200:
            data = response.json()
            content_text = data["choices"][0]["message"]["content"]
            
            # Clean JSON
            content_text = content_text.strip()
            if content_text.startswith("```json"):
                content_text = content_text[7:]
            if content_text.startswith("```"):
                content_text = content_text[3:]
            if content_text.endswith("```"):
                content_text = content_text[:-3]
            content_text = content_text.strip()
            
            slides_data = json.loads(content_text)
            return slides_data["slides"]
        else:
            # Enhanced error handling
            if response.status_code == 429:
                st.error(f"⏱️ Rate Limit: Model is temporarily unavailable")
                st.info("💡 **Solutions:**\n- Wait 30-60 seconds and try again\n- Switch to a different free model above\n- Use Claude 3.5 Sonnet (paid but reliable)")
                raise Exception("Rate limit - retry needed")
            elif response.status_code == 402:
                st.error("💳 Insufficient credits! Reduce slides or add credits.")
            else:
                st.error(f"API Error: {response.text}")
            return None
            
    except json.JSONDecodeError as e:
        st.error(f"JSON parsing error: {str(e)}")
        return None
    except Exception as e:
        if "Rate limit" in str(e):
            raise
        st.error(f"Error: {str(e)}")
        return None

def generate_content_with_retry(api_key, topic, category, slide_count, tone, audience, key_points, model_choice, language, max_retries=3):
    """Generate content with automatic retry on rate limit"""
    for attempt in range(max_retries):
        try:
            result = generate_content_with_claude(api_key, topic, category, slide_count, tone, audience, key_points, model_choice, language)
            if result:
                return result
        except Exception as e:
            if "Rate limit" in str(e) or "429" in str(e):
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 5
                    st.warning(f"⏳ Rate limit hit. Retrying in {wait_time} seconds... (Attempt {attempt + 2}/{max_retries})")
                    time.sleep(wait_time)
                else:
                    st.error("❌ Rate limit persists after retries.")
                    return None
            else:
                return None
    return None

# ============ ANALYSIS FUNCTIONS ============

def analyze_presentation(slides_content):
    """Analyze presentation quality"""
    issues = []
    suggestions = []
    score = 100
    
    for i, slide in enumerate(slides_content, 1):
        bullet_count = len(slide.get('bullets', []))
        if bullet_count > 5:
            issues.append(f"Slide {i}: Too many bullets ({bullet_count})")
            suggestions.append(f"Slide {i}: Reduce to 3-5 key points")
            score -= 5
        
        title_len = len(slide['title'])
        if title_len > 60:
            issues.append(f"Slide {i}: Title too long ({title_len} chars)")
            suggestions.append(f"Slide {i}: Shorten title to <60 characters")
            score -= 3
        
        total_text = sum(len(b) for b in slide.get('bullets', []))
        if total_text > 500:
            issues.append(f"Slide {i}: Too much text ({total_text} chars)")
            suggestions.append(f"Slide {i}: Use more visuals, less text")
            score -= 5
    
    score = max(0, score)
    return issues, suggestions, score

def generate_slide_preview(slide_data, theme_colors):
    """Generate preview image of slide"""
    fig, ax = plt.subplots(figsize=(10, 7.5), facecolor='#' + ''.join([format(c, '02x') for c in theme_colors['bg']]))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7.5)
    ax.axis('off')
    
    title_color = '#' + ''.join([format(c, '02x') for c in theme_colors['accent']])
    ax.text(5, 6.5, slide_data['title'][:50], ha='center', fontsize=18, weight='bold', color=title_color)
    
    text_color = '#' + ''.join([format(c, '02x') for c in theme_colors['text']])
    y_pos = 5.5
    for bullet in slide_data.get('bullets', [])[:5]:
        bullet_text = bullet[:60] + "..." if len(bullet) > 60 else bullet
        ax.text(1, y_pos, f"• {bullet_text}", fontsize=10, color=text_color)
        y_pos -= 0.6
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=80, facecolor=fig.get_facecolor())
    plt.close(fig)
    buf.seek(0)
    return buf

# ============ EXPORT FUNCTIONS ============

def export_to_pdf(slides_content, topic):
    """Export to PDF"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    title = Paragraph(f"<b>{topic}</b>", styles['Title'])
    story.append(title)
    story.append(Spacer(1, 12))
    
    for i, slide in enumerate(slides_content, 1):
        slide_title = Paragraph(f"<b>Slide {i}: {slide['title']}</b>", styles['Heading2'])
        story.append(slide_title)
        story.append(Spacer(1, 6))
        
        for bullet in slide.get('bullets', []):
            bullet_text = Paragraph(f"• {bullet}", styles['BodyText'])
            story.append(bullet_text)
        
        if slide.get('speaker_notes'):
            notes = Paragraph(f"<i>Notes: {slide['speaker_notes']}</i>", styles['Italic'])
            story.append(Spacer(1, 6))
            story.append(notes)
        
        story.append(Spacer(1, 20))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

def export_to_google_slides_json(slides_content, topic, theme):
    """Export to Google Slides JSON"""
    google_slides_data = {
        "title": topic,
        "theme": theme,
        "slides": []
    }
    
    for slide in slides_content:
        google_slide = {
            "title": slide['title'],
            "content": slide.get('bullets', []),
            "notes": slide.get('speaker_notes', ''),
            "imagePrompt": slide.get('image_prompt', '')
        }
        google_slides_data['slides'].append(google_slide)
    
    return json.dumps(google_slides_data, indent=2)

# ============ POWERPOINT CREATION ============

def create_powerpoint(slides_content, theme, image_mode, google_api_key, google_cx, use_unsplash, use_pexels, pexels_key, category, audience, topic, image_position, logo_data):
    """Create PowerPoint presentation"""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    themes = {
        "Corporate Blue": {"bg": RGBColor(240, 248, 255), "accent": RGBColor(31, 119, 180), "text": RGBColor(0, 0, 0)},
        "Gradient Modern": {"bg": RGBColor(240, 242, 246), "accent": RGBColor(138, 43, 226), "text": RGBColor(0, 0, 0)},
        "Minimal Dark": {"bg": RGBColor(30, 30, 30), "accent": RGBColor(255, 215, 0), "text": RGBColor(255, 255, 255)},
        "Pastel Soft": {"bg": RGBColor(255, 250, 240), "accent": RGBColor(255, 182, 193), "text": RGBColor(60, 60, 60)},
        "Professional Green": {"bg": RGBColor(245, 255, 250), "accent": RGBColor(34, 139, 34), "text": RGBColor(0, 0, 0)},
        "Elegant Purple": {"bg": RGBColor(250, 245, 255), "accent": RGBColor(128, 0, 128), "text": RGBColor(0, 0, 0)}
    }
    
    color_scheme = themes.get(theme, themes["Corporate Blue"])
    
    positions = {
        "Right Side": {"left": Inches(6.5), "top": Inches(2), "width": Inches(3)},
        "Left Side": {"left": Inches(0.5), "top": Inches(2), "width": Inches(3)},
        "Top Right Corner": {"left": Inches(8), "top": Inches(0.5), "width": Inches(1.5)},
        "Bottom": {"left": Inches(3.5), "top": Inches(5.5), "width": Inches(3)},
        "Center": {"left": Inches(3.5), "top": Inches(2.5), "width": Inches(3)}
    }
    
    img_pos = positions.get(image_position, positions["Right Side"])
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, slide_data in enumerate(slides_content):
        status_text.text(f"Creating slide {idx + 1}/{len(slides_content)}...")
        progress_bar.progress((idx + 1) / len(slides_content))
        
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = color_scheme["bg"]
        
        if logo_data:
            try:
                logo_stream = io.BytesIO(logo_data)
                slide.shapes.add_picture(logo_stream, Inches(9), Inches(0.2), width=Inches(0.8))
            except:
                pass
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8.5), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = slide_data["title"]
        title_frame.paragraphs[0].font.size = Pt(36 if idx == 0 else 28)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = color_scheme["accent"]
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER if idx == 0 else PP_ALIGN.LEFT
        
        if idx > 0 and slide_data.get("bullets"):
            bullet_width = Inches(5.5) if image_mode == "With Images" else Inches(9)
            bullet_box = slide.shapes.add_textbox(Inches(0.5), Inches(2), bullet_width, Inches(4.5))
            text_frame = bullet_box.text_frame
            text_frame.word_wrap = True
            
            for bullet in slide_data["bullets"]:
                p = text_frame.add_paragraph()
                p.text = bullet
                p.level = 0
                p.font.size = Pt(18)
                p.font.color.rgb = color_scheme["text"]
                p.space_after = Pt(12)
        
        if idx == 0:
            subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3), Inches(9), Inches(1))
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = f"{category} Presentation | {audience}"
            subtitle_frame.paragraphs[0].font.size = Pt(20)
            subtitle_frame.paragraphs[0].font.color.rgb = color_scheme["text"]
            subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        if slide_data.get("speaker_notes"):
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = slide_data["speaker_notes"]
        
        if idx > 0 and image_mode == "With Images":
            with st.expander(f"🖼️ Slide {idx + 1}: {slide_data['title']}", expanded=False):
                image_prompt = slide_data.get("image_prompt", "")
                
                image_data = get_topic_relevant_image(
                    main_topic=topic,
                    slide_title=slide_data["title"],
                    image_prompt=image_prompt,
                    google_api_key=google_api_key,
                    google_cx=google_cx,
                    use_unsplash=use_unsplash,
                    use_pexels=use_pexels,
                    pexels_key=pexels_key
                )
                
                if image_data:
                    try:
                        image_stream = io.BytesIO(image_data)
                        slide.shapes.add_picture(
                            image_stream, 
                            img_pos["left"], 
                            img_pos["top"], 
                            width=img_pos["width"]
                        )
                        st.success(f"   ✅ Image added!")
                    except Exception as e:
                        st.error(f"   ❌ Failed: {str(e)}")
                else:
                    st.warning(f"   ⚠️ No image found")
            
            time.sleep(0.5)
    
    progress_bar.progress(1.0)
    status_text.text("✅ Presentation created!")
    
    return prs

# ============ TEMPLATE FUNCTIONS ============

def save_template(category, slide_count, tone, audience, theme, image_mode, language):
    """Save template"""
    template = {
        "category": category,
        "slide_count": slide_count,
        "tone": tone,
        "audience": audience,
        "theme": theme,
        "image_mode": image_mode,
        "language": language
    }
    return json.dumps(template, indent=2)

def load_template(template_file):
    """Load template"""
    try:
        template = json.loads(template_file.read())
        return template
    except:
        return None

# ============ MAIN UI ============

tab1, tab2, tab3 = st.tabs(["📝 Create Presentation", "📊 Bulk Generate", "⚙️ Templates"])

with tab1:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📝 Your Topic")
        topic = st.text_input("Enter Topic *", placeholder="e.g., Space Exploration, Digital Marketing...")
        st.caption("💡 Be specific for better images!")
        
        category = st.selectbox("Category *", ["Business", "Pitch", "Marketing", "Technical", "Academic", "Training", "Sales"])
        
        col1_1, col1_2 = st.columns(2)
        with col1_1:
            slide_count = st.number_input("Slides *", min_value=3, max_value=20, value=6)
        with col1_2:
            language = st.selectbox("Language 🌍", ["English", "Hindi (हिंदी)", "Spanish", "French", "German"])
        
        tone = st.selectbox("Tone *", ["Formal", "Neutral", "Inspirational", "Educational", "Persuasive"])

    with col2:
        st.subheader("🎨 Style & Images")
        audience = st.selectbox("Audience *", ["Investors", "Students", "Corporate", "Clients", "Managers"])
        theme = st.selectbox("Theme *", ["Corporate Blue", "Gradient Modern", "Minimal Dark", "Pastel Soft", "Professional Green", "Elegant Purple"])
        
        image_mode = st.selectbox("Image Mode *", ["With Images", "No Images"])
        
        if image_mode == "With Images":
            image_position = st.selectbox("Position", ["Right Side", "Left Side", "Top Right Corner", "Bottom", "Center"])
        else:
            image_position = "Right Side"

    st.subheader("➕ Additional Points (Optional)")
    key_points = st.text_area("Key points", placeholder="- Point 1\n- Point 2", height=80)

    export_format = st.selectbox("Export Format", ["PowerPoint (.pptx)", "PowerPoint + PDF", "Google Slides (JSON)"])

    st.markdown("---")
    
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        generate_button = st.button("🚀 Generate PowerPoint", use_container_width=True, type="primary")
    with col_btn2:
        save_template_btn = st.button("💾 Template", use_container_width=True)

    if save_template_btn:
        template_json = save_template(category, slide_count, tone, audience, theme, image_mode, language)
        st.download_button("📥 Download", template_json, "template.json", "application/json")
        st.success("✅ Ready!")

    if generate_button:
        if not claude_api_key:
            st.error("⚠️ Enter OpenRouter API key")
        elif not topic:
            st.error("⚠️ Enter a topic")
        elif image_mode == "With Images" and not google_cx:
            st.error("⚠️ Enter Google Search Engine ID for images")
        else:
            with st.spinner("🤖 Generating..."):
                slides_content = generate_content_with_retry(
                    claude_api_key, topic, category, slide_count, 
                    tone, audience, key_points, model_choice, language
                )
                
                if slides_content:
                    st.session_state.slides_content = slides_content
                    st.session_state.generation_count += 1
                    st.session_state.total_slides += len(slides_content)
                    
                    st.success("✅ Creating presentation...")
                    
                    prs = create_powerpoint(
                        slides_content, theme, image_mode,
                        google_api_key, google_cx,
                        use_unsplash_fallback, use_pexels_fallback, 
                        pexels_api_key if use_pexels_fallback else None,
                        category, audience, topic, 
                        image_position, logo_data
                    )
                    
                    pptx_io = io.BytesIO()
                    prs.save(pptx_io)
                    pptx_io.seek(0)
                    st.session_state.final_pptx = pptx_io.getvalue()
                    
                    st.success("🎉 Ready!")
                    
                    st.markdown("---")
                    st.markdown('<div class="download-section">', unsafe_allow_html=True)
                    st.markdown("### 🎉 Your Presentation is Ready!")
                    
                    if export_format == "PowerPoint (.pptx)":
                        col_dl = st.columns([1, 2, 1])
                        with col_dl[1]:
                            st.download_button(
                                label="📥 DOWNLOAD NOW",
                                data=st.session_state.final_pptx,
                                file_name=f"{topic.replace(' ', '_')}.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                                type="primary"
                            )
                    
                    elif export_format == "PowerPoint + PDF":
                        col_dl1, col_dl2 = st.columns(2)
                        with col_dl1:
                            st.download_button(
                                "📥 PowerPoint",
                                st.session_state.final_pptx,
                                f"{topic.replace(' ', '_')}.pptx",
                                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                                type="primary"
                            )
                        with col_dl2:
                            pdf_buffer = export_to_pdf(slides_content, topic)
                            st.download_button(
                                "📄 PDF",
                                pdf_buffer,
                                f"{topic.replace(' ', '_')}.pdf",
                                "application/pdf",
                                use_container_width=True,
                                type="primary"
                            )
                    
                    elif export_format == "Google Slides (JSON)":
                        google_json = export_to_google_slides_json(slides_content, topic, theme)
                        col_dl = st.columns([1, 2, 1])
                        with col_dl[1]:
                            st.download_button(
                                "📥 Download JSON",
                                google_json,
                                f"{topic.replace(' ', '_')}.json",
                                "application/json",
                                use_container_width=True,
                                type="primary"
                            )
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                    st.markdown("---")
                    
                    # Stats
                    st.subheader("📊 Stats")
                    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
                    
                    with col_stat1:
                        st.metric("Slides", len(slides_content))
                    with col_stat2:
                        total_words = 0
                        for s in slides_content:
                            bullets = s.get('bullets', [])
                            if bullets and isinstance(bullets, list):
                                for bullet in bullets:
                                    if isinstance(bullet, str):
                                        total_words += len(bullet.split())
                        st.metric("Words", total_words)
                    with col_stat3:
                        bullet_counts = [len(s.get('bullets', [])) for s in slides_content if isinstance(s.get('bullets', []), list)]
                        avg_bullets = sum(bullet_counts) / len(bullet_counts) if bullet_counts else 0
                        st.metric("Avg Bullets", f"{avg_bullets:.1f}")
                    with col_stat4:
                        est_time = len(slides_content) * 2
                        st.metric("Est. Time", f"{est_time} min")
                    
                    # Coach
                    with st.expander("🎓 AI Coach", expanded=False):
                        issues, suggestions, score = analyze_presentation(slides_content)
                        col_score1, col_score2 = st.columns([1, 3])
                        with col_score1:
                            score_color = "🟢" if score >= 80 else "🟡" if score >= 60 else "🔴"
                            st.markdown(f"### {score_color} {score}/100")
                        with col_score2:
                            if score >= 80:
                                st.success("Excellent!")
                            elif score >= 60:
                                st.warning("Good!")
                            else:
                                st.error("Needs work!")
                        
                        if suggestions:
                            for suggestion in suggestions:
                                st.write(f"- {suggestion}")

with tab2:
    st.subheader("📊 Bulk Generate")
    st.info("Upload CSV with topics")
    
    sample_df = pd.DataFrame({
        'topic': ['AI', 'Marketing'],
        'category': ['Tech', 'Business'],
        'slide_count': [6, 8],
        'audience': ['Students', 'Corporate']
    })
    st.dataframe(sample_df)
    
    csv_sample = sample_df.to_csv(index=False)
    st.download_button("📥 Sample CSV", csv_sample, "sample.csv", "text/csv")

with tab3:
    st.subheader("⚙️ Templates")
    st.info("Save/Load settings")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>🎯 <strong>AI PowerPoint Generator Pro - Google API Edition</strong></p>
    <p>✨ With Google Custom Search | Multi-Language | AI Coach!</p>
</div>
""", unsafe_allow_html=True)
