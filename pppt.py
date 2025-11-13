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
    
    # Pexels API (Optional but recommended)
    pexels_api_key = st.text_input(
        "Pexels API Key (Optional)", 
        type="password",
        help="FREE! Get better topic-relevant images. Sign up at: https://www.pexels.com/api/"
    )
    
    if pexels_api_key:
        st.success("✅ Pexels connected - will search for topic-specific images!")
    
    stability_api_key = st.text_input("Stability AI API Key (Optional)", type="password", 
                                      help="Optional: For AI-generated images")
    
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
    2. **Optional:** Add Pexels key for better images
    3. Enter your presentation topic
    4. Click Generate!
    5. Edit slides if needed
    6. Download in your format!
    """)
    st.markdown("---")
    st.markdown("### 🔗 Get API Keys")
    st.markdown("🆓 [Pexels API (FREE)](https://www.pexels.com/api/)")
    st.markdown("[OpenRouter API](https://openrouter.ai/keys)")
    st.markdown("[Stability AI](https://platform.stability.ai)")

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

def get_unsplash_image(query, width=800, height=600):
    """Get image from Unsplash"""
    try:
        clean_query = query.strip().replace(' ', ',')
        url = f"https://source.unsplash.com/{width}x{height}/?{clean_query}"
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(url, timeout=15, allow_redirects=True, headers=headers)
        
        if response.status_code == 200 and len(response.content) > 5000:
            # Validate image
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
    """Get image from Pexels API"""
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

def get_topic_relevant_image(main_topic, slide_title, image_prompt, pexels_key=None):
    """Get highly relevant image for the topic"""
    
    st.write(f"   🎯 Topic: {main_topic}")
    st.write(f"   📄 Slide: {slide_title}")
    
    # Generate search terms
    search_terms = generate_topic_search_terms(main_topic, slide_title, image_prompt)
    st.write(f"   🔍 Will try {len(search_terms)} search variations")
    
    # Try each search term
    for i, term in enumerate(search_terms, 1):
        st.write(f"      → Search {i}: '{term}'")
        
        # Try Pexels first if available
        if pexels_key:
            image_data = get_pexels_image(term, pexels_key)
            if image_data:
                st.write(f"      ✅ Found on Pexels!")
                return image_data
        
        # Try Unsplash
        image_data = get_unsplash_image(term)
        if image_data:
            st.write(f"      ✅ Found on Unsplash!")
            return image_data
    
    # Fallback to generic topic
    st.write(f"   🆘 Trying generic fallback...")
    fallback = main_topic.split()[0] if main_topic else "business"
    image_data = get_unsplash_image(fallback)
    if image_data:
        st.write(f"      ✅ Got fallback image")
        return image_data
    
    return None

def generate_image_stability(api_key, prompt, image_style):
    """Generate AI image using Stability AI"""
    try:
        style_prompts = {
            "Professional": "professional photography, corporate, clean",
            "Minimalist": "minimal, simple, clean lines, white background",
            "Colorful": "vibrant colors, dynamic, energetic",
            "Corporate": "business, formal, blue tones, professional",
            "Creative": "artistic, creative, unique perspective",
            "Infographic": "flat design, infographic style, icons"
        }
        
        enhanced_prompt = f"{prompt}, {style_prompts.get(image_style, style_prompts['Professional'])}"
        
        url = "https://api.stability.ai/v2beta/stable-image/generate/core"
        
        response = requests.post(
            url,
            headers={
                "authorization": f"Bearer {api_key.strip()}",
                "accept": "image/*"
            },
            files={"none": ''},
            data={
                "prompt": enhanced_prompt,
                "output_format": "png",
            },
        )
        
        if response.status_code == 200:
            return response.content
        return None
    except:
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
                error_data = response.json()
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
            raise  # Re-raise for retry logic
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
                    wait_time = (attempt + 1) * 5  # 5, 10, 15 seconds
                    st.warning(f"⏳ Rate limit hit. Retrying in {wait_time} seconds... (Attempt {attempt + 2}/{max_retries})")
                    time.sleep(wait_time)
                else:
                    st.error("❌ Rate limit persists after retries. Please:\n1. Wait 1-2 minutes\n2. Switch to different free model\n3. Use Claude model (paid)")
                    return None
            else:
                return None
    return None

# ============ ANALYSIS FUNCTIONS ============

def analyze_presentation(slides_content):
    """Analyze presentation quality with AI Coach"""
    issues = []
    suggestions = []
    score = 100
    
    for i, slide in enumerate(slides_content, 1):
        # Too many bullets
        bullet_count = len(slide.get('bullets', []))
        if bullet_count > 5:
            issues.append(f"Slide {i}: Too many bullets ({bullet_count})")
            suggestions.append(f"Slide {i}: Reduce to 3-5 key points for better retention")
            score -= 5
        
        # Title too long
        title_len = len(slide['title'])
        if title_len > 60:
            issues.append(f"Slide {i}: Title too long ({title_len} chars)")
            suggestions.append(f"Slide {i}: Shorten title to <60 characters")
            score -= 3
        
        # Check for text-heavy slides
        total_text = sum(len(b) for b in slide.get('bullets', []))
        if total_text > 500:
            issues.append(f"Slide {i}: Too much text ({total_text} chars)")
            suggestions.append(f"Slide {i}: Use more visuals, less text. Aim for <400 chars")
            score -= 5
        
        # Check for speaker notes
        if not slide.get('speaker_notes') or len(slide.get('speaker_notes', '')) < 20:
            suggestions.append(f"Slide {i}: Add detailed speaker notes for better delivery")
            score -= 2
    
    score = max(0, score)
    return issues, suggestions, score

def generate_slide_preview(slide_data, theme_colors):
    """Generate preview image of slide"""
    fig, ax = plt.subplots(figsize=(10, 7.5), facecolor='#' + ''.join([format(c, '02x') for c in theme_colors['bg']]))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 7.5)
    ax.axis('off')
    
    # Title
    title_color = '#' + ''.join([format(c, '02x') for c in theme_colors['accent']])
    ax.text(5, 6.5, slide_data['title'][:50], ha='center', fontsize=18, weight='bold', color=title_color)
    
    # Bullets
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
    """Export presentation to PDF format"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title = Paragraph(f"<b>{topic}</b>", styles['Title'])
    story.append(title)
    story.append(Spacer(1, 12))
    
    # Content
    for i, slide in enumerate(slides_content, 1):
        # Slide title
        slide_title = Paragraph(f"<b>Slide {i}: {slide['title']}</b>", styles['Heading2'])
        story.append(slide_title)
        story.append(Spacer(1, 6))
        
        # Bullets
        for bullet in slide.get('bullets', []):
            bullet_text = Paragraph(f"• {bullet}", styles['BodyText'])
            story.append(bullet_text)
        
        # Speaker notes
        if slide.get('speaker_notes'):
            notes = Paragraph(f"<i>Notes: {slide['speaker_notes']}</i>", styles['Italic'])
            story.append(Spacer(1, 6))
            story.append(notes)
        
        story.append(Spacer(1, 20))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

def export_to_google_slides_json(slides_content, topic, theme):
    """Export to Google Slides compatible JSON format"""
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

def create_powerpoint(slides_content, theme, image_mode, stability_key, pexels_key, category, audience, topic, image_position, image_style, logo_data):
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
    
    # Image position settings
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
        
        # Background
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = color_scheme["bg"]
        
        # Add logo if provided
        if logo_data:
            try:
                logo_stream = io.BytesIO(logo_data)
                logo_left = Inches(9)
                logo_top = Inches(0.2)
                logo_width = Inches(0.8)
                slide.shapes.add_picture(logo_stream, logo_left, logo_top, width=logo_width)
            except:
                pass
        
        # Title
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8.5), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = slide_data["title"]
        title_frame.paragraphs[0].font.size = Pt(36 if idx == 0 else 28)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].font.color.rgb = color_scheme["accent"]
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER if idx == 0 else PP_ALIGN.LEFT
        
        # Bullets
        if idx > 0 and slide_data.get("bullets"):
            bullet_width = Inches(5.5) if image_mode != "None" and image_position == "Right Side" else Inches(9)
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
        
        # Subtitle on first slide
        if idx == 0:
            subtitle_box = slide.shapes.add_textbox(Inches(0.5), Inches(3), Inches(9), Inches(1))
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = f"{category} Presentation | {audience}"
            subtitle_frame.paragraphs[0].font.size = Pt(20)
            subtitle_frame.paragraphs[0].font.color.rgb = color_scheme["text"]
            subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Speaker notes
        if slide_data.get("speaker_notes"):
            notes_slide = slide.notes_slide
            notes_slide.notes_text_frame.text = slide_data["speaker_notes"]
        
        # Add images to content slides
        if idx > 0 and image_mode != "None":
            with st.expander(f"🖼️ Slide {idx + 1}: {slide_data['title']}", expanded=False):
                image_prompt = slide_data.get("image_prompt", "")
                image_data = None
                
                if image_mode == "AI Generated (Paid)" and stability_key:
                    st.write("   🤖 Generating AI image...")
                    image_data = generate_image_stability(stability_key, image_prompt, image_style)
                
                if not image_data and image_mode != "None":
                    st.write("   🔍 Searching for topic-relevant image...")
                    image_data = get_topic_relevant_image(
                        main_topic=topic,
                        slide_title=slide_data["title"],
                        image_prompt=image_prompt,
                        pexels_key=pexels_key
                    )
                
                if image_data:
                    try:
                        image_stream = io.BytesIO(image_data)
                        pic = slide.shapes.add_picture(
                            image_stream, 
                            img_pos["left"], 
                            img_pos["top"], 
                            width=img_pos["width"]
                        )
                        st.success(f"   ✅ Image added successfully!")
                    except Exception as e:
                        st.error(f"   ❌ Failed to add image: {str(e)}")
                else:
                    st.warning(f"   ⚠️ No image found")
            
            time.sleep(0.5)
    
    progress_bar.progress(1.0)
    status_text.text("✅ Presentation created!")
    
    return prs

# ============ TEMPLATE FUNCTIONS ============

def save_template(category, slide_count, tone, audience, theme, image_mode, language):
    """Save current settings as template"""
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
    """Load template from uploaded file"""
    try:
        template = json.loads(template_file.read())
        return template
    except:
        return None

# ============ MAIN UI ============

# Tab navigation
tab1, tab2, tab3 = st.tabs(["📝 Create Presentation", "📊 Bulk Generate", "⚙️ Templates"])

with tab1:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📝 Your Topic")
        topic = st.text_input("Enter Topic *", placeholder="e.g., Space Exploration, Digital Marketing, Climate Change...")
        st.caption("💡 Be specific! The more detailed your topic, the better the images will match.")
        
        category = st.selectbox("Category *", ["Business", "Pitch", "Marketing", "Technical", "Academic", "Training", "Sales"])
        
        col1_1, col1_2 = st.columns(2)
        with col1_1:
            slide_count = st.number_input("Number of Slides *", min_value=3, max_value=20, value=6)
        with col1_2:
            language = st.selectbox("Language 🌍", ["English", "Hindi (हिंदी)", "Spanish", "French", "German", "Chinese", "Japanese"])
        
        tone = st.selectbox("Tone *", ["Formal", "Neutral", "Inspirational", "Educational", "Persuasive"])

    with col2:
        st.subheader("🎨 Style & Images")
        audience = st.selectbox("Audience *", ["Investors", "Students", "Corporate", "Clients", "Managers", "General Public"])
        theme = st.selectbox("Theme *", ["Corporate Blue", "Gradient Modern", "Minimal Dark", "Pastel Soft", "Professional Green", "Elegant Purple"])
        
        image_mode = st.selectbox(
            "Image Mode *",
            ["Free Images (Auto)", "AI Generated (Paid)", "None"],
            help="Auto: Uses Unsplash/Pexels for topic-relevant free images"
        )
        
        if image_mode != "None":
            col2_1, col2_2 = st.columns(2)
            with col2_1:
                image_position = st.selectbox("Image Position", ["Right Side", "Left Side", "Top Right Corner", "Bottom", "Center"])
            with col2_2:
                image_style = st.selectbox("Image Style", ["Professional", "Minimalist", "Colorful", "Corporate", "Creative", "Infographic"])

    st.subheader("➕ Additional Points (Optional)")
    key_points = st.text_area("Key points to cover", placeholder="- Point 1\n- Point 2", height=80)

    # Export format
    export_format = st.selectbox("Export Format", ["PowerPoint (.pptx)", "PowerPoint + PDF", "Google Slides (JSON)"])

    st.markdown("---")
    
    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        generate_button = st.button("🚀 Generate PowerPoint", use_container_width=True)
    with col_btn2:
        save_template_btn = st.button("💾 Save as Template", use_container_width=True)

    if save_template_btn:
        template_json = save_template(category, slide_count, tone, audience, theme, image_mode, language)
        st.download_button(
            label="📥 Download Template",
            data=template_json,
            file_name="presentation_template.json",
            mime="application/json"
        )
        st.success("✅ Template ready to download!")

    if generate_button:
        if not claude_api_key:
            st.error("⚠️ Enter OpenRouter API key in sidebar")
        elif not topic:
            st.error("⚠️ Enter a topic")
        else:
            with st.spinner("🤖 Generating your presentation..."):
                slides_content = generate_content_with_retry(
                    claude_api_key, topic, category, slide_count, 
                    tone, audience, key_points, model_choice, language
                )
                
                if slides_content:
                    st.session_state.slides_content = slides_content
                    st.session_state.generation_count += 1
                    st.session_state.total_slides += len(slides_content)
                    
                    st.success("✅ Content generated!")
                    
                    # Presentation Statistics
                    st.markdown("---")
                    st.subheader("📊 Presentation Statistics")
                    col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
                    
                    with col_stat1:
                        st.metric("Total Slides", len(slides_content))
                    
                    with col_stat2:
                        total_words = sum(len(' '.join(s.get('bullets', []))).split() for s in slides_content)
                        st.metric("Total Words", total_words)
                    
                    with col_stat3:
                        avg_bullets = sum(len(s.get('bullets', [])) for s in slides_content) / len(slides_content)
                        st.metric("Avg Bullets/Slide", f"{avg_bullets:.1f}")
                    
                    with col_stat4:
                        est_time = len(slides_content) * 2
                        st.metric("Est. Time", f"{est_time} min")
                    
                    # AI Presentation Coach
                    st.markdown("---")
                    with st.expander("🎓 AI Presentation Coach", expanded=True):
                        issues, suggestions, score = analyze_presentation(slides_content)
                        
                        col_score1, col_score2 = st.columns([1, 3])
                        with col_score1:
                            score_color = "🟢" if score >= 80 else "🟡" if score >= 60 else "🔴"
                            st.markdown(f"### {score_color} Score: {score}/100")
                        
                        with col_score2:
                            if score >= 80:
                                st.success("Excellent! Your presentation follows best practices.")
                            elif score >= 60:
                                st.warning("Good, but there's room for improvement.")
                            else:
                                st.error("Consider revising based on suggestions below.")
                        
                        if issues:
                            st.markdown("#### ⚠️ Issues Found:")
                            for issue in issues:
                                st.write(f"- {issue}")
                        
                        if suggestions:
                            st.markdown("#### 💡 Suggestions:")
                            for suggestion in suggestions:
                                st.write(f"- {suggestion}")
                    
                    # Edit Slides Section
                    st.markdown("---")
                    st.subheader("✏️ Review & Edit Slides")
                    enable_edit = st.checkbox("Enable slide editing")
                    
                    if enable_edit:
                        edited_slides = []
                        for i, slide in enumerate(slides_content):
                            with st.expander(f"Slide {i+1}: {slide['title']}", expanded=False):
                                new_title = st.text_input(f"Title", slide['title'], key=f"edit_title_{i}")
                                new_bullets = st.text_area(
                                    f"Content (one per line)", 
                                    "\n".join(slide.get('bullets', [])), 
                                    height=150,
                                    key=f"edit_bullets_{i}"
                                )
                                new_notes = st.text_area(
                                    f"Speaker Notes",
                                    slide.get('speaker_notes', ''),
                                    height=100,
                                    key=f"edit_notes_{i}"
                                )
                                edited_slides.append({
                                    "title": new_title,
                                    "bullets": [b.strip() for b in new_bullets.split("\n") if b.strip()],
                                    "image_prompt": slide.get('image_prompt', ''),
                                    "speaker_notes": new_notes
                                })
                        
                        if st.button("✅ Apply Edits"):
                            st.session_state.slides_content = edited_slides
                            slides_content = edited_slides
                            st.success("✅ Edits applied!")
                            st.rerun()
                    
                    # Preview Slides
                    st.markdown("---")
                    st.subheader("👀 Slide Previews")
                    
                    themes_for_preview = {
                        "Corporate Blue": {"bg": (240, 248, 255), "accent": (31, 119, 180), "text": (0, 0, 0)},
                        "Gradient Modern": {"bg": (240, 242, 246), "accent": (138, 43, 226), "text": (0, 0, 0)},
                        "Minimal Dark": {"bg": (30, 30, 30), "accent": (255, 215, 0), "text": (255, 255, 255)},
                        "Pastel Soft": {"bg": (255, 250, 240), "accent": (255, 182, 193), "text": (60, 60, 60)},
                        "Professional Green": {"bg": (245, 255, 250), "accent": (34, 139, 34), "text": (0, 0, 0)},
                        "Elegant Purple": {"bg": (250, 245, 255), "accent": (128, 0, 128), "text": (0, 0, 0)}
                    }
                    
                    theme_colors = themes_for_preview.get(theme, themes_for_preview["Corporate Blue"])
                    
                    preview_cols = st.columns(3)
                    for i, slide in enumerate(slides_content[:6]):
                        with preview_cols[i % 3]:
                            preview_img = generate_slide_preview(slide, theme_colors)
                            st.image(preview_img, caption=f"Slide {i+1}", use_container_width=True)
                    
                    if len(slides_content) > 6:
                        st.info(f"Showing first 6 of {len(slides_content)} slides")
                    
                    # Content Preview
                    st.markdown("---")
                    with st.expander("📄 Full Content Preview"):
                        for i, slide in enumerate(slides_content):
                            st.markdown(f"### Slide {i+1}: {slide['title']}")
                            if slide.get('bullets'):
                                for bullet in slide['bullets']:
                                    st.write(f"- {bullet}")
                            if slide.get('speaker_notes'):
                                st.caption(f"📝 Notes: {slide['speaker_notes']}")
                            if slide.get('image_prompt'):
                                st.caption(f"🖼️ Image: {slide['image_prompt']}")
                            st.markdown("---")
                    
                    # Generate Final Presentation
                    st.markdown("---")
                    if st.button("🎨 Create Final Presentation", use_container_width=True):
                        with st.spinner("Creating your presentation with images..."):
                            prs = create_powerpoint(
                                slides_content, theme, image_mode, 
                                stability_api_key, pexels_api_key,
                                category, audience, topic, 
                                image_position if image_mode != "None" else "Right Side",
                                image_style if image_mode == "AI Generated (Paid)" else "Professional",
                                logo_data
                            )
                            
                            st.success("🎉 PowerPoint ready!")
                            
                            # Export based on format
                            if export_format == "PowerPoint (.pptx)":
                                pptx_io = io.BytesIO()
                                prs.save(pptx_io)
                                pptx_io.seek(0)
                                
                                st.download_button(
                                    label="📥 Download PowerPoint",
                                    data=pptx_io,
                                    file_name=f"{topic.replace(' ', '_')}.pptx",
                                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                    use_container_width=True
                                )
                            
                            elif export_format == "PowerPoint + PDF":
                                col_dl1, col_dl2 = st.columns(2)
                                
                                with col_dl1:
                                    pptx_io = io.BytesIO()
                                    prs.save(pptx_io)
                                    pptx_io.seek(0)
                                    
                                    st.download_button(
                                        label="📥 Download PowerPoint",
                                        data=pptx_io,
                                        file_name=f"{topic.replace(' ', '_')}.pptx",
                                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                        use_container_width=True
                                    )
                                
                                with col_dl2:
                                    pdf_buffer = export_to_pdf(slides_content, topic)
                                    st.download_button(
                                        label="📄 Download PDF",
                                        data=pdf_buffer,
                                        file_name=f"{topic.replace(' ', '_')}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
                            
                            elif export_format == "Google Slides (JSON)":
                                google_json = export_to_google_slides_json(slides_content, topic, theme)
                                st.download_button(
                                    label="📥 Download Google Slides JSON",
                                    data=google_json,
                                    file_name=f"{topic.replace(' ', '_')}_google_slides.json",
                                    mime="application/json",
                                    use_container_width=True
                                )
                                st.info("💡 Import this JSON into Google Slides using Apps Script")

with tab2:
    st.subheader("📊 Bulk Presentation Generation")
    st.info("Upload a CSV file with multiple topics to generate presentations in batch")
    
    # Sample CSV template
    st.markdown("### 📋 CSV Format Required:")
    sample_df = pd.DataFrame({
        'topic': ['Digital Marketing', 'Climate Change'],
        'category': ['Business', 'Academic'],
        'slide_count': [6, 8],
        'audience': ['Corporate', 'Students']
    })
    st.dataframe(sample_df)
    
    # Download sample
    csv_sample = sample_df.to_csv(index=False)
    st.download_button("📥 Download Sample CSV", csv_sample, "sample_bulk.csv", "text/csv")
    
    st.markdown("---")
    
    # Upload CSV
    bulk_file = st.file_uploader("📂 Upload CSV for Bulk Generation", type="csv")
    
    if bulk_file:
        try:
            df = pd.read_csv(bulk_file)
            st.success(f"✅ Loaded {len(df)} topics")
            st.dataframe(df)
            
            if st.button("🚀 Generate All Presentations"):
                if not claude_api_key:
                    st.error("⚠️ Enter OpenRouter API key in sidebar")
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        progress = st.progress(0)
                        
                        for idx, row in df.iterrows():
                            st.write(f"Generating: {row['topic']}...")
                            
                            slides = generate_content_with_retry(
                                claude_api_key, 
                                row['topic'], 
                                row.get('category', 'Business'),
                                int(row.get('slide_count', 6)),
                                tone,
                                row.get('audience', 'Corporate'),
                                "",
                                model_choice,
                                language
                            )
                            
                            if slides:
                                prs = create_powerpoint(
                                    slides, theme, "None", 
                                    None, None,
                                    row.get('category', 'Business'),
                                    row.get('audience', 'Corporate'),
                                    row['topic'],
                                    "Right Side",
                                    "Professional",
                                    logo_data
                                )
                                
                                pptx_io = io.BytesIO()
                                prs.save(pptx_io)
                                zip_file.writestr(f"{row['topic'].replace(' ', '_')}.pptx", pptx_io.getvalue())
                            
                            progress.progress((idx + 1) / len(df))
                        
                        st.success("✅ All presentations generated!")
                        
                        zip_buffer.seek(0)
                        st.download_button(
                            "📦 Download All (ZIP)",
                            zip_buffer.getvalue(),
                            "presentations_bulk.zip",
                            "application/zip",
                            use_container_width=True
                        )
        except Exception as e:
            st.error(f"Error processing CSV: {str(e)}")

with tab3:
    st.subheader("⚙️ Template Management")
    
    col_temp1, col_temp2 = st.columns(2)
    
    with col_temp1:
        st.markdown("### 💾 Save Template")
        st.info("Save your current settings as a reusable template")
        
        if st.button("Save Current Settings", use_container_width=True):
            template_json = save_template(
                category if 'category' in locals() else "Business",
                slide_count if 'slide_count' in locals() else 6,
                tone if 'tone' in locals() else "Formal",
                audience if 'audience' in locals() else "Corporate",
                theme if 'theme' in locals() else "Corporate Blue",
                image_mode if 'image_mode' in locals() else "Free Images (Auto)",
                language if 'language' in locals() else "English"
            )
            st.download_button(
                "📥 Download Template",
                template_json,
                "my_template.json",
                "application/json",
                use_container_width=True
            )
    
    with col_temp2:
        st.markdown("### 📂 Load Template")
        st.info("Upload a saved template to quickly apply settings")
        
        template_file = st.file_uploader("Upload Template JSON", type="json")
        
        if template_file:
            loaded_template = load_template(template_file)
            if loaded_template:
                st.success("✅ Template loaded!")
                st.json(loaded_template)
                st.info("Go to 'Create Presentation' tab and manually apply these settings")
            else:
                st.error("❌ Invalid template file")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>🎯 <strong>AI PowerPoint Generator Pro v2.0</strong></p>
    <p>✨ With Multi-Language | AI Coach | Bulk Generation | Templates & More!</p>
    <p>🆓 <strong>Get Pexels API</strong> (free) for best results: <a href="https://www.pexels.com/api/">pexels.com/api</a></p>
</div>
""", unsafe_allow_html=True)
