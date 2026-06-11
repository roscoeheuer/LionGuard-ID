import os
import re
from datetime import date
import pandas as pd
import streamlit as st
from PIL import Image

try:
    from supabase import create_client
except Exception:
    create_client = None

APP_NAME = "LionGuard ID"
KNOWN_LIONS_CSV = "known_lions.csv"
SUBMISSIONS_CSV = "lion_id_submissions.csv"
SIGHTINGS_FILE = "lion_sightings.xlsx"
UPLOAD_FOLDER = "lion_uploads"
ASSETS_FOLDER = "assets"
KNOWN_LION_IMAGES_FOLDER = "known_lion_images"
SUPABASE_BUCKET = "lion-images"

HERO_IMAGE = os.path.join(ASSETS_FOLDER, "hero_lion.png")
LION_CUBS_IMAGE = os.path.join(ASSETS_FOLDER, "lion_cubs.png")
LIONESS_IMAGE = os.path.join(ASSETS_FOLDER, "lioness.png")
LION_FAMILY_IMAGE = os.path.join(ASSETS_FOLDER, "lion_family.png")
LION_MALE_IMAGE = os.path.join(ASSETS_FOLDER, "lion_male.png")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(KNOWN_LION_IMAGES_FOLDER, exist_ok=True)

st.set_page_config(page_title=APP_NAME, page_icon="🦁", layout="wide")

st.markdown("""
<style>
:root {
    --maroon:#6b1e24;
    --maroon-dk:#4d1519;
    --gold:#c8973f;
    --white:#ffffff;
    --off-white:#f8f5f0;
    --ink:#2c2c2c;
    --muted:#6e6e6e;
    --border:#e0d8ce;
}
.stApp { background:white; color:var(--ink); }
.block-container { padding:0 2rem 4rem !important; max-width:1240px !important; }
h1,h2,h3,h4 { color:var(--maroon) !important; font-family:Georgia,serif !important; }
section[data-testid="stSidebar"] { background:var(--maroon-dk); border-right:4px solid var(--gold); }
section[data-testid="stSidebar"] * { color:white !important; }
.topbar { background:var(--maroon); color:white; padding:14px 32px; margin:-2rem -2rem 0; border-bottom:4px solid var(--gold); display:flex; gap:14px; align-items:center; }
.topbar-logo { font-size:22px; font-weight:900; color:white !important; font-family:Georgia,serif; }
.topbar-tagline { font-size:12px; color:#e8b96a !important; letter-spacing:.1em; text-transform:uppercase; }
.page-band { background:var(--off-white); border-bottom:3px solid var(--gold); padding:32px; margin:0 -2rem 36px; }
.page-band h1 { margin:0 0 6px !important; font-size:36px !important; }
.page-band p { margin:0; color:var(--muted) !important; }
.hero { background:var(--maroon); padding:56px 48px; margin:0 -2rem 0; border-bottom:5px solid var(--gold); }
.hero-title { color:white !important; font-family:Georgia,serif; font-size:56px; font-weight:900; }
.hero-sub { color:rgba(255,255,255,.85) !important; font-size:18px; line-height:1.6; max-width:600px; }
.hero-btn { background:var(--gold); color:var(--maroon-dk) !important; padding:12px 28px; display:inline-block; font-weight:800; margin-top:20px; }
.stat-card { background:var(--maroon); padding:30px; text-align:center; border-bottom:4px solid var(--gold); }
.stat-num { color:white !important; font-size:48px; font-family:Georgia,serif; font-weight:900; }
.stat-lbl { color:#e8b96a !important; text-transform:uppercase; font-size:12px; letter-spacing:.12em; }
.lion-card,.card { background:white; border:1px solid var(--border); border-left:5px solid var(--maroon); padding:24px; margin-bottom:20px; box-shadow:0 2px 16px rgba(107,30,36,.10); }
.profile-card { background:white; border:1px solid var(--border); border-top:5px solid var(--maroon); padding:28px; margin-bottom:24px; box-shadow:0 2px 16px rgba(107,30,36,.10); }
.lion-name { font-family:Georgia,serif; font-size:24px; font-weight:800; color:var(--maroon) !important; }
.lion-id-tag { font-size:13px; color:var(--muted) !important; margin-left:8px; }
.badge { display:inline-block; padding:4px 12px; margin:5px 5px 5px 0; background:var(--off-white); border:1px solid var(--border); color:var(--maroon) !important; border-radius:99px; font-size:11px; font-weight:800; text-transform:uppercase; }
.fld { margin-bottom:14px; }
.fld-key { font-size:10px; font-weight:800; letter-spacing:.14em; text-transform:uppercase; color:var(--muted) !important; }
.fld-val { font-size:15px; color:var(--ink) !important; line-height:1.55; }
.img-placeholder { background:var(--off-white); border:1px solid var(--border); height:200px; display:flex; align-items:center; justify-content:center; font-size:36px; }
.stButton > button { background:var(--maroon) !important; color:white !important; border:none !important; border-radius:2px !important; font-weight:800 !important; text-transform:uppercase !important; letter-spacing:.08em !important; }
.stButton > button:hover { background:#8c2830 !important; }
.site-footer { background:var(--maroon-dk); color:white; text-align:center; padding:20px; margin:48px -2rem -4rem; border-top:4px solid var(--gold); }
.site-footer span { color:#e8b96a !important; }
</style>
""", unsafe_allow_html=True)

KNOWN_COLUMNS = [
    "lion_id", "lion_name", "sex", "age",
    "general_description", "whisker_pattern_description",
    "whisker_pattern_image", "reference_image",
    "all_reference_images_folder", "gallery_image_urls"
]

SUBMISSION_COLUMNS = [
    "submission_id", "image_path", "date_seen", "location",
    "sex", "age_class", "side_visible", "mane_description",
    "scars_or_injuries", "ear_notches", "whisker_spot_notes",
    "tail_or_body_marks", "behavior_context", "suspected_lion",
    "confidence", "reviewer_notes", "expert_final_id", "expert_status"
]

SIGHTING_COLUMNS = [
    "Name of Lion", "Date Sighted", "Conservancy", "Location",
    "Recorder", "Status", "Age", "Sex", "Notes"
]

IMAGE_EXTS = (".jpg", ".jpeg", ".png", ".JPG", ".JPEG", ".PNG", ".heic", ".HEIC")


def ensure_csv(path, cols):
    if not os.path.exists(path):
        pd.DataFrame(columns=cols).to_csv(path, index=False)


def load_known_lions():
    ensure_csv(KNOWN_LIONS_CSV, KNOWN_COLUMNS)
    df = pd.read_csv(KNOWN_LIONS_CSV)

    for col in KNOWN_COLUMNS:
        if col not in df.columns:
            if col == "age":
                df[col] = "Unknown"
            else:
                df[col] = ""

    df["age"] = df["age"].fillna("Unknown")
    df.loc[df["age"].astype(str).str.strip() == "", "age"] = "Unknown"

    return df[KNOWN_COLUMNS]


def load_submissions():
    ensure_csv(SUBMISSIONS_CSV, SUBMISSION_COLUMNS)
    df = pd.read_csv(SUBMISSIONS_CSV)
    for col in SUBMISSION_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[SUBMISSION_COLUMNS]


def load_sightings():
    if not os.path.exists(SIGHTINGS_FILE):
        return pd.DataFrame(columns=SIGHTING_COLUMNS)

    try:
        df = pd.read_excel(SIGHTINGS_FILE)
    except Exception as e:
        st.warning(f"Could not read {SIGHTINGS_FILE}: {e}")
        return pd.DataFrame(columns=SIGHTING_COLUMNS)

    for col in SIGHTING_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    df["Name of Lion"] = df["Name of Lion"].astype(str).str.strip()
    df["Date Sighted"] = pd.to_datetime(df["Date Sighted"], errors="coerce")
    return df[SIGHTING_COLUMNS]


def save_submissions(df):
    df.to_csv(SUBMISSIONS_CSV, index=False)


def save_known_lions(df):
    df.to_csv(KNOWN_LIONS_CSV, index=False)


def clean_value(value):
    if pd.isna(value):
        return "—"
    value = str(value).strip()
    if value == "" or value.upper() in ["N/A", "NA", "NONE", "UNKNOWN", "NAN"]:
        return "—"
    return value


def is_valid_path(value):
    if pd.isna(value):
        return False
    value = str(value).strip()
    return value not in ["", "N/A", "NA", "None", "Unknown", "nan"] and os.path.exists(value)


def show_image(path, caption=None):
    if pd.isna(path):
        st.markdown("<div class='img-placeholder'>🦁</div>", unsafe_allow_html=True)
        return

    path = str(path).strip()

    if path.startswith("http://") or path.startswith("https://"):
        st.image(path, caption=caption, use_container_width=True)
    elif is_valid_path(path):
        st.image(path, caption=caption, use_container_width=True)
    else:
        st.markdown("<div class='img-placeholder'>🦁</div>", unsafe_allow_html=True)


def field(key, val):
    st.markdown(
        f"<div class='fld'><div class='fld-key'>{key}</div><div class='fld-val'>{clean_value(val)}</div></div>",
        unsafe_allow_html=True
    )


def page_band(title, subtitle=""):
    st.markdown(f"<div class='page-band'><h1>{title}</h1><p>{subtitle}</p></div>", unsafe_allow_html=True)


def get_lion_images(folder):
    if not is_valid_path(folder):
        return []
    images = []
    for file in sorted(os.listdir(folder)):
        if file.endswith(IMAGE_EXTS):
            images.append(os.path.join(folder, file))
    return images


def set_selected_lion(lion_name):
    st.session_state["selected_lion_name"] = lion_name


def safe_folder_name(name):
    name = str(name).strip()
    name = re.sub(r"[^A-Za-z0-9 _-]", "", name)
    name = name.replace(" ", "_")
    return name or "New_Lion"
def get_supabase_client():
    if create_client is None:
        return None

    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except Exception:
        return None


def upload_photo_to_supabase(file, lion_name):
    supabase = get_supabase_client()

    if supabase is None:
        return None

    safe_lion = safe_folder_name(lion_name)
    original_name = os.path.splitext(file.name)[0]
    ext = os.path.splitext(file.name)[1].lower()

    if ext not in [".jpg", ".jpeg", ".png"]:
        ext = ".jpg"

    safe_file = safe_folder_name(original_name)
    storage_path = f"{safe_lion}/{safe_file}{ext}"

    try:
        supabase.storage.from_(SUPABASE_BUCKET).upload(
            path=storage_path,
            file=file.getvalue(),
            file_options={
                "content-type": file.type or "image/jpeg",
                "upsert": "true"
            }
        )

        return supabase.storage.from_(SUPABASE_BUCKET).get_public_url(storage_path)

    except Exception as e:
        st.error(f"Supabase upload failed: {e}")
        return None


def save_uploaded_photo(file, lion_name, fallback_folder=None):
    public_url = upload_photo_to_supabase(file, lion_name)

    if public_url:
        return public_url

    if fallback_folder:
        os.makedirs(fallback_folder, exist_ok=True)
        save_path = os.path.join(fallback_folder, file.name)

        with open(save_path, "wb") as f:
            f.write(file.getbuffer())

        return save_path

    return ""

def split_gallery_urls(value):
    if pd.isna(value):
        return []

    value = str(value).strip()
    if not value:
        return []

    return [url.strip() for url in value.split("|") if url.strip()]


def join_gallery_urls(urls):
    clean_urls = []
    for url in urls:
        url = str(url).strip()
        if url and url not in clean_urls:
            clean_urls.append(url)
    return "|".join(clean_urls)


def save_uploaded_photo(file, lion_name, fallback_folder=None):
    if not fallback_folder:
        fallback_folder = os.path.join(KNOWN_LION_IMAGES_FOLDER, safe_folder_name(lion_name))

    os.makedirs(fallback_folder, exist_ok=True)

    safe_file_name = safe_folder_name(os.path.splitext(file.name)[0])
    ext = os.path.splitext(file.name)[1].lower()

    if ext not in [".jpg", ".jpeg", ".png"]:
        ext = ".jpg"

    save_path = os.path.join(fallback_folder, f"{safe_file_name}{ext}")

    with open(save_path, "wb") as f:
        f.write(file.getbuffer())

    return save_path


def get_last_sighting(sightings_df, lion_name):
    if sightings_df.empty or not lion_name:
        return None, pd.DataFrame()

    lion_sightings = sightings_df[
        sightings_df["Name of Lion"].astype(str).str.lower().str.strip() == str(lion_name).lower().strip()
    ].copy()

    lion_sightings = lion_sightings.dropna(subset=["Date Sighted"])

    if lion_sightings.empty:
        return None, lion_sightings

    lion_sightings = lion_sightings.sort_values("Date Sighted", ascending=False)
    return lion_sightings.iloc[0], lion_sightings


def show_last_sighting(sightings_df, lion_name):
    st.subheader("Last Known Sighting")

    if sightings_df.empty:
        st.info("No lion_sightings.xlsx file found yet. Put it in the LionIDApp folder to display last known sightings.")
        return

    last_sighting, lion_sightings = get_last_sighting(sightings_df, lion_name)

    if last_sighting is None:
        st.info(f"No sighting records found for {lion_name}.")
        return

    st.markdown("<div class='profile-card'>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)

    with c1:
        field("Date Sighted", last_sighting["Date Sighted"].strftime("%B %d, %Y"))
        field("Conservancy", last_sighting.get("Conservancy", ""))
        field("Location", last_sighting.get("Location", ""))

    with c2:
        field("Recorder", last_sighting.get("Recorder", ""))
        field("Status", last_sighting.get("Status", ""))
        field("Age", last_sighting.get("Age", ""))

    with c3:
        field("Sex", last_sighting.get("Sex", ""))
        field("Notes", last_sighting.get("Notes", ""))

    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander("View full sighting history for this lion"):
        st.dataframe(lion_sightings, use_container_width=True, hide_index=True)


def score_lion_match(row, sex_input, description_input, whisker_input):
    score = 0
    reasons = []

    lion_sex = clean_value(row.get("sex", "")).lower()
    user_sex = sex_input.lower()

    if user_sex != "unknown" and lion_sex == user_sex:
        score += 30
        reasons.append("Sex matches")

    combined_lion_text = " ".join([
        str(row.get("general_description", "")),
        str(row.get("whisker_pattern_description", "")),
    ]).lower()

    user_text = " ".join([description_input, whisker_input]).lower()

    user_words = [
        w.strip(".,;:!?()[]")
        for w in user_text.split()
        if len(w.strip(".,;:!?()[]")) >= 4
    ]

    matched_words = []
    for word in user_words:
        if word in combined_lion_text and word not in matched_words:
            matched_words.append(word)

    score += min(len(matched_words) * 10, 50)

    if matched_words:
        reasons.append("Matched words: " + ", ".join(matched_words[:8]))

    if score == 0:
        reasons.append("No strong text match yet")

    return score, reasons


known_lions = load_known_lions()
submissions = load_submissions()
sightings_df = load_sightings()

with st.sidebar:
    st.markdown("""
    <div style="text-align:center;padding:20px 0 16px;">
        <div style="font-size:38px;">🦁</div>
        <div style="font-family:Georgia,serif;font-size:20px;font-weight:900;">LionGuard ID</div>
        <div style="font-size:10px;letter-spacing:.18em;text-transform:uppercase;color:#c8973f;">Ewaso Lions</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    page = st.radio(
        "Navigate",
        [
            "Home", "Known Lions", "Lion Profiles", "Identify Lion", "Full Image Archive",
            "Submit Sighting", "Review Queue", "Add New Lion", "Edit Lion Profile", "Data Dashboard", "About"
        ],
        label_visibility="collapsed"
    )

st.markdown("""
<div class="topbar">
    <span style="font-size:28px;">🦁</span>
    <div>
        <div class="topbar-logo">LionGuard ID</div>
        <div class="topbar-tagline">Community-powered identification · Ewaso Lions Field Tool</div>
    </div>
</div>
""", unsafe_allow_html=True)

if page == "Home":
    col1, col2 = st.columns([1.25, 1], gap="large")
    with col1:
        st.markdown("""
        <div class="hero">
            <div style="color:#e8b96a;font-size:12px;letter-spacing:.18em;text-transform:uppercase;">Ewaso Lions · Field Identification Tool</div>
            <div class="hero-title">LionGuard ID</div>
            <div class="hero-sub">Community-powered lion identification built for the people who know these animals best.</div>
            <span class="hero-btn">Human-in-the-loop · Expert reviewed</span>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
        show_image(HERO_IMAGE)

    pending = len(submissions[submissions["expert_status"] == "Pending expert review"]) if not submissions.empty else 0
    a, b, c, d = st.columns(4)
    stats = [
        (a, len(known_lions), "Known Lions"),
        (b, len(submissions), "Submitted Sightings"),
        (c, pending, "Pending Reviews"),
        (d, len(sightings_df), "Archive Sightings"),
    ]
    for col, num, label in stats:
        with col:
            st.markdown(f"<div class='stat-card'><div class='stat-num'>{num}</div><div class='stat-lbl'>{label}</div></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    x, y, z = st.columns(3)
    for col, img, title, desc in [
        (x, LION_MALE_IMAGE, "Known Lion Archive", "Browse profiles, photos, descriptions, and whisker notes."),
        (y, LIONESS_IMAGE, "Identify Lion", "Compare observed features against known lions."),
        (z, LION_CUBS_IMAGE, "Expert Review", "Support expert confirmation of sightings.")
    ]:
        with col:
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            show_image(img)
            st.subheader(title)
            st.write(desc)
            st.markdown("</div>", unsafe_allow_html=True)

elif page == "Known Lions":
    page_band("Known Lions Archive", "Browse individually identified lions and open full profiles.")

    search = st.text_input("Search", placeholder="Search by name, ID, sex, description, or whisker pattern...", label_visibility="collapsed")
    filtered = known_lions.copy()

    if search:
        q = search.lower()
        filtered = filtered[filtered.apply(lambda r: q in " ".join(r.astype(str)).lower(), axis=1)]

    st.caption(f"{len(filtered)} record(s)")

    for _, row in filtered.iterrows():
        st.markdown("<div class='lion-card'>", unsafe_allow_html=True)
        img_col, info_col = st.columns([1, 2.8], gap="large")

        with img_col:
            show_image(row.get("reference_image", ""))

        with info_col:
            lion_name = clean_value(row.get("lion_name", "Unknown Lion"))
            lion_id = clean_value(row.get("lion_id", ""))
            st.markdown(f"<div class='lion-name'>{lion_name}<span class='lion-id-tag'>#{lion_id}</span></div>", unsafe_allow_html=True)
            st.markdown(
                f"<span class='badge'>{clean_value(row.get('sex', 'Unknown'))}</span>"
                f"<span class='badge'>Age: {clean_value(row.get('age', 'Unknown'))}</span>",
                unsafe_allow_html=True
            )

            field("General Description", row.get("general_description", ""))
            field("Whisker Pattern", row.get("whisker_pattern_description", ""))

            last_sighting, _ = get_last_sighting(sightings_df, lion_name)
            if last_sighting is not None:
                field("Last Known Sighting", f"{last_sighting['Date Sighted'].strftime('%B %d, %Y')} · {clean_value(last_sighting.get('Location', ''))}")

            if st.button(f"View Full Profile: {lion_name}", key=f"profile_{lion_id}"):
                set_selected_lion(lion_name)
                st.info("Open the Lion Profiles tab to view the full profile and gallery.")

        st.markdown("</div>", unsafe_allow_html=True)

elif page == "Lion Profiles":
    page_band("Lion Profiles", "Full profile pages with descriptions, whisker notes, last known sightings, and photo galleries.")

    lion_names = known_lions["lion_name"].dropna().tolist()

    if not lion_names:
        st.warning("No known lions available.")
    else:
        default_index = 0
        if "selected_lion_name" in st.session_state and st.session_state["selected_lion_name"] in lion_names:
            default_index = lion_names.index(st.session_state["selected_lion_name"])

        selected_lion = st.selectbox("Choose a lion profile", lion_names, index=default_index)
        row = known_lions[known_lions["lion_name"] == selected_lion].iloc[0]

        st.markdown("<div class='profile-card'>", unsafe_allow_html=True)
        col_img, col_info = st.columns([1, 2], gap="large")

        with col_img:
            show_image(row.get("reference_image", ""), caption="Main Reference Image")
            whisker_img = row.get("whisker_pattern_image", "")
            if is_valid_path(whisker_img):
                st.image(whisker_img, caption="Whisker Pattern Image", use_container_width=True)

        with col_info:
            st.title(clean_value(row.get("lion_name", "Unknown Lion")))
            st.markdown(
                f"<span class='badge'>{clean_value(row.get('lion_id', ''))}</span>"
                f"<span class='badge'>{clean_value(row.get('sex', 'Unknown'))}</span>"
                f"<span class='badge'>Age: {clean_value(row.get('age', 'Unknown'))}</span>",
                unsafe_allow_html=True
            )
            field("General Description", row.get("general_description", ""))
            field("Whisker Pattern Description", row.get("whisker_pattern_description", ""))

        st.markdown("</div>", unsafe_allow_html=True)

        show_last_sighting(sightings_df, selected_lion)

        st.subheader("Photo Gallery")
        folder = row.get("all_reference_images_folder", "")
        local_images = get_lion_images(folder)
        url_images = split_gallery_urls(row.get("gallery_image_urls", ""))
        images = local_images + url_images

        if not images:
            st.warning("No gallery images found for this lion.")
        else:
            st.write(f"{len(images)} photo(s) found.")
            cols = st.columns(3)
            for i, img in enumerate(images):
                with cols[i % 3]:
                    st.image(img, use_container_width=True)
                    if str(img).startswith("http"):
                        st.caption("Supabase image")
                    else:
                        st.caption(os.path.basename(img))

elif page == "Identify Lion":
    page_band("Identify Lion", "Enter visible features and get possible matches from the known lion database.")

    st.write("This first version uses sex, description keywords, and whisker notes. Photo-based matching will come later.")

    col1, col2 = st.columns([1, 1.3], gap="large")

    with col1:
        uploaded_match_photo = st.file_uploader("Optional: upload lion photo", type=["jpg", "jpeg", "png"], key="identify_upload")
        if uploaded_match_photo:
            image = Image.open(uploaded_match_photo).convert("RGB")
            st.image(image, caption="Uploaded lion photo", use_container_width=True)
        else:
            st.info("Image similarity is not active yet. Use the fields on the right for now.")

    with col2:
        sex_input = st.selectbox("Sex", ["Unknown", "Male", "Female"])
        description_input = st.text_area(
            "General visible features",
            placeholder="Example: left ear notch, scar on nose, dark mane, mark below eye..."
        )
        whisker_input = st.text_area(
            "Whisker pattern notes",
            placeholder="Example: 5 rows on left, dense right side, missing lower spot..."
        )
        run_match = st.button("Find Possible Matches")

    if run_match:
        results = []

        for _, row in known_lions.iterrows():
            score, reasons = score_lion_match(row, sex_input, description_input, whisker_input)
            results.append({
                "lion_id": row.get("lion_id", ""),
                "lion_name": row.get("lion_name", ""),
                "sex": row.get("sex", ""),
                "score": score,
                "reasons": "; ".join(reasons),
                "reference_image": row.get("reference_image", ""),
                "general_description": row.get("general_description", ""),
                "whisker_pattern_description": row.get("whisker_pattern_description", ""),
            })

        results_df = pd.DataFrame(results).sort_values("score", ascending=False)

        st.subheader("Possible Matches")
        top_results = results_df.head(5)

        if top_results["score"].max() == 0:
            st.warning("No strong match found. This may be a new lion or the notes may need more detail.")

        for _, match in top_results.iterrows():
            st.markdown("<div class='lion-card'>", unsafe_allow_html=True)
            img_col, info_col = st.columns([1, 2.5], gap="large")

            with img_col:
                show_image(match["reference_image"])

            with info_col:
                st.markdown(
                    f"<div class='lion-name'>{clean_value(match['lion_name'])}<span class='lion-id-tag'>#{clean_value(match['lion_id'])}</span></div>",
                    unsafe_allow_html=True,
                )
                st.markdown(
                    f"<span class='badge'>{clean_value(match['sex'])}</span><span class='badge'>Match score: {int(match['score'])}</span>",
                    unsafe_allow_html=True,
                )

                field("Why this match", match["reasons"])
                field("General Description", match["general_description"])
                field("Whisker Pattern", match["whisker_pattern_description"])

                if st.button(f"Open Profile: {match['lion_name']}", key=f"open_match_{match['lion_id']}"):
                    set_selected_lion(match["lion_name"])
                    st.info("Open the Lion Profiles tab to view the full profile.")

            st.markdown("</div>", unsafe_allow_html=True)

elif page == "Full Image Archive":
    page_band("Full Image Archive", "Browse all reference photos from the known lion folders.")

    lion_options = ["All Lions"] + known_lions["lion_name"].dropna().tolist()
    selected = st.selectbox("Choose a lion", lion_options)

    if selected == "All Lions":
        selected_rows = known_lions.copy()
    else:
        selected_rows = known_lions[known_lions["lion_name"] == selected].copy()

    all_images = []
    for _, lion_row in selected_rows.iterrows():
        folder = lion_row.get("all_reference_images_folder", "")
        all_images.extend(get_lion_images(folder))
        all_images.extend(split_gallery_urls(lion_row.get("gallery_image_urls", "")))

    st.write(f"Showing {len(all_images)} image(s).")

    if not all_images:
        st.warning("No images found.")
    else:
        cols = st.columns(4)
        for i, img_path in enumerate(all_images):
            with cols[i % 4]:
                st.image(img_path, use_container_width=True)
                st.caption(os.path.basename(img_path))

elif page == "Submit Sighting":
    page_band("Submit a Sighting", "Upload a photo and record visible identification details.")

    uploaded = st.file_uploader("Upload lion photo", type=["jpg", "jpeg", "png"])

    if uploaded:
        image = Image.open(uploaded).convert("RGB")
        img_col, form_col = st.columns([1, 1.4], gap="large")

        with img_col:
            st.markdown("<div class='card'>", unsafe_allow_html=True)
            st.image(image, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with form_col:
            tab1, tab2, tab3 = st.tabs(["Basic Info", "Identifying Features", "Notes"])

            with tab1:
                date_seen = st.date_input("Date seen", value=date.today())
                location = st.text_input("Location / conservancy / area")
                sex = st.selectbox("Sex", ["Unknown", "Male", "Female"])
                age_class = st.selectbox("Age class", ["Unknown", "Cub", "Subadult", "Adult", "Old adult"])
                side_visible = st.selectbox("Side visible", ["Unknown", "Left side", "Right side", "Front face", "Both sides", "Rear/body only"])

            with tab2:
                scars_or_injuries = st.text_area("Scars, injuries, nose or eye marks")
                ear_notches = st.text_area("Ear notches / ear damage")
                whisker_spot_notes = st.text_area("Whisker spot pattern notes")
                mane_description = st.text_area("Mane description")
                tail_or_body_marks = st.text_area("Tail, coat, or other body markings")

            with tab3:
                behavior_context = st.text_area("Behavior / context")
                suspected_lion = st.text_input("Suspected lion name or ID, if any")
                confidence = st.selectbox("Observer confidence", ["Low", "Medium", "High"])
                reviewer_notes = st.text_area("Additional notes")

            if st.button("Save Sighting"):
                existing = load_submissions()
                submission_id = len(existing) + 1
                img_fname = f"submission_{submission_id}_{uploaded.name}"
                fallback_path = os.path.join(UPLOAD_FOLDER, img_fname)
                img_path = save_uploaded_photo(uploaded, "sightings", UPLOAD_FOLDER)

                if not img_path:
                    img_path = fallback_path

                new_row = {
                    "submission_id": submission_id,
                    "image_path": img_path,
                    "date_seen": date_seen,
                    "location": location,
                    "sex": sex,
                    "age_class": age_class,
                    "side_visible": side_visible,
                    "mane_description": mane_description,
                    "scars_or_injuries": scars_or_injuries,
                    "ear_notches": ear_notches,
                    "whisker_spot_notes": whisker_spot_notes,
                    "tail_or_body_marks": tail_or_body_marks,
                    "behavior_context": behavior_context,
                    "suspected_lion": suspected_lion,
                    "confidence": confidence,
                    "reviewer_notes": reviewer_notes,
                    "expert_final_id": "",
                    "expert_status": "Pending expert review",
                }

                updated = pd.concat([existing, pd.DataFrame([new_row])], ignore_index=True)
                save_submissions(updated)
                st.success("Sighting saved and queued for expert review.")
    else:
        st.info("Upload a lion photo to begin.")

elif page == "Review Queue":
    page_band("Review Queue", "Expert review area for submitted lion sightings.")

    if submissions.empty:
        st.info("No sightings have been submitted yet.")
    else:
        status_filter = st.selectbox("Filter", ["All", "Pending expert review", "Reviewed"], label_visibility="collapsed")
        filtered = submissions.copy()
        if status_filter != "All":
            filtered = filtered[filtered["expert_status"] == status_filter]

        for idx, row in filtered.iterrows():
            st.markdown("<div class='lion-card'>", unsafe_allow_html=True)
            img_col, info_col = st.columns([1, 2.2], gap="large")

            with img_col:
                show_image(row.get("image_path", ""))

            with info_col:
                st.markdown(f"<div class='lion-name'>Submission #{clean_value(row.get('submission_id', ''))}</div>", unsafe_allow_html=True)
                field("Date", row.get("date_seen", ""))
                field("Location", row.get("location", ""))
                field("Sex", row.get("sex", ""))
                field("Whisker Notes", row.get("whisker_spot_notes", ""))
                field("Suspected Lion", row.get("suspected_lion", ""))

                lion_opts = ["Unknown / New Lion"] + [f"{r['lion_id']} — {r['lion_name']}" for _, r in known_lions.iterrows()]
                selected = st.selectbox("Expert final ID", lion_opts, key=f"sel_{idx}")

                if st.button("Save Expert Decision", key=f"save_{idx}"):
                    orig = submissions.index[submissions["submission_id"] == row["submission_id"]][0]
                    submissions.loc[orig, "expert_final_id"] = selected
                    submissions.loc[orig, "expert_status"] = "Reviewed"
                    save_submissions(submissions)
                    st.success("Expert decision saved.")

            st.markdown("</div>", unsafe_allow_html=True)

elif page == "Add New Lion":
    page_band("Add New Lion", "Create a new lion profile and upload reference photos from inside LionGuard ID.")

    st.markdown("### New Lion Details")
    st.info("This page saves the lion into known_lions.csv and creates a photo folder inside known_lion_images/.")

    suggested_next_id = f"LG{len(known_lions) + 1:03d}"

    with st.form("add_new_lion_form", clear_on_submit=False):
        lion_name = st.text_input("Lion name", placeholder="Example: Sikiria")
        lion_id = st.text_input("Lion ID", value=suggested_next_id, placeholder="Example: LG019")
        sex = st.selectbox("Sex", ["Unknown", "Male", "Female"])
        age = st.text_input("Age", value="Unknown", placeholder="Unknown, Adult, Subadult, Cub, etc.")

        general_description = st.text_area(
            "General description",
            placeholder="Example: Adult lioness with pale coat and small scar near left eye."
        )

        whisker_pattern_description = st.text_area(
            "Whisker pattern description",
            placeholder="Example: Dense lower whisker spots on right side."
        )

        reference_photos = st.file_uploader(
            "Upload reference photos",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True
        )

        submitted = st.form_submit_button("Add Lion to Archive")

    if submitted:
        clean_lion_name = lion_name.strip()
        clean_lion_id = lion_id.strip()

        if not clean_lion_name:
            st.error("Please enter a lion name.")
        elif not clean_lion_id:
            st.error("Please enter a lion ID.")
        elif known_lions["lion_name"].astype(str).str.lower().str.strip().eq(clean_lion_name.lower()).any():
            st.error("A lion with this name already exists.")
        elif known_lions["lion_id"].astype(str).str.lower().str.strip().eq(clean_lion_id.lower()).any():
            st.error("A lion with this ID already exists.")
        else:
            folder_name = safe_folder_name(clean_lion_name)
            lion_folder = os.path.join(KNOWN_LION_IMAGES_FOLDER, folder_name)
            os.makedirs(lion_folder, exist_ok=True)

            saved_photo_paths = []

            for photo in reference_photos:
                saved_path = save_uploaded_photo(photo, clean_lion_name, lion_folder)
                if saved_path:
                    saved_photo_paths.append(saved_path)

            reference_image = saved_photo_paths[0] if saved_photo_paths else ""

            new_lion = {
                "lion_id": clean_lion_id,
                "lion_name": clean_lion_name,
                "sex": sex,
                "age": age.strip() if age.strip() else "Unknown",
                "general_description": general_description,
                "whisker_pattern_description": whisker_pattern_description,
                "whisker_pattern_image": "",
                "reference_image": reference_image,
                "all_reference_images_folder": lion_folder,
                "gallery_image_urls": join_gallery_urls(saved_photo_paths),
            }

            updated_lions = pd.concat([known_lions, pd.DataFrame([new_lion])], ignore_index=True)
            save_known_lions(updated_lions)

            st.success(f"{clean_lion_name} has been added to the known lion archive.")
            st.write(f"Photo folder created: `{lion_folder}`")
            st.info("The new lion will appear after you refresh the browser or switch pages. If it does not, stop Streamlit and run it again.")

elif page == "Edit Lion Profile":
    page_band("Edit Lion Profile", "Update lion information, age, descriptions, and reference images.")

    lion_names = known_lions["lion_name"].dropna().tolist()

    if not lion_names:
        st.warning("No lions available to edit.")
    else:
        selected_lion = st.selectbox("Select lion to edit", lion_names)

        lion_index = known_lions[known_lions["lion_name"] == selected_lion].index[0]
        row = known_lions.loc[lion_index]

        st.subheader("Current Profile")
        c1, c2 = st.columns([1, 2], gap="large")

        with c1:
            show_image(row.get("reference_image", ""), caption="Current Reference Image")

        with c2:
            field("Lion ID", row.get("lion_id", ""))
            field("Name", row.get("lion_name", ""))
            field("Sex", row.get("sex", ""))
            field("Age", row.get("age", "Unknown"))
            field("General Description", row.get("general_description", ""))
            field("Whisker Pattern", row.get("whisker_pattern_description", ""))

        st.divider()
        st.subheader("Edit Profile")

        with st.form("edit_lion_profile_form"):
            new_id = st.text_input("Lion ID", value=str(row.get("lion_id", "")))
            new_name = st.text_input("Lion Name", value=str(row.get("lion_name", "")))

            sex_options = ["Unknown", "Male", "Female"]
            current_sex = str(row.get("sex", "Unknown"))
            sex_index = sex_options.index(current_sex) if current_sex in sex_options else 0
            new_sex = st.selectbox("Sex", sex_options, index=sex_index)

            new_age = st.text_input(
                "Age",
                value=str(row.get("age", "Unknown") or "Unknown"),
                placeholder="Unknown, Adult, Subadult, Cub, 5 years, etc."
            )

            new_description = st.text_area(
                "General Description",
                value=str(row.get("general_description", ""))
            )

            new_whiskers = st.text_area(
                "Whisker Pattern Description",
                value=str(row.get("whisker_pattern_description", ""))
            )

            uploaded_photos = st.file_uploader(
                "Add new reference/gallery photos",
                type=["jpg", "jpeg", "png"],
                accept_multiple_files=True
            )

            make_first_uploaded_reference = st.checkbox(
                "Use first uploaded photo as the main reference image",
                value=False
            )

            save_changes = st.form_submit_button("Save Profile Changes")

        if save_changes:
            clean_name = new_name.strip()
            clean_id = new_id.strip()

            if not clean_name:
                st.error("Lion name cannot be blank.")
            elif not clean_id:
                st.error("Lion ID cannot be blank.")
            else:
                lion_folder = str(row.get("all_reference_images_folder", "")).strip()
                if not lion_folder:
                    lion_folder = os.path.join(KNOWN_LION_IMAGES_FOLDER, safe_folder_name(clean_name))

                existing_urls = split_gallery_urls(row.get("gallery_image_urls", ""))
                new_urls = []

                for photo in uploaded_photos:
                    saved_path = save_uploaded_photo(photo, clean_name, lion_folder)
                    if saved_path:
                        new_urls.append(saved_path)

                all_gallery_urls = existing_urls + new_urls

                known_lions.loc[lion_index, "lion_id"] = clean_id
                known_lions.loc[lion_index, "lion_name"] = clean_name
                known_lions.loc[lion_index, "sex"] = new_sex
                known_lions.loc[lion_index, "age"] = new_age.strip() if new_age.strip() else "Unknown"
                known_lions.loc[lion_index, "general_description"] = new_description
                known_lions.loc[lion_index, "whisker_pattern_description"] = new_whiskers
                known_lions.loc[lion_index, "all_reference_images_folder"] = lion_folder
                known_lions.loc[lion_index, "gallery_image_urls"] = join_gallery_urls(all_gallery_urls)

                if new_urls and make_first_uploaded_reference:
                    known_lions.loc[lion_index, "reference_image"] = new_urls[0]

                save_known_lions(known_lions)

                st.success(f"{clean_name}'s profile has been updated.")
                st.info("Refresh or switch pages to see the updated profile.")

elif page == "Data Dashboard":
    page_band("Data Dashboard", "Overview of known lions, submitted sightings, and sighting archive data.")
    pending = len(submissions[submissions["expert_status"] == "Pending expert review"]) if not submissions.empty else 0
    reviewed = len(submissions[submissions["expert_status"] == "Reviewed"]) if not submissions.empty else 0

    a, b, c, d = st.columns(4)
    a.metric("Known Lions", len(known_lions))
    b.metric("Pending Reviews", pending)
    c.metric("Reviewed Sightings", reviewed)
    d.metric("Archive Sightings", len(sightings_df))

    st.subheader("Known Lions")
    st.dataframe(known_lions, use_container_width=True, hide_index=True)

    st.subheader("Submissions")
    st.dataframe(submissions, use_container_width=True, hide_index=True)

    st.subheader("Sighting Archive")
    if sightings_df.empty:
        st.info("No lion_sightings.xlsx file found yet.")
    else:
        st.dataframe(sightings_df.sort_values("Date Sighted", ascending=False), use_container_width=True, hide_index=True)

elif page == "About":
    page_band("About LionGuard ID")
    col1, col2 = st.columns([1, 1.3], gap="large")

    with col1:
        show_image(LION_FAMILY_IMAGE)
        show_image(LIONESS_IMAGE)

    with col2:
        st.markdown("""
        <div class="card">
        <h2>Purpose</h2>
        <p>LionGuard ID is a prototype tool for recording lion sightings, organizing local ecological knowledge, and supporting expert identification of known lions.</p>
        <h2>Current Version</h2>
        <p>This version includes known lion profiles, full image galleries, sighting submissions, expert review, last known sightings, and an Add New Lion page.</p>
        <h2>Future Direction</h2>
        <p>The long-term goal is AI-assisted identification using reference photos, whisker patterns, scars, ear notches, and Ewaso Lions field knowledge.</p>
        </div>
        """, unsafe_allow_html=True)

st.markdown("""
<div class="site-footer">
    <span>LionGuard ID</span> · Ewaso Lions · Prototype V8 · Human-in-the-loop
</div>
""", unsafe_allow_html=True)
