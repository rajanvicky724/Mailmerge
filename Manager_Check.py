import streamlit as st
import time
import random
import speech_recognition as sr
from thefuzz import process
import io
import base64
import asyncio
import edge_tts  # PIP INSTALL edge-tts
import datetime

# --- PAGE CONFIG ---
st.set_page_config(page_title="ஞான ஜோதிடம் Premium", page_icon="🕉️", layout="wide")

# --- CSS STYLING ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Hind+Madurai:wght@400;700&display=swap');
    .stApp { background-color: #0E1117; color: #E0E0E0; font-family: 'Hind Madurai', sans-serif; }
    .premium-card {
        background: linear-gradient(160deg, #1a1a2e, #16213e); 
        padding: 40px; border-radius: 20px; border: 2px solid #D4AF37; 
        box-shadow: 0 0 30px rgba(212, 175, 55, 0.2); color: white;
    }
    .header-gold { color: #D4AF37; text-align: center; margin-bottom: 5px; font-weight: bold; }
    .sub-header { color: #A0A0A0; text-align: center; font-size: 18px; margin-bottom: 25px; }
    .section-title {
        color: #00FFFF; font-weight: bold; font-size: 22px; margin-top: 25px; 
        border-bottom: 1px solid #333; padding-bottom: 5px;
    }
    .content-text { font-size: 18px; color: #FFF; line-height: 1.8; margin-top: 10px; }
    .planet-box {
        background-color: rgba(0, 0, 0, 0.3); padding: 15px; border-left: 4px solid #D4AF37;
        margin-top: 20px; color: #FFD700;
    }
    .roast-box {
        background-color: #330000; color: #FF9999; padding: 15px; border-radius: 10px;
        margin-top: 25px; text-align: center; font-weight: bold; border: 1px dashed #FF5555; font-size: 18px;
    }
    .quote-box {
        background: linear-gradient(90deg, #141E30, #243B55);
        text-align: center; font-style: italic; font-size: 18px; color: #FFD700; 
        margin-top: 30px; padding: 20px; border-radius: 10px; border: 1px solid #D4AF37;
    }
    .hidden-audio { display: none; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #D4AF37;'>✨ ஞான ஜோதிடம் Premium ✨</h1>", unsafe_allow_html=True)

# --- 🔊 AUDIO ENGINE ---
async def generate_audio_edge_async(text, filename, voice, rate, pitch):
    if not text or not text.strip(): return None
    try:
        communicate = edge_tts.Communicate(text, voice, rate=rate, pitch=pitch)
        await communicate.save(filename)
        return filename
    except: return None

def generate_audio(text, filename, voice="ta-IN-ValluvarNeural", rate="+0%", pitch="+0Hz"):
    try:
        try: loop = asyncio.get_event_loop()
        except RuntimeError: loop = asyncio.new_event_loop(); asyncio.set_event_loop(loop)
        if loop.is_closed(): loop = asyncio.new_event_loop(); asyncio.set_event_loop(loop)
        loop.run_until_complete(generate_audio_edge_async(text, filename, voice, rate, pitch))
        return filename
    except: return None

def autoplay_audio_hidden(file_path):
    try:
        with open(file_path, "rb") as f: data = f.read()
        b64 = base64.b64encode(data).decode()
        md = f"""<audio autoplay="true" class="hidden-audio"><source src="data:audio/mp3;base64,{b64}" type="audio/mp3"></audio>"""
        st.markdown(md, unsafe_allow_html=True)
    except: pass

def recognize_audio_bytes(audio_input, lang="en-IN"):
    if not audio_input: return None
    r = sr.Recognizer()
    try:
        audio_bytes = audio_input.getvalue()
        audio_file = io.BytesIO(audio_bytes)
        with sr.AudioFile(audio_file) as source:
            r.adjust_for_ambient_noise(source, duration=0.5)
            audio_data = r.record(source)
            return r.recognize_google(audio_data, language=lang)
    except: return None

# --- TAMIL TO ENGLISH STAR MAPPING ---
TAMIL_TO_ENGLISH_STARS = {
    "அசுவினி": "Ashwini", "பரணி": "Bharani", "கிருத்திகை": "Krittikai", "ரோகிணி": "Rohini", 
    "மிருகசீரிடம்": "Mrigashirsham", "திருவாதிரை": "Thiruvathirai", "புனர்பூசம்": "Punarpoosam", 
    "பூசம்": "Poosam", "ஆயில்யம்": "Ayilyam", "மகம்": "Magam", "பூரம்": "Pooram", 
    "உத்திரம்": "Uthiram", "ஹஸ்தம்": "Hastham", "சித்திரை": "Chithirai", "சுவாதி": "Swathi", 
    "விசாகம்": "Visakam", "அனுஷம்": "Anusham", "கேட்டை": "Kettai", "மூலம்": "Moolam", 
    "பூராடம்": "Pooradam", "உத்திராடம்": "Uthiradam", "திருவோணம்": "Thiruvonam", "அவிட்டம்": "Avittam", 
    "சதயம்": "Sathayam", "பூரட்டாதி": "Poorattathi", "உத்திரட்டாதி": "Uthirattathi", "ரேவதி": "Revathi"
}

def smart_match(voice_text, options_list):
    if not voice_text: return None
    for tamil_name, eng_name in TAMIL_TO_ENGLISH_STARS.items():
        if tamil_name in voice_text:
            voice_text = eng_name
            break
    clean_options = [opt.split(" (")[0].strip() for opt in options_list]
    best_match_clean, score = process.extractOne(voice_text, clean_options)
    if score > 50:
        index = clean_options.index(best_match_clean)
        return options_list[index]
    return None

def get_name_analysis(name):
    if not name: return ""
    first = name[0].upper()
    return {"A": "சூரியனின் ஆதிக்கம். நீங்கள் பிறருக்கு கட்டளையிடுவதை விரும்புவீர்கள்.", 
            "B": "பாசமானவர். பணத்தை விட உறவுக்கு முக்கியத்துவம் கொடுப்பவர்.", 
            "C": "நகைச்சுவை உணர்வு மிக்கவர். கவலையை வெளிக்காட்ட மாட்டீர்கள்.", 
            "D": "நம்பிக்கையானவர். எதையும் துணிச்சலாக எதிர்கொள்வீர்கள்.", 
            "M": "கடின உழைப்பாளி. குடும்பத்திற்காக எதையும் செய்வீர்கள்.", 
            "S": "கவர்ச்சியானவர். கூட்டத்தில் தனித்து தெரிவீர்கள்.", 
            "R": "உதவும் குணம். மற்றவர் நலனில் அக்கறை கொள்வீர்கள்.", 
            "V": "வெற்றியாளர். தோல்வியை கண்டு துவள மாட்டீர்கள்."}.get(first, "உங்கள் பெயரில் ஒரு தனித்துவமான காந்த சக்தி உள்ளது.")

def calculate_life_path_number(dob):
    s = str(dob).replace('-', '')
    total = sum(int(digit) for digit in s)
    while total > 9: total = sum(int(digit) for digit in str(total))
    return total

# --- TEAM MANAGER DATA ---
MANAGER_DATA = {
    "Prabha": "Prabha! Neenga our thivira muruga bhaktar. Verithanamana Rajini fan. Pathfinder la neenga oru AI manidhar.",
    "Suji": "Suji! You are good dancer. Recently your dance in insta crossed 3 million views. After karakata kanaga you have more fans for karakattam. people call you 'Rajamadha sivakami' devi.",
    "Ramzy": "Ramzy! kaara sweet pola irukum neenga tha inga saara sweet ah fast and furious ku our paul walker uh pathfinder ku neenga tha win diesel uh jakkamma \"Bhai amma\"",
    "Nandha Kumar S": "Nandha Kumar! NANDHA means PFBA’s Kavalan SOS. Magalirku kashtamna karnana mariduvan intha Nandhavanathu Nandha.",
    "Sugumar C": "Sugumar! Meeting start aagum… Aana Sugumar pesinaa, silence automatic-aa varum. Yennaa… Anga vaarthai illa, decision irukkum. Discussion illa, direction irukkum. Neenga sound raise panna maatinga… Aana clarity raise pannuvinga.",
    "Krishna Kumar M": "Krishna! Deadline-ah miss panninaa… reason ketka maatten. Root cause ketpen. B’coz I am Strict TL",
    "Avinash H": "Avinash… nenga search panina! Ethum solla matinga… Aana ellarukum theriyum, problem-e neenga kandupidikirenga! Task unga dhaan illa… Aana solution unga pakkam irukku. Problem side-la confidence irundhaa… Avinash side-la clarity irukkum!"
}

# --- DATA ---
RASI_NAMES = ["மேஷம் (Aries)", "ரிஷபம் (Taurus)", "மிதுனம் (Gemini)", "கடகம் (Cancer)", "சிம்மம் (Leo)", "கன்னி (Virgo)", "துலாம் (Libra)", "விருச்சிகம் (Scorpio)", "தனுசு (Sagittarius)", "மகரம் (Capricorn)", "கும்பம் (Aquarius)", "மீனம் (Pisces)"]

RASI_TO_STARS = {
    "மேஷம் (Aries)": ["Ashwini", "Bharani", "Krittikai (1)"], 
    "ரிஷபம் (Taurus)": ["Krittikai (2-4)", "Rohini", "Mrigashirsham (1-2)"], 
    "மிதுனம் (Gemini)": ["Mrigashirsham (3-4)", "Thiruvathirai", "Punarpoosam (1-3)"], 
    "கடகம் (Cancer)": ["Punarpoosam (4)", "Poosam", "Ayilyam"], 
    "சிம்மம் (Leo)": ["Magam", "Pooram", "Uthiram (1)"], 
    "கன்னி (Virgo)": ["Uthiram (2-4)", "Hastham", "Chithirai (1-2)"], 
    "துலாம் (Libra)": ["Chithirai (3-4)", "Swathi", "Visakam (1-3)"], 
    "விருச்சிகம் (Scorpio)": ["Visakam (4)", "Anusham", "Kettai"], 
    "தனுசு (Sagittarius)": ["Moolam", "Pooradam", "Uthiradam (1)"], 
    "மகரம் (Capricorn)": ["Uthiradam (2-4)", "Thiruvonam", "Avittam (1-2)"], 
    "கும்பம் (Aquarius)": ["Avittam (3-4)", "Sathayam", "Poorattathi (1-3)"], 
    "மீனம் (Pisces)": ["Poorattathi (4)", "Uthirattathi", "Revathi"]
}

RASI_DESC = {
    "மேஷம் (Aries)": "நீங்கள் தைரியமானவர். எதற்கும் அஞ்சாதவர். எடுத்த முடிவை மாற்ற மாட்டீர்கள். கோபம் வந்தாலும் உடனே ஆறிவிடுவீர்கள்.",
    "ரிஷபம் (Taurus)": "நீங்கள் நிதானமானவர். பொறுமைசாலி. கலை மற்றும் அழகில் ஈடுபாடு அதிகம். நம்பியவர்களை கைவிட மாட்டீர்கள்.",
    "மிதுனம் (Gemini)": "நீங்கள் புத்திசாலி. பேச்சாலேயே எதையும் சாதிப்பீர்கள். ஒரே நேரத்தில் பல வேலைகளை செய்வீர்கள். நகைச்சுவை உணர்வு அதிகம்.",
    "கடகம் (Cancer)": "நீங்கள் பாசமானவர். குடும்பத்திற்கு முக்கியத்துவம் கொடுப்பவர். மன உறுதி மிக்கவர். பிறருக்கு உதவும் குணம் உண்டு.",
    "சிம்மம் (Leo)": "நீங்கள் ராஜ கம்பீரம் கொண்டவர். எதிலும் முதலிடத்தில் இருக்க விரும்புவீர்கள். தாராள குணம் கொண்டவர். உங்கள் சொல்லுக்கு மதிப்பு அதிகம்.",
    "கன்னி (Virgo)": "நீங்கள் நுணுக்கமான அறிவு கொண்டவர். சுத்தம் மற்றும் ஒழுக்கத்தை விரும்புவீர்கள். எதையும் திட்டமிட்டு செய்வீர்கள்.",
    "துலாம் (Libra)": "நீங்கள் நீதி நேர்மை தவறாதவர். அனைவரிடமும் அன்பாக பழகுவீர்கள். கலை ரசனை மிக்கவர். சமாதானத்தை விரும்புவீர்கள்.",
    "விருச்சிகம் (Scorpio)": "நீங்கள் சுறுசுறுப்பானவர். ரகசியங்களை காப்பதில் வல்லவர். எதையும் ஆழமாக சிந்திப்பவர். எதிரிகளை வெல்லும் திறமை கொண்டவர்.",
    "தனுசு (Sagittarius)": "நீங்கள் நேர்மையானவர். ஆன்மீகத்தில் ஈடுபாடு அதிகம். குறிக்கோளை அடைய போராடுவீர்கள். சுதந்திரத்தை விரும்புவீர்கள்.",
    "மகரம் (Capricorn)": "நீங்கள் கடின உழைப்பாளி. பொறுமை மற்றும் சகிப்புத்தன்மை கொண்டவர். லட்சியத்தை அடைய விடாமல் முயற்சிப்பவர்.",
    "கும்பம் (Aquarius)": "நீங்கள் புதுமையை விரும்புபவர். நண்பர்களுக்கு முக்கியத்துவம் கொடுப்பவர். ஆராய்ச்சி மனப்பான்மை கொண்டவர். சுதந்திர பறவை.",
    "மீனம் (Pisces)": "நீங்கள் இரக்க குணம் கொண்டவர். கற்பனை வளம் அதிகம். பிறர் கஷ்டத்தை கண்டு வருந்துவீர்கள். தெய்வீக நம்பிக்கை அதிகம்."
}

PLANET_CONTEXT = {
    "மேஷம் (Aries)": "சனி 12-ல் விரய சனி. செலவுகள் அதிகரிக்கும் ஆனால் குரு 4-ல் இருப்பதால் வீடு, வாகனம் வாங்கும் யோகம் உண்டு.",
    "ரிஷபம் (Taurus)": "சனி 11-ல் லாப சனி. தொட்டதெல்லாம் துலங்கும் பொற்காலம். குரு 3-ல் இருப்பதால் சிறு இடமாற்றம் ஏற்படலாம்.",
    "மிதுனம் (Gemini)": "சனி 10-ல் கர்ம சனி. வேலைப்பளு அதிகரிக்கும். ஆனால் குரு 2-ல் இருப்பதால் தன வரவு தாராளமாக இருக்கும்.",
    "கடகம் (Cancer)": "சனி 9-ல் பாக்கிய சனி. தந்தை வழி சொத்துக்கள் கிடைக்கும். குரு ஜென்மத்தில் இருப்பதால் ஆரோக்கியத்தில் கவனம் தேவை.",
    "சிம்மம் (Leo)": "சனி 8-ல் அஷ்டம சனி. யாருக்கும் ஜாமீன் கையெழுத்து போட வேண்டாம். குரு 12-ல் இருப்பதால் சுப விரயங்கள் நடக்கும்.",
    "கன்னி (Virgo)": "சனி 7-ல் கண்டக சனி. கணவன்-மனைவி இடையே விட்டுக்கொடுத்து செல்லவும். குரு 11-ல் இருப்பதால் நினைத்த காரியம் நடக்கும்.",
    "துலாம் (Libra)": "சனி 6-ல் விபரீத ராஜயோகம். எதிரிகள் விலகுவார்கள். கடன் தொல்லை தீரும். குரு 10-ல் இருப்பதால் தொழில் மாற்றம் வரலாம்.",
    "விருச்சிகம் (Scorpio)": "சனி 5-ல் பூர்வ புண்ணிய சனி. குலதெய்வ அருள் கிடைக்கும். குரு 9-ல் இருப்பதால் பாக்கியங்கள் தேடி வரும்.",
    "தனுசு (Sagittarius)": "சனி 4-ல் அர்த்தாஷ்டம சனி. தாயார் உடல்நிலையில் கவனம் தேவை. குரு 8-ல் இருப்பதால் திடீர் அதிர்ஷ்டம் உண்டாகும்.",
    "மகரம் (Capricorn)": "சனி 3-ல் தைரிய சனி. ஏழரை சனி முடிவுக்கு வருகிறது. குரு 7-ல் இருப்பதால் திருமண யோகம் கைகூடும்.",
    "கும்பம் (Aquarius)": "சனி 2-ல் ஜென்ம சனி. பேச்சில் நிதானம் தேவை. குரு 6-ல் இருப்பதால் கடன் வாங்குவதை தவிர்க்கவும்.",
    "மீனம் (Pisces)": "சனி 1-ல் விரய சனி. அலைச்சல் அதிகரிக்கும். குரு 5-ல் இருப்பதால் புத்திர பாக்கியம் மற்றும் பூர்வ புண்ணிய பலன்கள் கிடைக்கும்."
}

CAREER_DATA = {
    "மேஷம் (Aries)": "வெளிநாட்டு வேலை வாய்ப்புகள் தேடி வரும். ஏற்றுமதி இறக்குமதி தொழில் செய்பவர்களுக்கு லாபம் அதிகரிக்கும்.",
    "ரிஷபம் (Taurus)": "நீண்ட நாட்களாக எதிர்பார்த்த பதவி உயர்வு கிடைக்கும். சுயதொழில் செய்பவர்களுக்கு அபரிமிதமான லாபம் உண்டு.",
    "மிதுனம் (Gemini)": "வேலைப்பளு கூடினாலும் உழைப்புக்கு ஏற்ற ஊதியம் கிடைக்கும். மேலதிகாரிகளின் பாராட்டு கிடைக்கும்.",
    "கடகம் (Cancer)": "தலைமை பண்பு அதிகரிக்கும். புதிய பொறுப்புகள் உங்களை தேடி வரும். அரசு வேலைக்கு முயற்சிப்பவர்களுக்கு வெற்றி.",
    "சிம்மம் (Leo)": "வேலையில் திடீர் இடமாற்றம் ஏற்பட வாய்ப்புள்ளது. சக ஊழியர்களிடம் கவனமாக இருக்கவும்.",
    "கன்னி (Virgo)": "தொழிலில் இருந்த போட்டி பொறாமைகள் விலகும். புதிய தொழில் ஒப்பந்தங்கள் கையெழுத்தாகும்.",
    "துலாம் (Libra)": "இது ராஜயோக காலம். எதிரிகள் காணாமல் போவார்கள். வியாபாரத்தில் கொடிகட்டி பறப்பீர்கள்.",
    "விருச்சிகம் (Scorpio)": "கடைசி நேரத்தில் அதிர்ஷ்டம் கைகொடுக்கும். பங்குச்சந்தை முதலீடுகளில் நிதானம் தேவை.",
    "தனுசு (Sagittarius)": "ஆவணங்களில் கையெழுத்திடும் முன் கவனம் தேவை. ரியல் எஸ்டேட் தொழில் மந்தமாக இருக்கும்.",
    "மகரம் (Capricorn)": "ஏழரை சனி முடிவதால் தொழில் முடக்கம் நீங்கும். பணப்புழக்கம் அதிகரிக்கும்.",
    "கும்பம் (Aquarius)": "பொறுப்புகள் கூடும். கடன் வாங்கி தொழில் விரிவாக்கம் செய்வதை தவிர்க்கவும்.",
    "மீனம் (Pisces)": "சிக்கனம் தேவை. தேவையற்ற செலவுகளை குறைப்பது தொழிலுக்கு நல்லது. வெளிநாட்டு பயணம் சாத்தியம்."
}

RELATIONSHIP_DATA = {
    "Single": "நண்பர்கள் மூலமாக காதல் மலரும். பெற்றோர்கள் சம்மதத்துடன் திருமணம் கைகூடும்.",
    "Married": "கணவன்-மனைவி இடையே இருந்த கருத்து வேறுபாடுகள் நீங்கி ஒற்றுமை பலப்படும். குழந்தை பாக்கியம் உண்டு."
}

HOBBY_DATA = {
    "Ashwini": "குதிரை ஏற்றம் மற்றும் வேகமான விளையாட்டு.", "Bharani": "ஓவியம் மற்றும் கலை பொருட்கள் சேகரிப்பு.", 
    "Krittikai": "சமையல் கலை மற்றும் விவாதம்.", "Rohini": "ஆடை வடிவமைப்பு (Fashion) மற்றும் தோட்டம்.", 
    "Mrigashirsham": "பயணம் மற்றும் புதிய இடங்களை தேடுதல்.", "Thiruvathirai": "கணினி மென்பொருள் மற்றும் தொழில்நுட்பம்.", 
    "Punarpoosam": "வில்வித்தை மற்றும் ஆன்மீக பயணம்.", "Poosam": "தியானம் மற்றும் விவசாயம்.", 
    "Ayilyam": "பங்கு வர்த்தகம் மற்றும் உளவியல்.", "Magam": "வரலாறு மற்றும் பாரம்பரிய கலை.", 
    "Pooram": "புகைப்படம் எடுத்தல் மற்றும் அலங்காரம்.", "Uthiram": "சமூக சேவை மற்றும் நிர்வாகம்.", 
    "Hastham": "கையெழுத்து கலை மற்றும் கைவினை.", "Chithirai": "கட்டிடக்கலை மற்றும் வடிவமைப்பு.", 
    "Swathi": "மார்க்கெட்டிங் மற்றும் ட்ரோன் இயக்குதல்.", "Visakam": "மேடை பேச்சு மற்றும் வியாபாரம்.", 
    "Anusham": "நிகழ்ச்சி ஒருங்கிணைப்பு (Event Mgmt).", "Kettai": "ஜோதிடம் மற்றும் துப்பறிதல்.", 
    "Moolam": "யோகா மற்றும் மூலிகை மருத்துவம்.", "Pooradam": "நீர் விளையாட்டு மற்றும் நீச்சல்.", 
    "Uthiradam": "அரசு தேர்வுகள் மற்றும் சட்டம்.", "Thiruvonam": "இசை மற்றும் பாட்காஸ்ட்.", 
    "Avittam": "இசைக்கருவி வாசித்தல் (Drums).", "Sathayam": "வானியல் மற்றும் விண்வெளி ஆய்வு.", 
    "Poorattathi": "நிதி மேலாண்மை மற்றும் வங்கி.", "Uthirattathi": "தத்துவம் மற்றும் தியானம்.", 
    "Revathi": "நீச்சல் மற்றும் பிராணிகள் வளர்ப்பு."
}

FUNNY_ROASTS = {
    "மேஷம் (Aries)": "எச்சரிக்கை: உங்கள் கோபம் 5G இண்டர்நெட்டை விட வேகமானது. கொஞ்சம் அமைதி ப்ளீஸ்!",
    "ரிஷபம் (Taurus)": "இதை படிக்கும் போது கூட எதாவது சாப்பிட்டுட்டு இருப்பீங்க. சாப்பாடு தான் உலகம்!",
    "மிதுனம் (Gemini)": "உங்க பிரவுசர்ல 50 டேப் (Tab) ஓபன்ல இருக்கும், அதே மாதிரி தான் உங்க மூளையும். எதை ஃபாலோ பண்றது?",
    "கடகம் (Cancer)": "பழைய பகையை மறக்கவே மாட்டீங்க, 10 வருஷம் ஆனாலும் வச்சு செய்வீங்க. பயங்கரமான ஆளு நீங்க!",
    "சிம்மம் (Leo)": "சூரியன் உங்களை சுத்தி வரல, கொஞ்சம் பந்தாவை குறைங்க பாஸ். இவ்ளோ கெத்து ஆகாது!",
    "கன்னி (Virgo)": "சாப்பிடும் போது கூட டேபிள் துடைப்பீங்க, சுத்தம் முக்கியம் பிகிலு! ஆனா மத்தவங்கள டார்ச்சர் பண்ணாதீங்க.",
    "துலாம் (Libra)": "ஒரு ஹோட்டல்ல என்ன சாப்பிடணும்னு முடிவு பண்ணவே உங்களுக்கு 1 மணி நேரம் ஆகும். சீக்கிரம் முடிவு எடுங்க!",
    "விருச்சிகம் (Scorpio)": "உங்க போன்ல ஒரு சீக்ரெட் ஃபோல்டர் இருக்குன்னு எங்களுக்கு தெரியும். பாஸ்வேர்ட் என்ன?",
    "தனுசு (Sagittarius)": "கையில காசு இருக்காது, ஆனா டூர் (Tour) போக பிளான் பண்ணுவீங்க. பட்ஜெட் பாருங்க பாஸ்!",
    "மகரம் (Capricorn)": "வேலையையே கட்டிட்டு அழுறீங்க, கொஞ்சம் ஜாலியா இருங்க. வாழ்க்கை வாழ்வதற்கே!",
    "கும்பம் (Aquarius)": "மெசேஜ் பார்த்தா கூட 3 நாள் கழிச்சு தான் ரிப்ளை பண்ணுவீங்க. நீங்க வேற்றுகிரக வாசிகள்!",
    "மீனம் (Pisces)": "ஒரு நாளைக்கு 5 தடவ லவ் பண்ணுவீங்க (கனவுல!). ரியாலிட்டிக்கு வாங்க!"
}

QUOTES_DATA = [
    "முயற்சி உடையார் இகழ்ச்சி அடையார்.",
    "வல்லவனுக்கு புல்லும் ஆயுதம்.",
    "எண்ணம் போல் வாழ்க்கை.",
    "காலம் பொன் போன்றது.",
    "வாய்மையே வெல்லும்.",
    "கற்றது கைமண் அளவு, கல்லாதது உலகளவு.",
    "நோயற்ற வாழ்வே குறைவற்ற செல்வம்.",
    "அன்பே சிவம்.",
    "பொறுமை கடலினும் பெரிது.",
    "வினையை விதைத்தவன் வினையை அறுப்பான்."
]

NUMEROLOGY_DATA = {1: "தலைவர் (Leader)", 2: "உணர்ச்சியாளர் (Emotional)", 3: "அறிவாளி (Intellectual)", 4: "வித்தியாசமானவர் (Unique)", 5: "புத்திசாலி (Smart)", 6: "கலைஞர் (Artistic)", 7: "சிந்தனையாளர் (Thinker)", 8: "உழைப்பாளி (Hardworker)", 9: "வீரமானவர் (Brave)"}

# --- SIDEBAR: VOICE CONTROL ---
with st.sidebar:
    st.header("✨ வாய்ஸ் அசிஸ்டென்ட்")
    
    voice_map = {"பெண் (Pallavi - IN)": "ta-IN-PallaviNeural", "ஆண் (Valluvar - IN)": "ta-IN-ValluvarNeural"}
    voice_choice = st.selectbox("Select Voice", list(voice_map.keys()), index=1)
    selected_voice = voice_map[voice_choice]
    st.divider()

    # --- NEW: TEAM MANAGER SELECTOR ---
    st.header("👥 Team")
    manager_choice = st.selectbox("Select Manager", ["Select..."] + list(MANAGER_DATA.keys()))
    
    if manager_choice != "Select...":
        if st.button("🔊 Play for " + manager_choice):
            m_text = MANAGER_DATA[manager_choice]
            # Use English compatible voice for mixed text if possible, else Tamil one is fine
            generate_audio(m_text, "manager_speech.mp3", selected_voice)
            autoplay_audio_hidden("manager_speech.mp3")
            st.success(f"Playing audio for {manager_choice}...")

    st.divider()

    if 'v_name' not in st.session_state: st.session_state.v_name = ""
    if 'v_rasi' not in st.session_state: st.session_state.v_rasi = RASI_NAMES[0]
    if 'v_star' not in st.session_state: st.session_state.v_star = RASI_TO_STARS[RASI_NAMES[0]][0]

    # --- 1. NAME ---
    st.markdown("### 1. பெயர் (Name)")
    if st.button("🔊 Guide", key="guide_name"):
        generate_audio("பெயரை சொல்லுங்கள்", "guide_name.mp3", selected_voice)
        autoplay_audio_hidden("guide_name.mp3")

    audio_name_input = st.audio_input("Record Name", key="name_rec_native")
    if audio_name_input:
        text = recognize_audio_bytes(audio_name_input, lang="en-IN")
        if text:
            clean_name = text.title()
            st.session_state.v_name = clean_name
            st.success(f"Name: {clean_name}")
            generate_audio(f"வணக்கம் {clean_name}, ராசியை சொல்லுங்கள்.", "reply_name.mp3", selected_voice)
            autoplay_audio_hidden("reply_name.mp3")

    name = st.text_input("பெயர்", value=st.session_state.v_name)

    # --- 2. RASI ---
    st.markdown("### 2. ராசி (Rasi)")
    if st.button("🔊 Guide", key="guide_rasi"):
        generate_audio("ராசியை சொல்லுங்கள்", "guide_rasi.mp3", selected_voice)
        autoplay_audio_hidden("guide_rasi.mp3")

    audio_rasi_input = st.audio_input("Record Rasi", key="rasi_rec_native")
    if audio_rasi_input:
        text = recognize_audio_bytes(audio_rasi_input, lang="ta-IN")
        if not text: text = recognize_audio_bytes(audio_rasi_input, lang="en-IN")
        
        matched_rasi = smart_match(text, RASI_NAMES)
        if matched_rasi:
            st.session_state.v_rasi = matched_rasi
            st.success(f"Rasi: {matched_rasi}")
            st.session_state.v_star = RASI_TO_STARS[matched_rasi][0]
            desc_text = f"{matched_rasi} ராசி. {RASI_DESC[matched_rasi]}. நட்சத்திரம் சொல்லுங்கள்."
            generate_audio(desc_text, "reply_rasi.mp3", selected_voice)
            autoplay_audio_hidden("reply_rasi.mp3")

    rasi = st.selectbox("ராசி", RASI_NAMES, index=RASI_NAMES.index(st.session_state.v_rasi))
    st.info(f"📜 **{rasi} குணங்கள்:**\n\n{RASI_DESC[rasi]}")

    # --- 3. STAR ---
    st.markdown("### 3. நட்சத்திரம் (Star)")
    if st.button("🔊 Guide", key="guide_star"):
        generate_audio("நட்சத்திரத்தை சொல்லுங்கள்", "guide_star.mp3", selected_voice)
        autoplay_audio_hidden("guide_star.mp3")

    audio_star_input = st.audio_input("Record Star", key="star_rec_native")
    if audio_star_input:
        text = recognize_audio_bytes(audio_star_input, lang="ta-IN")
        if not text: text = recognize_audio_bytes(audio_star_input, lang="en-IN")
        if text:
            all_stars = [s for stars in RASI_TO_STARS.values() for s in stars]
            matched_star = smart_match(text, all_stars)
            if matched_star:
                st.session_state.v_star = matched_star
                st.success(f"Star: {matched_star}")
                for r, s_list in RASI_TO_STARS.items():
                    if matched_star in s_list: st.session_state.v_rasi = r
                generate_audio(f"{matched_star} நட்சத்திரம். ஜாதகம் வருகிறது.", "reply_star.mp3", selected_voice)
                autoplay_audio_hidden("reply_star.mp3")
            else:
                st.warning(f"Not Recognized: {text} - Try saying just the star name.")

    available_stars = RASI_TO_STARS[rasi]
    if st.session_state.v_star not in available_stars: st.session_state.v_star = available_stars[0]
    star_full = st.selectbox("நட்சத்திரம்", available_stars, index=available_stars.index(st.session_state.v_star))

    st.divider()
    dob = st.date_input("பிறந்த தேதி", min_value=datetime.date(1950, 1, 1))

# --- MAIN ---
col1, col2, col3 = st.columns([1, 6, 1])

with col2:
    if st.button("🔮 ஜாதகம் கணிக்கவும்", use_container_width=True):
        if name:
            with st.spinner("கணிக்கப்படுகிறது..."):
                time.sleep(1.0)
                planet_info = PLANET_CONTEXT[rasi]
                career = CAREER_DATA[rasi]
                rel = RELATIONSHIP_DATA["Single"]
                name_vibes = get_name_analysis(name)
                skill = HOBBY_DATA.get(star_full.split(" (")[0], "தியானம்")
                roast = FUNNY_ROASTS[rasi]
                life_path = calculate_life_path_number(dob)
                numerology = NUMEROLOGY_DATA[life_path]
                selected_quote = random.choice(QUOTES_DATA)

                html_code = f"""
                <div class="premium-card">
                <h1 class="header-gold">ஜாதக பலன் 2026</h1>
                <div class="sub-header">{name} | {rasi} | {star_full}</div>
                <hr style="border-color: #333;">
                <div class="planet-box"><strong>🪐 கிரக நிலை:</strong><br>{planet_info}</div>
                <div class="section-title">🔢 எண் கணிதம் ({life_path})</div><p class="content-text">{numerology}</p>
                <div class="section-title">👤 பெயர் பலன்</div><p class="content-text">{name_vibes}</p>
                <div class="section-title">🚀 தொழில்</div><p class="content-text">{career}</p>
                <div class="section-title">❤️ உறவு</div><p class="content-text">{rel}</p>
                <div class="section-title">🎓 கலை</div><p class="content-text">{skill}</p>
                <div class="roast-box">😜 Roast:<br>{roast}</div>
                <div class="quote-box">📜 சிந்தனை துளி:<br>"{selected_quote}"</div>
                </div>
                """
                st.markdown(html_code, unsafe_allow_html=True)
                st.balloons()
                
                st.divider()
                st.markdown("### 🔊 Listen to Your Prediction")
                
                speech_text = f"""
                என்ன {name}, உங்களுடைய இந்த 2026 எப்படி இருக்குனு பாத்துறலாமா?
                
                உங்கள் ராசி {rasi}. நட்சத்திரம் {star_full}.
                
                முதலில் கிரக நிலை. {planet_info}.
                எண் கணித பலன். {numerology}.
                பெயர் அதிர்ஷ்டம். {name_vibes}.
                தொழில் வளர்ச்சி. {career}.
                உறவுமுறை பலன். {rel}.
                நீங்கள் கற்க வேண்டிய கலை. {skill}.
                ஒரு சின்ன கிண்டல். {roast}.
                
                இன்றைய சிந்தனை. {selected_quote}.
                
                சரி, இங்க இருந்து போயிட்டு, நான் சொன்னது சரியான்னு நெட்ல போட்டு பாக்காதீங்க!
                """
                
                safe_name = "".join(c for c in name if c.isalnum()).strip()
                audio_file = f"{safe_name}_2026.mp3"
                
                if generate_audio(speech_text, audio_file, selected_voice):
                    st.audio(audio_file, format="audio/mp3")
                    with open(audio_file, "rb") as f:
                        st.download_button("⬇️ Download Audio", f, file_name=audio_file)
                else:
                    st.error("Audio generation failed")
        else:
            st.error("பெயரை உள்ளிடவும்!")
