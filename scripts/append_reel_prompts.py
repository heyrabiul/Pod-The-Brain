"""Append new reel prompts to an existing "50 NEW T-Shirt-Focused Reels" docx.

Usage: python3 -I scripts/append_reel_prompts.py IN.docx DATA.json OUT.docx

The design wording, graphic description, design hashtags, age range and
man/woman mix are read from the entries already in IN.docx. DATA.json holds
the new scenes and dialogue pieces:
  {"scenes": [[location, action], ...], "openers": [...], "details": [...],
   "followups": [...], "audio": [...], "titles": [...], "tags": [...]}
New entries use the same paragraph XML as the existing ones and continue the
numbering, so they look identical in Word.
"""
import json
import random
import re
import sys
import zipfile
from xml.sax.saxutils import escape

MEN = [
    "a buzz cut", "a worn ball cap", "black-framed glasses", "tattooed forearms", "wavy brown hair",
    "a shaved head and goatee", "a knit beanie", "a bright smile and laugh lines",
    "a flannel-style lumberjack beard", "curly black hair", "a backwards cap", "round glasses",
]
OLDER_MEN = ["a salt-and-pepper beard", "short gray hair", "a trimmed white beard", "a friendly gray mustache"]
WOMEN = [
    "a messy bun", "a neat ponytail", "freckles and red hair", "light-brown shoulder-length hair",
    "curly auburn hair", "short dark pixie hair", "round glasses", "a headband", "a knit beanie",
    "long dark braids", "big hoop earrings", "wavy blond hair", "thick curly hair", "black-framed glasses",
    "a bright smile and laugh lines",
]
OLDER_WOMEN = ["a silver bob", "short gray hair"]
ENTRY_HEAD = re.compile(r"^(\d+)\. .* — (\d+)-year-old American (man|woman) with ")


def texts(doc):
    for para in re.findall(r"<w:p[ >].*?</w:p>", doc, flags=re.S):
        yield "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para))


def unescape(s):
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def read_design(doc):
    heads, image_prompt, first_tags = [], None, None
    for t in map(unescape, texts(doc)):
        m = ENTRY_HEAD.match(t)
        if m:
            heads.append((int(m.group(1)), int(m.group(2)), m.group(3)))
        elif t.startswith("IMAGE PROMPT: ") and image_prompt is None:
            image_prompt = t
        elif t.startswith("HASHTAGS: ") and first_tags is None:
            first_tags = t[len("HASHTAGS: "):].split()
    title = unescape(re.search(r'<w:pStyle w:val="Title"/></w:pPr><w:r><w:t>([^<]*)', doc).group(1))
    name = title.split("—")[-1].strip()
    wording = re.search(r"Preserve the exact wording (“.*?”), the same typography", image_prompt).group(1)
    look = re.search(r"graphic layout \((.*?)\), and the same black T-shirt", image_prompt).group(1)
    name_words = {w.lower().strip("'") for w in re.split(r"\W+", name) if w}
    design_tags = [t for t in first_tags if t[1:].lower() in name_words] or first_tags[:1]
    men = sum(1 for h in heads if h[2] == "man")
    return {
        "title": title, "wording": wording, "look": look, "design_tags": design_tags,
        "last": max(h[0] for h in heads), "count": len(heads),
        "ages": (min(h[1] for h in heads), max(h[1] for h in heads)), "men_share": men / len(heads),
    }


PROPER = {"Christmas", "Halloween", "Thanksgiving", "Fourth", "July", "Year's", "Eve", "Super", "Bowl",
          "March", "Madness", "Labor", "Memorial", "St.", "Patrick's", "Mardi", "Gras", "Vegas", "Texas",
          "American", "Valentine's", "Friday", "Tuesday"}


def location_phrase(location):
    # "Hardware Store Aisle" -> "hardware store aisle"; keeps DMV, Fourth of July, New Year's Eve, etc.
    words = location.split()
    out = []
    for i, w in enumerate(words):
        nxt = words[i + 1] if i + 1 < len(words) else ""
        prev = words[i - 1] if i else ""
        keep = (w.isupper() or w in PROPER or (w == "New" and nxt == "Year's")
                or (w == "Day" and prev in {"Labor", "Memorial", "Valentine's", "Patrick's"}))
        out.append(w if keep else w.lower())
    return " ".join(out)


def article(phrase):
    return "an" if phrase[0].lower() in "aeiou" else "a"


def para_heading(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{escape(text)}</w:t></w:r></w:p>'


def para_label(label, text):
    return (f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{label}: </w:t></w:r>'
            f'<w:r><w:t>{escape(text)}</w:t></w:r></w:p>')


def entry(n, location, action, design, data, rng):
    man = rng.random() < design["men_share"]
    age = rng.randint(*design["ages"])
    pool = (MEN + (OLDER_MEN if age >= 45 else [])) if man else (WOMEN + (OLDER_WOMEN if age >= 45 else []))
    who = f"{age}-year-old American {'man' if man else 'woman'} with {rng.choice(pool)}"
    loc = location_phrase(location)
    wording, look = design["wording"], design["look"]
    line = f"{rng.choice(data['openers'])} {rng.choice(data['details'])}"
    extra = [t for t in data["tags"] if t.lower() not in {d.lower() for d in design["design_tags"]}]
    tags = " ".join(design["design_tags"] + rng.sample(extra, 9 - len(design["design_tags"])) + ["#Reels"])

    image = (
        "Use the attached mockup image as the EXACT T-shirt reference: the shirt in the image must be the shirt the person wears, unchanged. "
        f"Preserve the exact wording {wording}, the same typography, ink colours and graphic layout ({look}), and the same black T-shirt presentation. "
        "Do not change, rewrite, shorten, add, remove, or rearrange any wording. "
        f"Create a highly photorealistic {who} in a realistic American setting at {article(loc)} {loc}. The person is {action}. "
        "Natural age-appropriate appearance, realistic skin texture, authentic hair, natural facial expression, realistic black cotton T-shirt with visible fabric weave and natural folds. "
        "The T-shirt and its graphic are the MAIN FOCUS. Keep the entire front design clearly visible, centered, sharp and readable. "
        "Real-life commercial lifestyle photography, natural lighting, realistic background, subtle depth of field, believable American environment, authentic camera perspective. "
        "IMAGE SIZE: 9:16 VERTICAL. No text anywhere except the shirt print itself: no logo, no watermark, no QR code, no altered typography, no additional graphic."
    )
    video = (
        "Create a 10-second photorealistic 9:16 vertical USA lifestyle Reel that looks like genuine footage captured by a real person on a modern smartphone camera, not an AI-generated video. "
        "You MUST use the attached mockup image as the exact T-shirt: the person wears that exact shirt, with the identical graphic, wording, typography, ink colours and layout from the attached image, nothing redrawn or invented. "
        f"Location: {loc}. Main subject: {who} wearing the exact T-shirt from the attached mockup image. "
        f"Preserve the exact printed wording {wording}. The person is {action}. "
        f"They speak naturally and spend most of the dialogue talking specifically about the T-shirt: “{line}” {rng.choice(data['followups'])} "
        "Keep the T-shirt as the visual MAIN FOCUS for most of the 10 seconds. The camera should provide a clear front view of the complete graphic, with natural slight handheld movement and realistic autofocus. "
        "The person may gently point to, touch, straighten, or display the shirt while talking about its design, message, bold lettering, joke appeal, or gift value. "
        "Use natural American voice, realistic lip sync, authentic body movement, real cotton fabric behavior, natural wrinkles and folds, realistic lighting, and location-appropriate ambient sound. "
        f"{rng.choice(data['audio'])} Make it feel like a real lifestyle Reel filmed in America. IMAGE/VIDEO FRAME: 9:16 VERTICAL. "
        "No CGI look, no 3D render, no cartoon, no plastic fabric, no artificial smooth movement, no distorted hands, no extra fingers, no altered wording, no added text, no captions or subtitles, no logo, no watermark, no QR code."
    )
    return "".join([
        para_heading(f"{n}. {location} — {who}"),
        para_label("TITLE", rng.choice(data["titles"])),
        para_label("HASHTAGS", tags),
        para_label("IMAGE SIZE", "9:16 VERTICAL"),
        para_label("IMAGE PROMPT", image),
        para_label("VIDEO PROMPT", video),
    ])


def main(src, data_path, dst):
    data = json.load(open(data_path, encoding="utf8"))
    scenes = data["scenes"]
    assert len({s[0].lower() for s in scenes}) == len(scenes), "duplicate location in DATA.json"
    with zipfile.ZipFile(src) as zin:
        doc = zin.read("word/document.xml").decode("utf8")
        design = read_design(doc)
        rng = random.Random(design["title"])
        start = design["last"] + 1
        new = "".join(entry(start + i, loc, act, design, data, rng) for i, (loc, act) in enumerate(scenes))
        assert doc.count("<w:sectPr") == 1
        doc = doc.replace("<w:sectPr", new + "<w:sectPr", 1)
        total = design["count"] + len(scenes)
        doc = re.sub(r"(<w:pStyle w:val=\"Title\"/></w:pPr><w:r><w:t>)\d+ NEW ", rf"\g<1>{total} ", doc, count=1)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                out = doc.encode("utf8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, out)
    print(f"{design['title']}: added {len(scenes)} entries ({start}-{start + len(scenes) - 1}), tags {design['design_tags']}")


if __name__ == "__main__":
    main(*sys.argv[1:4])
