"""Append 70 new reel prompts (entries 51-120) to the Equally Whack Best Friends docx.

Usage: python3 -I scripts/append_equally_whack_prompts.py IN.docx OUT.docx

Entries are written with the same paragraph XML as the existing ones, so they
look identical in Word. Scenes do not repeat the 50 already in the file.
"""
import random
import sys
import zipfile
from xml.sax.saxutils import escape

WORDING = "“'TURTLES ARE GREEN' / 'DUCKS GO QUACK' / ★ 'WE'RE' ★ / Best Friends / 'CAUSE OUR HEADS ARE / EQUALLY WHACK!”"
LOOK = "white condensed caps, hot-pink bold script Best Friends, small white stars, big bold EQUALLY WHACK!"

# (location, what she is doing with her best friend)
SCENES = [
    ("Pottery Class Wheel", "shaping clay at the wheel next to her best friend"),
    ("Day Spa Lounge", "sipping cucumber water next to her best friend"),
    ("Taco Tuesday Restaurant", "clinking margarita glasses with her best friend"),
    ("Ice Cream Parlor", "swapping ice cream cones with her best friend"),
    ("Drive-In Movie Tailgate", "sharing a blanket with her best friend on a pickup tailgate"),
    ("Roller Skating Rink", "holding hands with her best friend while skating slowly"),
    ("Ice Skating Rink", "wobbling arm in arm with her best friend"),
    ("Escape Room Lobby", "high-fiving her best friend after escaping"),
    ("Axe Throwing Lane", "cheering with her best friend after a bullseye"),
    ("Yoga Studio Lobby", "rolling up yoga mats with her best friend"),
    ("Grocery Store Snack Aisle", "debating two bags of chips with her best friend"),
    ("Home Decor Store Aisle", "holding up matching throw pillows with her best friend"),
    ("Furniture Showroom", "flopping onto a display couch with her best friend"),
    ("Warehouse Club Sample Station", "trying tiny samples with her best friend"),
    ("Craft Store Aisle", "filling a basket with glitter and yarn with her best friend"),
    ("Candle Making Workshop", "pouring wax into jars with her best friend"),
    ("Cooking Class Kitchen", "rolling out pasta dough with her best friend"),
    ("Vineyard Tasting Room", "swirling wine glasses with her best friend"),
    ("Brewery Patio", "toasting pint glasses with her best friend"),
    ("Rooftop Bar At Sunset", "taking a selfie with her best friend"),
    ("Food Truck Park", "splitting a giant burrito with her best friend"),
    ("Pizza Place Booth", "pulling apart a cheesy slice with her best friend"),
    ("Cupcake Decorating Class", "piping frosting on cupcakes with her best friend"),
    ("Christmas Tree Lot", "carrying a tree with her best friend"),
    ("Holiday Lights Walk", "holding hot cocoa with her best friend under string lights"),
    ("Halloween Costume Shop", "holding up silly masks with her best friend"),
    ("Fourth Of July Backyard", "waving sparklers with her best friend at dusk"),
    ("New Year's Eve Party", "blowing party horns with her best friend"),
    ("Galentine's Brunch", "swapping little gift bags with her best friend"),
    ("Birthday Dinner Restaurant", "blowing out a candle on a shared slice of cake with her best friend"),
    ("Bachelorette Weekend Kitchen", "dancing around the kitchen island with her best friend"),
    ("Bridal Shower Backyard", "sticking a gift bow on her best friend's head"),
    ("Baby Shower Living Room", "holding up tiny baby socks with her best friend"),
    ("New Apartment Moving Day", "carrying a cardboard box with her best friend"),
    ("Housewarming Party", "hanging a picture frame with her best friend"),
    ("Driveway Garage Sale", "holding up a ridiculous vintage hat with her best friend"),
    ("Antique Mall Booth", "trying on old sunglasses with her best friend"),
    ("Vintage Record Store", "holding up two records with her best friend"),
    ("Sunflower Farm", "holding bunches of sunflowers with her best friend"),
    ("Strawberry Picking Field", "holding a basket of strawberries with her best friend"),
    ("Holiday Craft Fair", "holding up handmade ornaments with her best friend"),
    ("Zoo Giraffe Deck", "feeding a giraffe with her best friend"),
    ("Aquarium Turtle Tank", "pointing at a passing sea turtle with her best friend"),
    ("Duck Pond Park", "tossing seed to the ducks with her best friend"),
    ("Botanical Garden Path", "smelling roses with her best friend"),
    ("Beach Umbrella Setup", "planting a beach umbrella with her best friend"),
    ("Pool Float Party", "holding an inflatable flamingo with her best friend"),
    ("River Tubing Launch", "holding inner tubes with her best friend"),
    ("Campground Tent Site", "struggling to set up a tent with her best friend"),
    ("RV Road Trip Stop", "sitting in camp chairs beside an RV with her best friend"),
    ("Gas Station Snack Run", "holding armfuls of road-trip snacks with her best friend"),
    ("Parked Car Singalong", "singing in the front seats with her best friend"),
    ("Coffee Truck Window", "grabbing two iced coffees with her best friend"),
    ("Gym Weight Room", "spotting her best friend and cracking up"),
    ("Fun Run Finish Line", "holding finisher medals with her best friend"),
    ("Golf Cart Path", "riding in a golf cart with her best friend"),
    ("Tennis Court", "resting rackets on their shoulders with her best friend"),
    ("Line Dancing Night", "doing the same dance move as her best friend"),
    ("Comedy Club Table", "cracking up at a joke with her best friend"),
    ("Book Club Living Room", "holding up the same book as her best friend"),
    ("Trivia Night Bar", "writing answers on a slip with her best friend"),
    ("Retro Arcade", "playing skee-ball side by side with her best friend"),
    ("Laundromat", "folding laundry and gossiping with her best friend"),
    ("Dog Park", "holding two leashes with her best friend"),
    ("Self-Serve Car Wash", "dodging the spray wand with her best friend"),
    ("Tattoo Shop", "showing matching tiny wrist tattoos with her best friend"),
    ("Hotel Pool Lounge", "lounging on chairs with her best friend"),
    ("Airport Baggage Claim", "grabbing matching suitcases with her best friend"),
    ("Train Window Seat", "sharing snacks by the window with her best friend"),
    ("Snowy Cabin Kitchen", "making hot chocolate with her best friend"),
]

FEATURES = [
    "a messy bun", "a headband", "round glasses", "big hoop earrings", "curly auburn hair",
    "long dark braids", "freckles and red hair", "thick curly hair", "a neat ponytail",
    "short dark pixie hair", "wavy blond hair", "light-brown shoulder-length hair",
    "black-framed glasses", "a bright smile and laugh lines", "a high ponytail", "a denim jacket",
    "sun-kissed beach waves", "a sleek black bob", "a scrunchie in her hair",
]
OLDER_FEATURES = ["a silver bob", "short gray hair"]  # 45+

OPENERS = [
    "Me and my best friend in one shirt.", "Bought this for my best friend, kept one for me.",
    "If you know, you know.", "This shirt is basically our friendship.", "My best friend made me wear this.",
    "Finally, a shirt that explains us.", "Tag the friend who's just as whack as you.",
    "Every best friend duo needs this shirt.", "Shirt check, bestie edition.", "We've been friends way too long for this shirt.",
    "Read it slowly, it gets better.", "This is the official shirt of our friendship.",
    "People keep reading this out loud to us.", "The most accurate shirt I own.",
]
DETAILS = [
    "The hot-pink Best Friends script really pops on the black.", "The white letters are crisp and easy to read.",
    "The little stars are such a cute touch.", "Equally whack, big and bold, right where it belongs.",
    "Soft cotton, and the print feels thick.", "Perfect birthday gift for your bestie.",
    "It reads like a little poem, and the ending gets everyone.", "The layout is clean, top to bottom.",
    "It still looks this good after a lot of washes.", "Get one for your best friend before her birthday.",
    "That pink against the black is my favorite part.", "The turtles and ducks line gets a laugh every time.",
]
FOLLOWUPS = [
    "Then she glances down at the shirt, laughs, and looks back at the camera.",
    "Then she points to the hot-pink Best Friends line and grins.",
    "Then she smooths the shirt flat with both hands.",
    "Then she points to each line one by one while her best friend nods.",
    "Then she bumps shoulders with her best friend and shows the shirt again.",
    "Then she holds the hem out with both hands to show the whole graphic.",
    "Then she points at her best friend, then at the words EQUALLY WHACK.",
    "Then she winks at the camera and taps the print.",
]
AUDIO = [
    "Soft, genuine giggles and light laughter from friends just off-camera can be heard, reacting to the shirt.",
    "Her best friend laughs on camera and says \"So true,\" mixed with natural ambient sound.",
    "Add warm, natural background laughter and a few amused giggles from friends nearby, as if they just read the shirt.",
    "Include unscripted giggling and a short burst of laughter off-camera, mixed with the natural ambient sound.",
]
TITLES = [
    "Tag Your Equally Whack Bestie", "Best Friend Gift Idea 😂", "Official Best Friend Shirt",
    "This Is Literally Us 😂", "Gift Idea: Equally Whack Best Friends", "Bestie Shirt Check 💕",
    "When Your Heads Are Equally Whack", "The Funniest Best Friend Shirt", "Send This To Your Best Friend",
    "Turtles Are Green, Ducks Go Quack 😂",
]
TAGS = [
    "#BestFriends", "#BFF", "#BestieGift", "#FriendGroup", "#GirlsTrip", "#FunnyGift", "#GiftIdea",
    "#FunnyTShirt", "#ShirtCheck", "#BestieGoals", "#FriendshipGoals", "#GalPals", "#BirthdayGift",
]


# Words that stay capitalized when the location is used mid-sentence.
PROPER = {"Taco", "Tuesday", "Christmas", "Halloween", "Fourth", "July", "New", "Year's", "Eve", "Galentine's", "RV"}


def para_heading(text):
    return f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{escape(text)}</w:t></w:r></w:p>'


def para_label(label, text):
    return (f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{label}: </w:t></w:r>'
            f'<w:r><w:t>{escape(text)}</w:t></w:r></w:p>')


def entry(n, location, action, rng):
    age = rng.randint(20, 58)
    feature = rng.choice(FEATURES + (OLDER_FEATURES if age >= 45 else []))
    who = f"{age}-year-old American woman with {feature}"
    loc = " ".join(w if w in PROPER else w.lower() for w in location.split())
    friend = "Her best friend is beside her in a plain casual top, so this shirt is the only printed graphic in the frame."
    line = f"{rng.choice(OPENERS)} {rng.choice(DETAILS)}"
    tags = " ".join(["#Equally", "#Best"] + rng.sample(TAGS, 7) + ["#Reels"])

    image = (
        "Use the attached mockup image as the EXACT T-shirt reference: the shirt in the image must be the shirt the person wears, unchanged. "
        f"Preserve the exact wording {WORDING}, the same typography, ink colours and graphic layout ({LOOK}), and the same black T-shirt presentation. "
        "Do not change, rewrite, shorten, add, remove, or rearrange any wording. "
        f"Create a highly photorealistic {who} in a realistic American setting at a {loc}. The person is {action}. {friend} "
        "Natural age-appropriate appearance, realistic skin texture, authentic hair, natural facial expression, realistic black cotton T-shirt with visible fabric weave and natural folds. "
        "The T-shirt and its graphic are the MAIN FOCUS. Keep the entire front design clearly visible, centered, sharp and readable. "
        "Real-life commercial lifestyle photography, natural lighting, realistic background, subtle depth of field, believable American environment, authentic camera perspective. "
        "IMAGE SIZE: 9:16 VERTICAL. No text anywhere except the shirt print itself: no logo, no watermark, no QR code, no altered typography, no additional graphic."
    )
    video = (
        "Create a 10-second photorealistic 9:16 vertical USA lifestyle Reel that looks like genuine footage captured by a real person on a modern smartphone camera, not an AI-generated video. "
        "You MUST use the attached mockup image as the exact T-shirt: the person wears that exact shirt, with the identical graphic, wording, typography, ink colours and layout from the attached image, nothing redrawn or invented. "
        f"Location: {loc}. Main subject: {who} wearing the exact T-shirt from the attached mockup image. "
        f"Preserve the exact printed wording {WORDING}. The person is {action}. {friend} "
        f"She speaks naturally and spends most of the dialogue talking specifically about the T-shirt: “{line}” {rng.choice(FOLLOWUPS)} "
        "Keep the T-shirt as the visual MAIN FOCUS for most of the 10 seconds. The camera should provide a clear front view of the complete graphic, with natural slight handheld movement and realistic autofocus. "
        "She may gently point to, touch, straighten, or display the shirt while talking about its design, message, bold lettering, joke appeal, or gift value. "
        "Use natural American voice, realistic lip sync, authentic body movement, real cotton fabric behavior, natural wrinkles and folds, realistic lighting, and location-appropriate ambient sound. "
        f"{rng.choice(AUDIO)} Make it feel like a real lifestyle Reel filmed in America. IMAGE/VIDEO FRAME: 9:16 VERTICAL. "
        "No CGI look, no 3D render, no cartoon, no plastic fabric, no artificial smooth movement, no distorted hands, no extra fingers, no altered wording, no added text, no captions or subtitles, no logo, no watermark, no QR code."
    )
    return "".join([
        para_heading(f"{n}. {location} — {who}"),
        para_label("TITLE", TITLES[n % len(TITLES)]),
        para_label("HASHTAGS", tags),
        para_label("IMAGE SIZE", "9:16 VERTICAL"),
        para_label("IMAGE PROMPT", image),
        para_label("VIDEO PROMPT", video),
    ])


def main(src, dst):
    assert len(SCENES) == 70 and len({s[0] for s in SCENES}) == 70
    rng = random.Random(20261007)
    with zipfile.ZipFile(src) as zin:
        doc = zin.read("word/document.xml").decode("utf8")
        new = "".join(entry(51 + i, loc, act, rng) for i, (loc, act) in enumerate(SCENES))
        doc = doc.replace("<w:sectPr", new + "<w:sectPr", 1)
        doc = doc.replace("50 NEW T-Shirt-Focused Reels", "120 T-Shirt-Focused Reels", 1)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = doc.encode("utf8") if item.filename == "word/document.xml" else zin.read(item.filename)
                zout.writestr(item, data)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
