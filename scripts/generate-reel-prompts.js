// Generates 100 Muse AI image + video Reel prompts (9:16) for one T-shirt design.
// Usage: node scripts/generate-reel-prompts.js [output.docx]
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, HeadingLevel } = require("docx");

const DESIGN = {
  name: "Anal Bead Tug O War Champion",
  // Visual description only. The wording itself comes from the attached mockup,
  // so the prompt text does not repeat the slogan (see README "Why Muse AI refuses").
  look: "stacked white retro bubble-script lettering with a swash tail, a hot-pink underlined bottom word, on a black T-shirt",
  tag: "#Champion",
};

// Adult-only settings. No kids, schools, family events, uniforms or workplaces.
const SCENES = [
  ["Karaoke Bar Stage", "holding a microphone and taking a little bow"],
  ["Axe-Throwing Lounge", "holding up a scorecard after a round"],
  ["Rooftop Bar At Sunset", "raising a cocktail glass"],
  ["Brewery Taproom", "holding a flight of beer samplers"],
  ["Lake Pontoon Boat", "sitting at the rail with sunglasses pushed up"],
  ["Darts League Night", "holding three darts and smirking"],
  ["Backyard BBQ Grill", "flipping burgers with a spatula"],
  ["Casino Night House Party", "stacking poker chips"],
  ["Adult Mini-Golf Night Out", "holding a putter and celebrating a hole-in-one"],
  ["Retro Arcade Bar", "standing next to a pinball machine with arms raised"],
  ["Escape Room Lobby", "holding an old brass key prop and cheering"],
  ["Wine Tasting Patio", "swirling a glass of red wine"],
  ["Music Festival Campground", "holding a camp mug outside a tent"],
  ["Country Bar Line-Dance Floor", "tipping a cowboy hat"],
  ["Tiki Bar", "holding a drink served in a pineapple"],
  ["Golf Clubhouse Patio", "resting a golf club on one shoulder"],
  ["Beach Boardwalk At Dusk", "holding a waffle cone"],
  ["Apartment Balcony With Friends", "leaning on the railing, laughing"],
  ["Friendsgiving Dinner Table", "carving a roast turkey"],
  ["Halloween House Party", "holding a carved pumpkin"],
  ["Fourth Of July Backyard Party", "holding a lit sparkler at night"],
  ["St. Patrick's Day Pub", "holding a pint of green beer"],
  ["Cinco De Mayo Patio Party", "holding a basket of chips and salsa"],
  ["Holiday Sweater Party", "holding a mug of eggnog"],
  ["40th Birthday Backyard Party", "blowing a party horn"],
  ["Retirement Party Private Dining Room", "raising a toast"],
  ["Divorce Party Lounge", "popping a bottle of champagne"],
  ["Engagement Party Patio", "clinking glasses with friends"],
  ["Lake Community Golf Cart", "driving slowly and waving"],
  ["Ski Lodge Apres-Ski Bar", "holding a mug of hot cocoa"],
  ["Cabin Deck Next To A Hot Tub", "holding a rolled towel"],
  ["Fishing Dock At Dawn", "holding a fishing rod"],
  ["Hunting Cabin Porch", "sipping coffee from a tin mug"],
  ["Rodeo Grandstand", "waving a cowboy hat"],
  ["Racetrack Infield Tailgate", "pointing toward the track with a cold drink"],
  ["Baseball Stadium Bleachers", "holding a hot dog"],
  ["Hockey Sports Bar", "banging the table after a goal"],
  ["Fantasy Football Draft Party", "holding a marker by the draft board"],
  ["Garage Poker Night", "fanning a hand of playing cards"],
  ["Speakeasy Cocktail Bar", "receiving a smoked cocktail"],
  ["Food Truck Rally", "holding a paper tray of tacos"],
  ["Oyster Bar", "holding up an oyster shell"],
  ["Backyard Crawfish Boil", "holding up a crawfish"],
  ["Chili Cook-Off Tent", "holding a ladle over a pot"],
  ["Hot Wing Challenge At A Sports Bar", "fanning their mouth"],
  ["Brewery Bingo Night", "dabbing a bingo card"],
  ["Dive Bar Jukebox", "picking a song on the jukebox"],
  ["Honky-Tonk Mechanical Bull Ring", "fixing their hat beside the mechanical bull"],
  ["Paintball Field Staging Area", "holding a paintball mask"],
  ["Go-Kart Track", "holding a racing helmet under one arm"],
  ["Disc Golf Course", "holding a flying disc"],
  ["Pickleball Court", "holding a paddle after a winning point"],
  ["Tennis Club Lounge", "resting a racket on their shoulder"],
  ["CrossFit Gym", "setting down a kettlebell"],
  ["Beer Run 5K Finish Area", "holding a finisher medal and a beer"],
  ["Hiking Trail Overlook", "holding trekking poles"],
  ["Kayak Launch", "holding a kayak paddle"],
  ["River Tubing Launch", "holding an inner tube"],
  ["RV Campground", "sitting in a camping chair beside an RV"],
  ["Glamping Tent", "holding a camping lantern"],
  ["Vineyard Picnic", "opening a picnic basket"],
  ["Orchard Cidery", "holding a glass of hard cider"],
  ["Distillery Tasting Room", "holding a whiskey tasting glass"],
  ["Whiskey Bar Leather Booth", "holding a rocks glass"],
  ["Jazz Club", "snapping fingers to the music"],
  ["Comedy Club Table", "laughing with a drink in hand"],
  ["Outdoor Amphitheater Lawn", "with both arms up during a concert"],
  ["Drive-In Movie Truck Bed", "holding a popcorn bucket"],
  ["Record Store Aisle", "flipping through vinyl records"],
  ["Tattoo Shop Waiting Area", "showing a freshly bandaged forearm tattoo"],
  ["Barbershop Chair", "checking a fresh haircut in the mirror"],
  ["Classic Car Show", "leaning on a classic muscle car"],
  ["Motorcycle Rally", "sitting on a parked motorcycle with a helmet in hand"],
  ["Home Garage Workshop", "wiping their hands with a shop rag"],
  ["Basement Home Bar", "pouring a drink behind the home bar"],
  ["Man Cave Sofa", "holding a TV remote with feet up"],
  ["Girls' Night Living Room", "holding a glass of rose on the couch"],
  ["Neon Party Bus", "dancing while holding the grab rail"],
  ["Resort Pool Cabana", "relaxing under a cabana with sunglasses"],
  ["Cruise Ship Deck Bar", "holding a frozen drink"],
  ["Beach Bar Deck", "holding a bottle of beer"],
  ["Hotel Rooftop Pool", "sitting on a lounge chair"],
  ["Scenic Road Trip Pull-Off", "leaning on a convertible"],
  ["Backyard Fire Pit", "roasting a marshmallow"],
  ["String-Light Patio Dinner Party", "serving a charcuterie board"],
  ["Pizza Night Kitchen", "tossing pizza dough"],
  ["Bottomless Brunch Patio", "holding a mimosa"],
  ["Late-Night Diner Booth", "holding a cup of coffee"],
  ["Lazy Sunday Couch", "holding coffee with sunglasses still on"],
  ["New Apartment Moving Day", "carrying a cardboard box"],
  ["Housewarming Party", "holding a potted plant gift"],
  ["Friend's Birthday Dinner", "pulling tissue paper out of a gift bag"],
  ["White Elephant Gift Exchange", "holding up an unwrapped gift box"],
  ["Anti-Valentine's Party", "holding a heart-shaped cookie"],
  ["Thanksgiving Backyard Football", "tossing a football"],
  ["Golf Simulator Lounge", "holding a driver after a big swing"],
  ["Board Game Cafe Adult Night", "rolling dice"],
  ["Lakefront Adirondack Chairs At Sunset", "holding a beer can"],
  ["Backyard Hammock", "sitting up in a hammock"],
  ["Rage Room", "holding a bat with safety goggles on"],
];

const WOMEN = [
  "wavy blond hair", "long dark braids", "freckles and red hair", "short dark pixie hair",
  "a messy bun", "big hoop earrings", "black-framed glasses", "curly auburn hair",
  "a neat ponytail", "a knit beanie", "a sleek black bob", "sun-kissed beach waves",
  "rose-gold glasses", "a high ponytail", "natural curly hair", "a denim jacket over the shirt left open",
];
const MEN = [
  "a buzz cut", "a shaved head and goatee", "wavy brown hair", "a worn ball cap",
  "a full lumberjack beard", "a trimmed mustache", "curly black hair", "a man bun",
  "a backwards cap", "round glasses", "a bucket hat", "a short boxed beard",
];
const OLDER_MEN = ["a salt-and-pepper beard", "close-cropped gray hair"]; // 40+
const OLDER_WOMEN = ["a silver-streaked bob", "reading glasses pushed up on her head"]; // 40+

const OPENERS = [
  "Okay, shirt check.", "New shirt day, and it's a good one.", "Everybody keeps asking about this shirt.",
  "I wore this on purpose today.", "Let's see who reads this one first.", "My friend gave me this and I can't stop wearing it.",
  "Be honest, is this the best shirt here?", "This shirt gets a reaction every single time.",
  "I didn't expect this many people to notice my shirt.", "Look what showed up in the mail.",
  "Wore this just to see the reactions.", "Quick close-up, because people keep squinting at it.",
  "Rate this shirt, one to ten.", "You'll want to read this one twice.", "Three people asked where I got this already.",
  "This might be my new favorite shirt.", "Stopped by the bar and got stopped for the shirt.", "Real talk, this shirt is the main event.",
  "Somebody had to wear it.", "I knew this shirt would start something.",
];
const DETAILS = [
  "That retro script lettering is so clean.", "The white print really pops on the black cotton.",
  "That hot-pink underline is the best part.", "It's soft, and the print doesn't crack.",
  "The bubble letters look straight out of the seventies.", "Perfect gift for the friend with the worst sense of humor.",
  "The swash on that tail is a nice touch.", "It still looks this sharp after a bunch of washes.",
  "The layout reads from across the room.", "Comfortable enough to wear all weekend.",
  "Funniest gift I've gotten all year.", "The fit is great and the print is crisp.",
  "That pink against the black is the whole vibe.", "You read it top to bottom and then it hits you.",
  "It's the lettering that sells it.", "Bold print, soft fabric, zero regrets.",
  "Every line is lined up perfectly.", "Get one for your group chat's funniest person.",
  "It looks even better in person.", "The contrast is ridiculous, in a good way.",
];
const FOLLOWUPS = [
  "Then they take a step back so the whole front is in frame.",
  "Then they straighten the shirt with both hands and face the camera.",
  "Then they point to each line of the print one by one.",
  "Then they cover their mouth, laugh, and show the shirt again.",
  "Then they turn slightly left and right to show the shirt from a natural angle.",
  "Then they shrug, grin, and tug the hem so the print sits flat.",
  "Then they gesture down at the shirt with both hands like a reveal.",
  "Then they look off-camera at a friend, laugh, and look back at the lens.",
];
const AUDIO = [
  "Include unscripted background giggling and a short burst of laughter from adult friends off-camera, mixed with the natural ambient sound.",
  "Background audio has natural chuckles from nearby adults reacting to the shirt, real and unpolished, not a laugh track.",
  "Add warm, natural background laughter from friends nearby, as if they just read the shirt.",
  "Mix in a friend off-camera saying \"No way\" and laughing, plus natural ambient sound.",
];
const TITLES = [
  "People Can't Stop Reading This", "Shirt Check 😂", "Read It Twice 😂", "The Gift Nobody Saw Coming",
  "Wore It Out And Everyone Noticed", "My Group Chat Needs This Shirt", "Best Gag Gift Of The Year",
  "Champion Energy Only 🏆", "Rate This Shirt 1–10", "Funniest Shirt In The Room", `${DESIGN.name} 😂`,
  "Wait For The Reactions 😂",
];
const TAGS = [
  "#FunnyTShirt", "#FunnyGift", "#GiftIdea", "#AdultHumor", "#SarcasticShirt", "#PartyShirt", "#FunnyShirts",
  "#TShirtLovers", "#ShirtCheck", "#FriendGroup", "#GagGift", "#HumorShirt", "#SarcasmLovers", "#GameDay",
];

// Small seeded PRNG so the output is stable between runs.
let seed = 20261007;
const rand = () => ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648);
const pick = (arr) => arr[Math.floor(rand() * arr.length)];
const shuffled = (arr) => arr.map((v) => [rand(), v]).sort((a, b) => a[0] - b[0]).map((x) => x[1]);

function person(i) {
  const woman = i % 2 === 0;
  const age = 25 + Math.floor(rand() * 31); // 25-55
  const pool = woman ? WOMEN.concat(age >= 40 ? OLDER_WOMEN : []) : MEN.concat(age >= 40 ? OLDER_MEN : []);
  return { age, gender: woman ? "woman" : "man", feature: pick(pool) };
}

function entry(i, [location, action]) {
  const p = person(i);
  const who = `${p.age}-year-old American ${p.gender} with ${p.feature}`;
  const line = `${OPENERS[i % OPENERS.length]} ${DETAILS[(i * 7 + 3) % DETAILS.length]}`;
  const tags = [DESIGN.tag, ...shuffled(TAGS).slice(0, 8), "#Reels"].join(" ");

  const image =
    `Use the attached mockup image as the EXACT T-shirt: the person wears that shirt unchanged, with the same wording, typography, ink colours and layout (${DESIGN.look}). ` +
    `Do not redraw, rewrite, add, remove or rearrange any lettering; copy it exactly from the attached image. ` +
    `Highly photorealistic ${who} in the USA, location: ${location}, ${action}. Adult-only setting: everyone in frame is an adult, no children. ` +
    `Natural age-appropriate appearance, realistic skin texture, natural expression, real black cotton T-shirt with visible fabric weave and natural folds. ` +
    `The T-shirt graphic is the MAIN FOCUS: the entire front print fully visible, centered, sharp and readable. ` +
    `Real-life commercial lifestyle photography, natural lighting, realistic background, subtle depth of field, authentic smartphone perspective. ` +
    `IMAGE SIZE: 9:16 VERTICAL. No text anywhere except the shirt print itself: no captions, no logo, no watermark, no QR code.`;

  const video =
    `Create a 10-second photorealistic 9:16 vertical USA lifestyle Reel that looks like real smartphone footage, not AI. ` +
    `Use the attached mockup image as the EXACT T-shirt: identical wording, typography, ink colours and layout, nothing redrawn or invented. ` +
    `Location: ${location}. Adult-only setting, no children in frame. ` +
    `Main subject: ${who}, wearing the exact T-shirt from the attached image, ${action}. ` +
    `They speak naturally about the shirt's look, not reading the print aloud: "${line}" ${FOLLOWUPS[(i * 3) % FOLLOWUPS.length]} ` +
    `Keep the T-shirt the visual MAIN FOCUS for most of the 10 seconds, with a clear front view of the complete graphic, slight natural handheld movement and realistic autofocus. ` +
    `Natural American voice, realistic lip sync, real cotton fabric movement, realistic lighting and location-appropriate ambient sound. ${AUDIO[i % AUDIO.length]} ` +
    `IMAGE/VIDEO FRAME: 9:16 VERTICAL. No CGI look, no cartoon, no plastic fabric, no distorted hands, no extra fingers, no altered wording, no added text, no captions or subtitles, no logo, no watermark.`;

  return { heading: `${i + 1}. ${location} — ${who}`, title: TITLES[(i * 5) % TITLES.length], tags, image, video };
}

if (SCENES.length !== 100) throw new Error(`expected 100 scenes, got ${SCENES.length}`);
if (new Set(SCENES.map((s) => s[0])).size !== 100) throw new Error("duplicate scene");
const entries = SCENES.map((s, i) => entry(i, s));

const label = (k, v) => new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: `${k}: `, bold: true }), new TextRun(v)] });
const body = [
  new Paragraph({ heading: HeadingLevel.HEADING_1, text: `100 NEW T-Shirt-Focused Reels — 9:16 Vertical — ${DESIGN.name}` }),
  new Paragraph({
    spacing: { after: 200 },
    text: "Always attach the design mockup image when generating; it is the exact T-shirt. All 100 scenes are adult-only and different from the first 50. " +
      "The prompts describe the design visually instead of quoting the slogan, and the dialogue talks about the shirt without reading it out.",
  }),
];
for (const e of entries) {
  body.push(
    new Paragraph({ heading: HeadingLevel.HEADING_2, text: e.heading, spacing: { before: 240 } }),
    label("TITLE", e.title),
    label("HASHTAGS", e.tags),
    label("IMAGE SIZE", "9:16 VERTICAL"),
    label("IMAGE PROMPT", e.image),
    label("VIDEO PROMPT", e.video),
  );
}

const out = process.argv[2] || path.join(__dirname, "..", "prompts", "Anal_Bead_Tug_O_War_Champion_100_Reels.docx");
fs.mkdirSync(path.dirname(out), { recursive: true });
const doc = new Document({
  styles: { default: { document: { run: { font: "Calibri", size: 21 } } } },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } }, children: body }],
});
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(out, buf);
  console.log(`wrote ${entries.length} entries to ${out}`);
});
