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

// Everyday errands, offices, travel and work sites (same style as the
// "Don't Push The Button" set). Every prompt still says no children in frame.
const SCENES = [
  ["Electrician Work Truck", "pulling a coil of wire from the truck bed"],
  ["Roofing Job Site Driveway", "carrying a bundle of shingles on one shoulder"],
  ["HVAC Service Van", "unloading a filter box from the back of the van"],
  ["Landscaping Trailer", "lowering the mower ramp"],
  ["Auto Parts Store Counter", "waiting while the clerk searches the computer"],
  ["Oil Change Waiting Room", "holding a paper cup of free coffee"],
  ["Car Wash Exit Lane", "drying a side mirror with a towel"],
  ["Car Dealership Showroom", "holding a stack of financing papers"],
  ["Insurance Office Lobby", "holding a claim folder with a deep breath"],
  ["Tax Prep Office", "sliding a shoebox of receipts across the desk"],
  ["Shipping Store Notary Desk", "waiting with a pen in hand"],
  ["Passport Photo Line", "holding a passport form"],
  ["Courthouse Hallway", "holding a parking ticket envelope"],
  ["City Hall Permit Office", "holding a rolled-up set of plans"],
  ["Social Security Office Waiting Area", "holding a numbered ticket with a blank stare"],
  ["Urgent Care Waiting Room", "filling out a clipboard form"],
  ["Dentist Office Front Desk", "checking in at the counter"],
  ["Eyeglass Store Display", "trying on frames in a small mirror"],
  ["Vet Clinic Lobby", "holding a dog leash while the dog sniffs around"],
  ["Laundromat", "folding towels on the counter"],
  ["Dry Cleaner Counter", "holding a garment ticket"],
  ["Apartment Leasing Office", "holding up a set of new keys"],
  ["Storage Unit Hallway", "rolling up the unit door"],
  ["Moving Truck Rental Lot", "holding a clipboard beside a box truck"],
  ["Home Improvement Paint Counter", "waiting while the paint gets mixed"],
  ["Garden Center Checkout", "holding a flat of flowers"],
  ["Furniture Store Showroom", "testing out a recliner"],
  ["Mattress Store", "sitting up on the edge of a mattress"],
  ["Appliance Delivery Driveway", "signing for a new washing machine"],
  ["Muffler Shop Bay", "looking under a lifted car with a flashlight"],
  ["Truck Stop Diner Counter", "stirring a mug of coffee"],
  ["Toll Plaza Line", "sitting in a pickup's driver seat with the window down"],
  ["Highway Rest Stop", "stretching beside a parked car"],
  ["EV Charging Station", "leaning on the car while it charges"],
  ["Bus Stop Shelter", "checking the time on a phone"],
  ["Commuter Train Platform", "holding a coffee while waiting for the train"],
  ["Subway Car", "holding the overhead grab bar"],
  ["Rideshare Pickup Curb", "scanning passing cars with a carry-on bag"],
  ["Airport Baggage Claim", "waiting while the carousel spins"],
  ["Airport Gate During A Delay", "sitting with a neck pillow and a sigh"],
  ["Hotel Lobby Coffee Station", "pouring a cup of coffee"],
  ["Hotel Ice Machine Hallway", "holding an ice bucket"],
  ["Coworking Space Phone Booth", "talking into a headset"],
  ["Home Office Video Call", "waving at the webcam"],
  ["Open-Plan Office Standing Desk", "typing at a standing desk"],
  ["Office Copy Room", "staring at a jammed printer"],
  ["Office Kitchen Fridge", "holding a labeled lunch container"],
  ["IT Help Desk", "handing over a laptop"],
  ["Company Town Hall Meeting", "sitting in a row of chairs with arms crossed"],
  ["Corporate Team-Building Ropes Course", "clipping into a safety harness"],
  ["Warehouse Loading Dock", "pushing a hand truck stacked with boxes"],
  ["Delivery Route Doorstep", "holding a package at a front door"],
  ["Construction Site Lunch Break", "opening a lunch cooler on a tailgate"],
  ["Farm Supply Store", "carrying a bag of feed on one shoulder"],
  ["Ranch Gate", "leaning on the gate with a coffee"],
  ["Boat Ramp", "standing by the truck while backing in a trailer"],
  ["Marina Fuel Dock", "holding the fuel nozzle"],
  ["Bait Shop Counter", "holding a container of bait"],
  ["Golf Pro Shop Counter", "holding a sleeve of golf balls"],
  ["Sporting Goods Store Aisle", "testing a folding camping chair"],
  ["Bike Repair Shop", "holding a flat bike tire"],
  ["Gym Front Desk", "scanning a membership card"],
  ["Nail Salon Chair", "drying a fresh manicure"],
  ["Hair Salon Waiting Couch", "flipping through a magazine"],
  ["Pet Store Aisle", "holding a giant bag of dog food"],
  ["Liquor Store Checkout", "setting a bottle on the counter"],
  ["Warehouse Club Sample Station", "holding a tiny sample cup"],
  ["Warehouse Club Parking Lot", "loading a giant pack of paper towels into a trunk"],
  ["Coffee Shop Order Line", "waiting with a phone in hand"],
  ["Bagel Shop Counter", "holding a paper bag of bagels"],
  ["Deli Counter", "holding a take-a-number ticket"],
  ["Butcher Shop Counter", "pointing at steaks in the case"],
  ["Pizza Pickup Counter", "holding a pizza box"],
  ["Chinese Takeout Counter", "holding two takeout bags"],
  ["Barbecue Joint Line", "holding a tray of brisket"],
  ["Mall Food Court", "holding a soft pretzel"],
  ["Shoe Store Bench", "lacing up a new pair of sneakers"],
  ["Electronics Store TV Wall", "comparing two big TVs"],
  ["Phone Repair Counter", "holding up a cracked phone"],
  ["Walk-Up ATM", "taking cash from the machine"],
  ["Credit Union Loan Desk", "signing loan paperwork"],
  ["Real Estate Open House Kitchen", "holding a flyer and looking around"],
  ["HOA Meeting Clubhouse", "raising a hand to speak"],
  ["Neighborhood Mailbox Cluster", "pulling out a stack of bills"],
  ["Apartment Laundry Room", "waiting beside a dryer"],
  ["Driveway Car Wash", "hosing down a car"],
  ["Front Yard Lawn Mowing", "stopping the mower to wipe their brow"],
  ["Front Yard Leaf Raking", "leaning on a rake"],
  ["Weekend Garage Sale", "holding up an old lamp"],
  ["Recycling Center Drop-Off", "tossing a flattened box into a bin"],
  ["Salvage Yard", "holding a salvaged side mirror"],
  ["Locked-Out Car In A Parking Lot", "standing by the driver door holding a coffee"],
  ["Windshield Repair Appointment", "pointing at a chip in the windshield"],
  ["Emissions Testing Station", "waiting beside their car"],
  ["Office Vending Machine", "tapping the glass at a stuck snack"],
  ["Trade Show Booth", "handing out a brochure"],
  ["Job Interview Lobby", "waiting in a chair holding a resume folder"],
  ["Bookstore Cafe", "holding a paperback and a latte"],
  ["Parking Garage Pay Station", "feeding a ticket into the machine"],
  ["Gas Station Convenience Store", "holding a fountain drink and a hot dog"],
];

const WOMEN = [
  "wavy blond hair", "long dark braids", "freckles and red hair", "short dark pixie hair",
  "a messy bun", "big hoop earrings", "black-framed glasses", "curly auburn hair",
  "a neat ponytail", "a knit beanie", "a sleek black bob", "sun-kissed beach waves",
  "rose-gold glasses", "a high ponytail", "natural curly hair", "a headband", "thick curly hair", "a denim jacket over the shirt left open",
];
const MEN = [
  "a buzz cut", "a shaved head and goatee", "wavy brown hair", "a worn ball cap",
  "a full lumberjack beard", "a trimmed mustache", "curly black hair", "a man bun",
  "a backwards cap", "round glasses", "a bucket hat", "a short boxed beard", "tattooed forearms", "thick curly hair",
];
const OLDER_MEN = ["a salt-and-pepper beard", "close-cropped gray hair", "a trimmed white beard"]; // 40+
const OLDER_WOMEN = ["a silver-streaked bob", "reading glasses pushed up on her head"]; // 40+

const OPENERS = [
  "Okay, shirt check.", "New shirt day, and it's a good one.", "Everybody keeps asking about this shirt.",
  "I wore this on purpose today.", "Let's see who reads this one first.", "My friend gave me this and I can't stop wearing it.",
  "Be honest, is this the best shirt here?", "This shirt gets a reaction every single time.",
  "I didn't expect this many people to notice my shirt.", "Look what showed up in the mail.",
  "Wore this just to see the reactions.", "Quick close-up, because people keep squinting at it.",
  "Rate this shirt, one to ten.", "You'll want to read this one twice.", "Three people asked where I got this already.",
  "This might be my new favorite shirt.", "Just running errands and everyone is reading my shirt.", "Real talk, this shirt is the main event.",
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
  const age = 25 + Math.floor(rand() * 38); // 25-62
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
    text: "Always attach the design mockup image when generating; it is the exact T-shirt. All 100 scenes are everyday American places (errands, offices, travel, work sites) with adults only. " +
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
