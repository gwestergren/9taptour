// Static snapshot of the live home page data (Nov 26, 2026) for the design mockups.
window.D = {
  champ: { name: "Miles Nerona", first: "Miles", last: "Nerona", where: "Angel of the Winds", city: "Arlington", date: "November 22, 2026", short: "Nov 22",
    head: "AOW = International 1st-Time Champion!!!", series: 841, won: "$2,678", entries: 157, paid: "$11,912", cut: 705,
    story: [
      "Starting their 6+ hour trip from Campbell River (on Vancouver Island) heading for Angel of the Winds at 3:00AM on Sunday, Miles Nerona was determined to make his long trip worth it. With the support of his wife Amy, he was able to focus, despite just one hour of sleep, making his approach to each shot look effortless.",
      "The Neronas run the Bowling Center in Campbell River (CR Bowling), contributing to his late start after a busy Saturday night. Finding a line that worked Sunday, he was able to string three great games toward a great 841 series, leading to his first Tour Championship."],
    top: [["Miles Nerona", 841, "$2,678"], ["Samson Ibale", 803, "$1,176"], ["Greg Burdsall", 790, "$603"]] },
  next: { mon: "Nov", day: 29, dow: "Sunday", center: "Lucky Strike", city: "Lynnwood, WA", rel: "In 3 days", squads: ["11:00", "1:00", "3:00"] },
  later: [["Dec 6", "Hi Joy Bowl", "Port Orchard", ""], ["Dec 13", "Bowlero Lakewood", "Lakewood", "Year End Championship · invite only"]],
  jackpot: "$12,451", record: "$13,120", progressive: "$2,260",
  jpNote: "Start a game with seven straight strikes to win 80%.", progNote: "Shared by the top three finishers who paid in. Unclaimed money rolls over.",
  stats: [["70", "Tournaments"], ["$170,857", "Won by champions"], ["$3,782", "Biggest check", "Ty Richardson"], ["$12,451", "9-Tap Jackpot now"]],
  wall: [["Nov 22", "Miles Nerona", "Angel of the Winds", "$2,678"], ["Nov 15", "Joel Brown", "Bowlero Paradise", "$2,340"], ["Nov 1", "Steve Nolan", "Kenmore Lanes", "$2,363"],
    ["Oct 25", "Joe Fincher Jr.", "Kent Bowl", "$2,080"], ["Oct 18", "Samson Ibale", "West Seattle Bowl", "$2,577"], ["Oct 11", "Samson Ibale", "Bowlero Lanes", "$2,581"],
    ["Oct 4", "Charles Daughery Jr", "Secoma Lanes", "$2,627"], ["Sep 19", "Paul Kelly", "Bowlero Paradise", "$2,910"]],
  jps: [["Nov 22", "Troy Breeding, Charles Clarke, Leo Morzo", "Front-4", "$150 each"], ["Nov 15", "Joyce Roberts, Bill Sprinkle, Andrey Williams", "Front-4", "$150 each"],
    ["Nov 1", "Kahlee Barrio", "Golden Shot (Front-6)", "$750"], ["Oct 25", "Kim Pohlot", "Front-4", "$150"], ["Oct 18", "Teddy Drake", "Middle-7", "$750"], ["Oct 11", "Marilyn Mateo", "Middle-7", "$750"]],
  years: ["2026", "2025"], months: ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"],
  nav: ["Home", "Schedule", "Standings", "Results", "BOTY", "Rules", "About", "Videos", "Contact"],
  sponsors: ["carstar", "miller-lite", "storm", "dexter", "heritage", "command", "south-tacoma-honda", "rain-city"]
};
// Simple sample-photo placeholder (trophy silhouette), colors vary per design.
window.PIC = (bg, fg = "#1b1b1b", tag = true) => `<div style="position:relative;width:100%;height:100%">${tag ? `<span style="position:absolute;left:8px;top:8px;font:700 10px/1 sans-serif;letter-spacing:.06em;background:rgba(0,0,0,.55);color:#fff;padding:4px 6px;border-radius:3px">SAMPLE PHOTO</span>` : ""}<svg viewBox="0 0 160 140" preserveAspectRatio="xMidYMid slice" role="img" aria-label="Sample photo" style="width:100%;height:100%;display:block"><rect width="160" height="140" fill="${bg}"/><path d="M0 105 L160 92 L160 140 L0 140Z" fill="${fg}" opacity=".18"/><rect x="18" y="58" width="5" height="16" rx="2.5" fill="#f3e3c3" opacity=".6"/><rect x="27" y="56" width="5" height="18" rx="2.5" fill="#f3e3c3" opacity=".6"/><rect x="134" y="54" width="5" height="18" rx="2.5" fill="#f3e3c3" opacity=".6"/><path d="M58 22h44c0 16-8 26-22 26S58 38 58 22z" fill="#e3c286"/><rect x="77" y="47" width="6" height="8" fill="#e3c286"/><rect x="70" y="54" width="20" height="5" rx="2" fill="#e3c286"/><path d="M60 30 L72 92 L88 92 L100 30 L94 28 L84 74 L76 74 L66 28Z" fill="${fg}"/><circle cx="80" cy="76" r="11" fill="${fg}"/><path d="M58 140 Q60 90 80 86 Q100 90 102 140Z" fill="${fg}"/></svg></div>`;
window.SP = (cls = "") => D.sponsors.map(s => `<img class="${cls}" src="../img/${s}.png" alt="${s.replace(/-/g, " ")}">`).join("") +
  `<span class="${cls} txt jd">JD &amp; Co.<small>Jewelry</small></span><span class="${cls} txt fq">Fitness<i>Quest</i></span>`;
