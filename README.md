<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>TPxi Software — 58 free tools for TouchPoint®</title>
<meta name="description" content="A browsable index of 58 free, open-source TouchPoint® tools built by Ben Swaby. Pick a tool to see what it does, then install it from GitHub." />
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700&family=Source+Serif+4:opsz,wght@8..60,400;8..60,600&display=swap" rel="stylesheet">
<style>
  :root{
    --paper:#f6f8fb;
    --card:#ffffff;
    --ink:#12253a;
    --ink-soft:#4a6076;
    --ink-faint:#8296aa;
    --rule:#dde5ee;
    --blue:#2563eb;
    --violet:#7c3aed;
    --gold:#b07d06;
    --focus:#2563eb;
    --shadow:0 1px 2px rgba(18,37,58,.05), 0 10px 30px rgba(18,37,58,.06);
  }
  *{box-sizing:border-box}
  html{-webkit-text-size-adjust:100%}
  body{
    margin:0;
    background:var(--paper);
    color:var(--ink);
    font-family:"Archivo",-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
    font-size:16px;
    line-height:1.5;
  }
  a{color:var(--blue)}
  .wrap{max-width:1320px;margin:0 auto;padding:0 24px}

  /* ---------- masthead ---------- */
  .masthead{border-bottom:1px solid var(--rule);background:var(--card)}
  .masthead .wrap{display:flex;align-items:center;gap:16px;height:68px}
  .brand{display:flex;align-items:center;gap:11px;text-decoration:none;color:inherit}
  .brand .name{font-weight:700;letter-spacing:-.015em;font-size:17px}
  .brand .tag{color:var(--ink-faint);font-size:13.5px;margin-left:2px;white-space:nowrap}
  .masthead nav{margin-left:auto;display:flex;align-items:center;gap:22px;font-size:14.5px}
  .masthead nav a{color:var(--ink-soft);text-decoration:none}
  .masthead nav a:hover{color:var(--ink)}
  .support-link{
    color:var(--ink)!important;font-weight:600;
    border:1px solid var(--rule);border-radius:999px;padding:6px 14px;
  }
  .support-link:hover{border-color:var(--blue);color:var(--blue)!important}

  /* ---------- hero / search ---------- */
  .lede{padding:44px 0 26px}
  .lede h1{
    font-size:clamp(30px,4.4vw,46px);
    line-height:1.06;letter-spacing:-.03em;font-weight:700;margin:0 0 12px;max-width:16ch;
  }
  .lede p{
    font-family:"Source Serif 4",Georgia,serif;
    font-size:18px;line-height:1.55;color:var(--ink-soft);
    max-width:62ch;margin:0;
  }
  .controls{
    position:sticky;top:0;z-index:20;
    background:var(--paper);
    padding:14px 0 12px;
    border-bottom:1px solid var(--rule);
  }
  .searchrow{display:flex;gap:12px;align-items:center;flex-wrap:wrap}
  .field{position:relative;flex:1 1 320px;min-width:0}
  .field svg{position:absolute;left:14px;top:50%;transform:translateY(-50%);color:var(--ink-faint)}
  #q{
    width:100%;height:46px;border:1px solid var(--rule);border-radius:10px;
    background:var(--card);padding:0 74px 0 42px;
    font:inherit;font-size:15.5px;color:var(--ink);
  }
  #q::placeholder{color:var(--ink-faint)}
  #q:focus{outline:2px solid var(--focus);outline-offset:1px;border-color:transparent}
  .slash{
    position:absolute;right:12px;top:50%;transform:translateY(-50%);
    border:1px solid var(--rule);border-radius:6px;padding:2px 7px;
    font-size:12px;color:var(--ink-faint);background:var(--paper);
  }
  #count{font-size:14px;color:var(--ink-soft);white-space:nowrap}
  .chips{display:flex;gap:7px;flex-wrap:wrap;margin-top:11px}
  .chip{
    font:inherit;font-size:13.5px;color:var(--ink-soft);
    background:var(--card);border:1px solid var(--rule);border-radius:999px;
    padding:5px 12px;cursor:pointer;
  }
  .chip:hover{color:var(--ink);border-color:var(--ink-faint)}
  .chip[aria-pressed="true"]{background:var(--ink);border-color:var(--ink);color:#fff}
  .chip:focus-visible{outline:2px solid var(--focus);outline-offset:2px}

  /* ---------- layout ---------- */
  .board{display:grid;grid-template-columns:1fr;gap:32px;padding:28px 0 64px;align-items:start}
  @media (min-width:1060px){ .board{grid-template-columns:minmax(0,1fr) 372px;gap:40px} }

  .cat{margin:0 0 34px}
  .cat-head{display:flex;align-items:baseline;gap:10px;border-bottom:2px solid var(--ink);padding-bottom:7px;margin-bottom:10px}
  .cat-head h2{font-size:17.5px;letter-spacing:-.012em;margin:0;font-weight:700}
  .cat-head .n{margin-left:auto;font-size:14px;color:var(--ink-faint);font-variant-numeric:tabular-nums}
  .list{display:grid;grid-template-columns:repeat(auto-fill,minmax(258px,1fr));gap:0 26px}
  .tool{
    display:flex;align-items:center;gap:10px;width:100%;text-align:left;
    background:none;border:0;border-bottom:1px solid var(--rule);
    padding:9px 4px;font:inherit;font-size:15px;color:var(--ink);cursor:pointer;
  }
  .tool:hover{color:var(--blue)}
  .tool:focus-visible{outline:2px solid var(--focus);outline-offset:-2px;border-radius:4px}
  .tool .dot{width:6px;height:6px;border-radius:50%;background:var(--rule);flex:none}
  .tool.fav .dot{background:var(--gold)}
  .tool .label{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  .tool .new{
    font-size:11px;color:var(--violet);border:1px solid #e2d6fb;border-radius:4px;
    padding:1px 5px;flex:none;
  }
  .tool[aria-current="true"]{color:var(--blue);font-weight:600}
  .tool[aria-current="true"] .dot{
    width:3px;height:16px;border-radius:2px;
    background:linear-gradient(180deg,var(--blue),var(--violet));
  }
  .empty{padding:40px 4px;color:var(--ink-soft);font-family:"Source Serif 4",Georgia,serif;font-size:17px}

  /* ---------- detail panel ---------- */
  .panel{
    background:var(--card);border:1px solid var(--rule);border-radius:14px;
    box-shadow:var(--shadow);padding:24px;
  }
  @media (min-width:1060px){ .panel{position:sticky;top:96px;max-height:calc(100vh - 120px);overflow:auto} }
  .panel .kicker{font-size:13.5px;color:var(--ink-faint);margin:0 0 6px}
  .panel h3{font-size:23px;line-height:1.15;letter-spacing:-.02em;margin:0 0 12px;font-weight:700}
  .panel .body{font-family:"Source Serif 4",Georgia,serif;font-size:17px;line-height:1.6;color:var(--ink-soft);margin:0 0 20px}
  .panel .marks{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 18px}
  .mark{font-size:12.5px;border-radius:5px;padding:3px 8px;border:1px solid var(--rule);color:var(--ink-soft)}
  .mark.fav{color:var(--gold);border-color:#eadfc0;background:#fdf9ef}
  .mark.new{color:var(--violet);border-color:#e2d6fb;background:#faf7ff}
  .actions{display:flex;flex-direction:column;gap:9px}
  .btn{
    display:inline-flex;align-items:center;justify-content:center;gap:8px;
    height:44px;border-radius:9px;text-decoration:none;font-size:15px;font-weight:600;
    border:1px solid transparent;cursor:pointer;font-family:inherit;
  }
  .btn-primary{background:var(--ink);color:#fff}
  .btn-primary:hover{background:#0b1a29}
  .btn-quiet{background:var(--card);border-color:var(--rule);color:var(--ink)}
  .btn-quiet:hover{border-color:var(--ink-faint)}
  .btn:focus-visible{outline:2px solid var(--focus);outline-offset:2px}
  .panel .hint{font-size:13.5px;color:var(--ink-faint);margin:16px 0 0;line-height:1.5}
  .panel .hint a{color:var(--ink-soft)}
  #closepanel{display:none}

  @media (max-width:1059px){
    .panel{
      position:fixed;left:0;right:0;bottom:0;z-index:40;
      border-radius:16px 16px 0 0;border-bottom:0;
      max-height:78vh;overflow:auto;
      box-shadow:0 -8px 40px rgba(18,37,58,.18);
      transform:translateY(101%);transition:transform .22s ease;
      padding-bottom:28px;
    }
    .panel.open{transform:translateY(0)}
    .panel.idle{display:none}
    #closepanel{
      display:block;position:absolute;top:14px;right:14px;
      background:none;border:0;font-size:22px;line-height:1;color:var(--ink-faint);cursor:pointer;padding:4px 8px;
    }
  }
  @media (prefers-reduced-motion:reduce){ .panel{transition:none} }

  /* ---------- support ---------- */
  .support{border-top:1px solid var(--rule);background:var(--card)}
  .support .wrap{padding:52px 24px}
  .support-grid{display:grid;grid-template-columns:1fr;gap:28px;max-width:960px}
  @media (min-width:820px){ .support-grid{grid-template-columns:1.15fr .85fr;gap:56px} }
  .support h2{font-size:26px;letter-spacing:-.02em;margin:0 0 12px;font-weight:700}
  .support p{font-family:"Source Serif 4",Georgia,serif;font-size:17.5px;line-height:1.6;color:var(--ink-soft);margin:0 0 14px;max-width:60ch}
  .give{border:1px solid var(--rule);border-radius:14px;padding:22px;background:var(--paper)}
  .give h3{margin:0 0 8px;font-size:17px}
  .give p{font-size:15.5px;margin:0 0 16px}
  .give .btn{width:100%}
  .give .amounts{display:flex;gap:8px;margin:0 0 12px}
  .amt{
    flex:1;text-align:center;text-decoration:none;font-size:15px;font-weight:600;color:var(--ink);
    border:1px solid var(--rule);border-radius:9px;padding:10px 0;background:var(--card);
  }
  .amt:hover{border-color:var(--blue);color:var(--blue)}
  .products{display:flex;flex-wrap:wrap;gap:10px;margin-top:18px}
  .product{
    flex:1 1 220px;border:1px solid var(--rule);border-radius:11px;padding:14px 16px;
    text-decoration:none;color:inherit;
  }
  .product:hover{border-color:var(--ink-faint)}
  .product strong{display:block;font-size:15.5px;margin-bottom:3px}
  .product span{font-size:14px;color:var(--ink-soft)}

  footer{border-top:1px solid var(--rule);padding:26px 0 40px;font-size:13px;color:var(--ink-faint)}
  footer .wrap{display:flex;flex-wrap:wrap;gap:14px 24px;align-items:center}
  footer a{color:var(--ink-soft);text-decoration:none}
  footer a:hover{color:var(--ink)}
  footer .fine{flex-basis:100%;line-height:1.55;max-width:80ch}
  .sr{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}
</style>
</head>
<body>

<header class="masthead">
  <div class="wrap">
    <a class="brand" href="https://tpxisoftware.com">
      <svg width="30" height="30" viewBox="0 0 64 64" aria-hidden="true">
        <defs><linearGradient id="bm" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stop-color="#2563eb"/><stop offset="1" stop-color="#7c3aed"/>
        </linearGradient></defs>
        <rect width="64" height="64" rx="14" fill="url(#bm)"/>
        <rect x="14" y="18" width="28" height="6" rx="3" fill="#fff"/>
        <rect x="20" y="29" width="28" height="6" rx="3" fill="#fff"/>
        <rect x="26" y="40" width="28" height="6" rx="3" fill="#fff"/>
      </svg>
      <span class="name">TPxi Software</span>
      <span class="tag">Less Admin. More Ministry.</span>
    </a>
    <nav>
      <a href="https://github.com/bswaby/Touchpoint">GitHub</a>
      <a href="https://bswaby.github.io/Touchpoint/DOC_SQLDocumentation.html">SQL docs</a>
      <a class="support-link" href="#support">Support this work</a>
    </nav>
  </div>
</header>

<div class="wrap lede">
  <h1>58 free tools for TouchPoint®</h1>
  <p>Every one of these started as a real problem at a real church, and every one is free to install. Search the index, pick a tool to read what it does, then take it to GitHub.</p>
</div>

<div class="controls">
  <div class="wrap">
    <div class="searchrow">
      <div class="field">
        <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
          <circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>
        </svg>
        <input id="q" type="search" autocomplete="off" placeholder="Search tools — try giving, attendance, volunteers, SQL" aria-label="Search tools" />
        <span class="slash" aria-hidden="true">/</span>
      </div>
      <span id="count" role="status">58 tools</span>
    </div>
    <div class="chips" id="chips"></div>
  </div>
</div>

<main class="wrap board">
  <div id="index"></div>

  <aside class="panel idle" id="panel" tabindex="-1" aria-live="polite">
    <button id="closepanel" aria-label="Close tool details">&times;</button>
    <div id="panelbody">
      <p class="kicker">Tool details</p>
      <h3>Pick a tool to see what it does</h3>
      <p class="body">Choose any tool from the index. New here? Start with Menu Editor — it&rsquo;s the one that puts every other tool in front of your staff.</p>
      <div class="actions">
        <button class="btn btn-quiet" type="button" onclick="selectTool('menu-editor')">Start with Menu Editor</button>
      </div>
    </div>
  </aside>
</main>

<section class="support" id="support">
  <div class="wrap">
    <div class="support-grid">
      <div>
        <h2>These stay free. Support is optional.</h2>
        <p>Fifty-eight tools, a hundred thousand lines of code, written in the evenings after the church day job. There is no account, no licence and no upsell attached to any of it.</p>
        <p>If a tool here saved your team hours, or kept a ministry moment from falling apart, you can put something back. It goes straight into the next tool.</p>
        <div class="products">
          <a class="product" href="https://displaycache.com">
            <strong>DisplayCache™</strong>
            <span>Church signage pulling live TouchPoint® data. $10/device/month.</span>
          </a>
          <a class="product" href="https://tpxigo.com">
            <strong>TPxi Go™</strong>
            <span>Your directory, caller ID and call logging, in Outlook and on your phone.</span>
          </a>
        </div>
      </div>

      <div class="give">
        <h3>Give once</h3>
        <p>Secure checkout through Stripe. No account needed, and nothing recurring unless you pick it.</p>
        <div class="amounts">
          <a class="amt" id="give25" href="#">$25</a>
          <a class="amt" id="give50" href="#">$50</a>
          <a class="amt" id="give100" href="#">$100</a>
        </div>
        <a class="btn btn-primary" id="giveAny" href="#">Choose your own amount</a>
        <p class="hint" style="margin-top:14px;font-size:13px;color:var(--ink-faint)">
          TPxi Software, LLC is a for-profit company, so gifts here are support, not tax-deductible donations.
        </p>
      </div>
    </div>
  </div>
</section>

<footer>
  <div class="wrap">
    <a href="https://tpxisoftware.com">tpxisoftware.com</a>
    <a href="https://github.com/bswaby/Touchpoint">github.com/bswaby/Touchpoint</a>
    <a href="https://tpxigo.com">TPxi Go™</a>
    <a href="https://displaycache.com">DisplayCache™</a>
    <a href="https://tpxiscan.com">TPxi Scan</a>
    <p class="fine">Built by Ben Swaby, Director of Technology Solutions at First Baptist Church Hendersonville. TouchPoint® is a registered trademark of Touchpoint Software, Inc. TPxi Software™, TPxi Go™ and DisplayCache™ are trademarks of TPxi Software, LLC. TPxi Software is not affiliated with, endorsed by, or sponsored by Touchpoint Software, Inc.</p>
  </div>
</footer>

<script>
/* ------------------------------------------------------------------
   1. YOUR STRIPE LINKS — replace these four, everything else is done.
   Stripe Dashboard → Payment links → New. One link per preset amount,
   plus one "customer chooses amount" link.
------------------------------------------------------------------- */
var STRIPE = {
  a25:  "https://buy.stripe.com/REPLACE_25",
  a50:  "https://buy.stripe.com/REPLACE_50",
  a100: "https://buy.stripe.com/REPLACE_100",
  any:  "https://buy.stripe.com/REPLACE_ANY"
};

var REPO = "https://github.com/bswaby/Touchpoint";

/* ------------------------------------------------------------------
   2. THE TOOLS. To add one: copy a line, change the fields.
      fav = community favorite (gold dot) · isNew = recent release
------------------------------------------------------------------- */
var CATEGORIES = [
  ["People & Attendance", "Who was here, how to reach them, what to celebrate"],
  ["Finance & Giving", "Money in, money out, and making it all reconcile"],
  ["Events & Registration", "Sign-ups, day-of operations, post-event cleanup"],
  ["Outreach & Engagement", "Lapsed members, prospects, and the wider community"],
  ["Ministry Insights", "Dashboards that show what is actually happening"],
  ["Volunteers, Tasks & Reports", "Coordinate the team, build the report you need"],
  ["System & Admin", "Keep the platform healthy and developers productive"]
];

var TOOLS = [
  // People & Attendance
  {n:"Live Search", c:0, fav:1, u:"/tree/main/TPxi/Live%20Search",
   d:"Type part of a name and get that person's full history on one screen. Log a note or a task without leaving the results."},
  {n:"Weekly Attendance (WAAG 2.0)", c:0, u:"/tree/main/TPxi/Weekly%20Attendance",
   d:"A rebuilt weekly attendance view for groups, with the comparisons and trend lines the standard report never gave you."},
  {n:"Attendance Markings", c:0, isNew:1, u:"/tree/main/TPxi/Attendance%20Markings",
   d:"Marks a whole roll present in one pass, then you uncheck the few who missed. About a fifth of the clicks for a full-attendance event."},
  {n:"Attendance Report Builder", c:0, isNew:1, u:"/tree/main/TPxi/Attendance%20Builder",
   d:"Build attendance reports across programs, divisions and involvements without writing any SQL."},
  {n:"In-Progress Registrations", c:0, u:"/tree/main/TPxi/In-Progress%20Registrations",
   d:"Shows the registrations people started and never finished, so you can follow up before the event fills."},
  {n:"Duplicate Involvements", c:0, u:"/tree/main/TPxi/Duplicate%20Involvements",
   d:"Finds people sitting in more than one involvement when they should only be in one."},
  {n:"Roll Sheet", c:0, fav:1, isNew:1, u:"/tree/main/TPxi/Roll%20Sheet",
   d:"Printable roll sheets built the way your teachers actually want them: which columns, what order, how much room to write."},
  {n:"Emergency List", c:0, u:"/tree/main/TPxi/Emergency%20List",
   d:"Pulls emergency contacts and medical notes for a group into one list you can hand to a leader."},
  {n:"Anniversaries Widget", c:0, u:"/tree/main/TPxi/Anniversaries",
   d:"Surfaces birthdays, membership anniversaries and other milestones on the home page so they don't slip past."},
  {n:"New Member Report", c:0, u:"/tree/main/TPxi/New%20Member%20Report",
   d:"Tracks new members through onboarding: who joined, what they've connected to, and who has gone quiet."},

  // Finance & Giving
  {n:"Weekly Contribution Report", c:1, fav:1, u:"/tree/main/TPxi/Contribution%20Report",
   d:"The weekly giving and reconciliation report most churches running these tools open every Monday morning."},
  {n:"Giving Dashboard", c:1, u:"/tree/main/TPxi/Giving%20Dashboard",
   d:"Giving trends, donor movement and fund performance in one view instead of five exports."},
  {n:"Statement Audit Dashboard", c:1, u:"/tree/main/TPxi/Statement%20Audit",
   d:"Finds the statement problems before your donors do: bad addresses, missing emails, electronic statements that never arrived."},
  {n:"Deposit Report", c:1, u:"/tree/main/TPxi/Deposit%20Report",
   d:"Reconciles a deposit end to end so the bank, the batch and the books agree."},
  {n:"Coupon Report", c:1, u:"/tree/main/TPxi/Coupon%20Report",
   d:"Shows where coupons are being used across ministries and what they are costing."},
  {n:"Envelope Number Report", c:1, u:"/tree/main/TPxi/Envelope%20Number%20Report",
   d:"A SQL report for giving envelope numbers: who has one, and who is actually using it."},
  {n:"Last 4 Search", c:1, u:"/tree/main/TPxi/Transaction%20Search",
   d:"Look up a transaction by the last four digits of the card or bank account, when that's all the donor can tell you."},
  {n:"Find Funds in Batch", c:1, u:"/tree/main/TPxi/Find%20Funds%20in%20Batch",
   d:"Tells you which batches contain a specific fund when you're chasing down a posting error."},
  {n:"Fortis Fees", c:1, u:"/tree/main/Finance/FortisFees",
   d:"Breaks down Fortis processing fees automatically instead of by hand every month."},
  {n:"QCD-Grant Letters", c:1, u:"/tree/main/TPxi/Finance%20Grant-QCD%20Letter",
   d:"Generates acknowledgement letters for qualified charitable distributions and donor-advised grants."},
  {n:"Payment Manager", c:1, fav:1, u:"/tree/main/TPxi/Payment%20Manager",
   d:"Tracks outstanding balances on fee-based registrations, sends receipts and handles the follow-up."},
  {n:"Involvement with Fees", c:1, u:"/tree/main/TPxi/Involvements%20with%20Fees",
   d:"Lists every involvement carrying a fee and shows what has been paid against it."},

  // Events & Registration
  {n:"Day of Registration", c:2, isNew:1, u:"/tree/main/TPxi/Day%20of%20Registration",
   d:"Assign walk-ins to classes and sessions at the table on event day, without slowing the line down."},
  {n:"Involvement Processor", c:2, fav:1, isNew:1, u:"/tree/main/TPxi/Involvement%20Processor",
   d:"The full registrant processing workflow in one screen. Replaces several disconnected steps you used to do in order."},
  {n:"Registration Export", c:2, u:"/tree/main/TPxi/Registration%20Export",
   d:"Exports registration data, question answers included, in a shape you can actually work with."},
  {n:"Registration Data Manager", c:2, u:"/tree/main/TPxi/Registration%20Data%20Manager",
   d:"Edit and clean registration answers in bulk instead of one record at a time."},
  {n:"FastLaneCheckIn", c:2, u:"/tree/main/TPxi/FastLaneCheckIn",
   d:"A stripped-down check-in flow for large events where the standard station can't keep up."},

  // Outreach & Engagement
  {n:"Communication Dashboard", c:3, u:"/tree/main/TPxi/Communication%20Dashboard",
   d:"Shows what you're sending, who is opening it, and which ministries have gone quiet."},
  {n:"Lapsed Attenders", c:3, u:"/tree/main/TPxi/Lapsed%20Attenders",
   d:"Finds people whose attendance has dropped off, early enough that a phone call still makes sense."},
  {n:"Prospector", c:3, isNew:1, u:"/tree/main/TPxi/Prospector",
   d:"A configurable pipeline for working prospects: assign, track, follow up, close the loop."},
  {n:"Auxiliary to Group Analytics", c:3, u:"/tree/main/TPxi/Auxiliary%20to%20Group%20Analytics",
   d:"Answers whether your events and auxiliary programs are actually moving people into groups."},
  {n:"TaskNote Activity Dashboard", c:3, u:"/tree/main/TPxi/TaskNote%20Activity%20Dashboard",
   d:"Shows task and note activity across the staff: what is moving, and what has stalled."},

  // Ministry Insights
  {n:"Ministry Structure", c:4, fav:1, u:"/tree/main/TPxi/Ministry%20Structure",
   d:"Your whole program, division and involvement hierarchy on one page, so you can see where it doesn't line up."},
  {n:"Enterprise Reporting", c:4, fav:1, isNew:1, u:"/tree/main/TPxi/Enterprise%20Reporting",
   d:"Over 100 reports behind a single dashboard. Replaces a folder of bookmarks and half-remembered saved queries."},
  {n:"Program Pulse", c:4, isNew:1, u:"/tree/main/TPxi/Program%20Pulse",
   d:"A quick read on what is actually happening in each program right now."},
  {n:"Missions Dashboard", c:4, isNew:1, u:"/tree/main/TPxi/Mission%20Dashboard",
   d:"Tracks mission trips end to end: teams, funding and who has gone."},
  {n:"Membership Analysis", c:4, u:"/tree/main/TPxi/Membership%20Analysis",
   d:"Demographic and trend analysis of your membership: who you have, and how that is changing."},
  {n:"Involvement Activity Dashboard", c:4, u:"/tree/main/TPxi/Involvement%20Activity%20Dashboard",
   d:"Engagement trends by involvement, so a leader can see their own group's direction without asking you."},
  {n:"Geographic Distribution Map", c:4, fav:1, u:"/tree/main/TPxi/Geographic%20Distribution%20Map",
   d:"Maps where your people live. Useful for campus, small group and outreach decisions."},

  // Volunteers, Tasks & Reports
  {n:"Volunteer Scheduler Report", c:5, u:"/tree/main/TPxi/Scheduler%20Report",
   d:"A full report on scheduled volunteers: who is serving, where you're short, who never confirmed."},
  {n:"Volunteer Widget", c:5, u:"/tree/main/TPxi/Volunteer%20Widget",
   d:"Shows the logged-in person their own upcoming assignments on the home page."},
  {n:"Task Runner", c:5, isNew:1, u:"/tree/main/TPxi/Task%20Runner",
   d:"Task management for you and your team, inside TouchPoint® rather than in yet another app."},
  {n:"QuickLinks", c:5, u:"/blob/main/TPxi/Quicklinks",
   d:"A permissioned quick-access menu with live counts on the links that matter."},
  {n:"Operations Checklists", c:5, isNew:1, u:"/tree/main/TPxi/Operations%20Checklists",
   d:"Recurring checks and reminders: the weekly and monthly things that get forgotten until they bite."},
  {n:"Report Writer", c:5, fav:1, isNew:1, u:"/tree/main/TPxi/Report%20Writer",
   d:"Build and save your own involvement and people reports without writing SQL."},

  // System & Admin
  {n:"Menu Editor", c:6, fav:1, isNew:1, u:"/tree/main/TPxi/Menu%20Editor",
   d:"Puts any script onto the People, Involvement, Finance, Admin or Blue Toolbar menus, with role permissions from a dropdown and an automatic backup before every save. Install this one first, and every other tool is one menu entry away from your staff."},
  {n:"API Explorer", c:6, u:"/tree/main/TPxi/API%20Explorer",
   d:"Try TouchPoint® API calls live and see what comes back before you write anything against them."},
  {n:"SQL Query Explorer", c:6, u:"/tree/main/TPxi/SQL%20Query%20Explorer",
   d:"Run and explore SQL against your database directly, from inside TouchPoint®."},
  {n:"Email Technical Diagnostics", c:6, u:"/tree/main/TPxi/Email%20Technical%20Diagnostics",
   d:"Traces email problems: what sent, what bounced, and what the receiving server actually said."},
  {n:"Account Security Monitor", c:6, u:"/tree/main/TPxi/Account%20Security%20Monitor",
   d:"Security analytics on user accounts: logins, roles and access that doesn't look right."},
  {n:"Person-Audit Detail", c:6, u:"/tree/main/TPxi/Person%20Attendance%20Audit",
   d:"Shows everywhere a person has served and attended. The report you want in hand when a question comes up about someone."},
  {n:"CSV Phone Matcher", c:6, u:"/tree/main/TPxi/CSV%20Phone%20Matcher.",
   d:"Matches a CSV of phone numbers back to people records."},
  {n:"Link Generator", c:6, u:"/tree/main/TPxi/Link%20Generator",
   d:"Creates pre-authenticated links, so people land where you sent them instead of at a login wall."},
  {n:"Attachment Link Downloader", c:6, u:"/tree/main/TPxi/Attachment%20Link%20Generator",
   d:"Bulk-downloads documents attached across records."},
  {n:"Involvement Sync", c:6, u:"/tree/main/TPxi/Involvement%20Sync",
   d:"Copies involvement settings across many groups at once so they stay consistent."},
  {n:"Involvement Owner Audit", c:6, u:"/tree/main/TPxi/Involvement%20Notification%20Audit%20Tool",
   d:"Finds involvements with no leader, no notification target, or the wrong person still attached."},
  {n:"User Activity", c:6, u:"/tree/main/TPxi/User%20Activity",
   d:"System usage by staff: who is in, what they use, and what they never touch."},
  {n:"TechStatus", c:6, fav:1, u:"/blob/main/Python%20Scripts/TechStatus/TechStatus",
   d:"Health and performance monitoring for your TouchPoint® instance."}
];

/* ------------------------------------------------------------------
   3. Page behaviour
------------------------------------------------------------------- */
var slug = function (s) {
  return s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, "");
};
TOOLS.forEach(function (t) { t.id = slug(t.n); t.href = REPO + t.u; });

var indexEl = document.getElementById("index");
var chipsEl = document.getElementById("chips");
var countEl = document.getElementById("count");
var qEl = document.getElementById("q");
var panel = document.getElementById("panel");
var panelBody = document.getElementById("panelbody");
var activeCat = -1;
var selected = null;

// category chips
var allChip = document.createElement("button");
allChip.className = "chip";
allChip.type = "button";
allChip.textContent = "All 58";
allChip.setAttribute("aria-pressed", "true");
allChip.onclick = function () { setCat(-1); };
chipsEl.appendChild(allChip);

CATEGORIES.forEach(function (c, i) {
  var b = document.createElement("button");
  b.className = "chip";
  b.type = "button";
  b.textContent = c[0];
  b.setAttribute("aria-pressed", "false");
  b.onclick = function () { setCat(i); };
  chipsEl.appendChild(b);
});

// build the index
CATEGORIES.forEach(function (c, i) {
  var sec = document.createElement("section");
  sec.className = "cat";
  sec.dataset.cat = i;

  var head = document.createElement("div");
  head.className = "cat-head";
  var h = document.createElement("h2");
  h.textContent = c[0];
  var n = document.createElement("span");
  n.className = "n";
  var list = document.createElement("div");
  list.className = "list";

  var mine = TOOLS.filter(function (t) { return t.c === i; });
  n.textContent = mine.length;
  head.appendChild(h); head.appendChild(n);
  sec.appendChild(head);

  mine.forEach(function (t) {
    var b = document.createElement("button");
    b.type = "button";
    b.className = "tool" + (t.fav ? " fav" : "");
    b.id = "t-" + t.id;
    b.dataset.id = t.id;
    b.setAttribute("aria-current", "false");
    b.innerHTML =
      '<span class="dot" aria-hidden="true"></span>' +
      '<span class="label"></span>' +
      (t.isNew ? '<span class="new">new</span>' : "");
    b.querySelector(".label").textContent = t.n;
    b.onclick = function () { selectTool(t.id); };
    list.appendChild(b);
    t.el = b;
  });

  sec.appendChild(list);
  indexEl.appendChild(sec);
});

var emptyEl = document.createElement("p");
emptyEl.className = "empty";
emptyEl.style.display = "none";
indexEl.appendChild(emptyEl);

function setCat(i) {
  activeCat = i;
  Array.prototype.forEach.call(chipsEl.children, function (c, idx) {
    c.setAttribute("aria-pressed", String(idx - 1 === i));
  });
  applyFilter();
}

function applyFilter() {
  var q = qEl.value.trim().toLowerCase();
  var shown = 0;

  TOOLS.forEach(function (t) {
    var hit = !q || (t.n + " " + t.d + " " + CATEGORIES[t.c][0] + " " + CATEGORIES[t.c][1])
      .toLowerCase().indexOf(q) > -1;
    if (activeCat > -1 && t.c !== activeCat) hit = false;
    t.el.style.display = hit ? "" : "none";
    if (hit) shown++;
  });

  Array.prototype.forEach.call(indexEl.querySelectorAll(".cat"), function (sec) {
    var any = Array.prototype.some.call(sec.querySelectorAll(".tool"), function (b) {
      return b.style.display !== "none";
    });
    sec.style.display = any ? "" : "none";
    if (any) {
      sec.querySelector(".n").textContent =
        Array.prototype.filter.call(sec.querySelectorAll(".tool"), function (b) {
          return b.style.display !== "none";
        }).length;
    }
  });

  countEl.textContent = shown === 1 ? "1 tool" : shown + " tools";
  if (shown === 0) {
    emptyEl.style.display = "";
    emptyEl.textContent = "Nothing matches “" + qEl.value.trim() +
      "”. Try a ministry word instead — giving, attendance, volunteers, registration, email.";
  } else {
    emptyEl.style.display = "none";
  }
}

function selectTool(id) {
  var t = TOOLS.filter(function (x) { return x.id === id; })[0];
  if (!t) return;

  if (selected) selected.el.setAttribute("aria-current", "false");
  selected = t;
  t.el.setAttribute("aria-current", "true");

  var marks = "";
  if (t.fav) marks += '<span class="mark fav">Community favorite</span>';
  if (t.isNew) marks += '<span class="mark new">Recent release</span>';
  marks += '<span class="mark">' + CATEGORIES[t.c][0] + "</span>";

  panelBody.innerHTML =
    '<p class="kicker">Tool details</p>' +
    "<h3></h3>" +
    '<div class="marks">' + marks + "</div>" +
    '<p class="body"></p>' +
    '<div class="actions">' +
      '<a class="btn btn-primary" href="' + t.href + '">Open on GitHub</a>' +
      '<button class="btn btn-quiet" type="button" id="copylink">Copy link to this tool</button>' +
    "</div>" +
    '<p class="hint">Copy the script into <strong>Admin → Advanced → Special Content → Python</strong>, then use ' +
    '<a href="' + REPO + '/tree/main/TPxi/Menu%20Editor">Menu Editor</a> to put it on a menu. ' +
    "The tool's own README covers anything specific to it.</p>";
  panelBody.querySelector("h3").textContent = t.n;
  panelBody.querySelector(".body").textContent = t.d;

  panelBody.querySelector("#copylink").onclick = function (e) {
    var url = location.href.split("#")[0] + "#" + t.id;
    var done = function () { e.target.textContent = "Link copied"; 
      setTimeout(function(){ e.target.textContent = "Copy link to this tool"; }, 1800); };
    if (navigator.clipboard) { navigator.clipboard.writeText(url).then(done, done); }
    else { done(); }
  };

  panel.classList.remove("idle");
  panel.classList.add("open");
  if (history.replaceState) history.replaceState(null, "", "#" + t.id);
  if (window.innerWidth < 1060) panel.focus();
}

document.getElementById("closepanel").onclick = function () {
  panel.classList.remove("open");
  setTimeout(function () { panel.classList.add("idle"); }, 220);
  if (selected) { selected.el.setAttribute("aria-current", "false"); selected.el.focus(); selected = null; }
};

qEl.addEventListener("input", applyFilter);

document.addEventListener("keydown", function (e) {
  if (e.key === "/" && document.activeElement !== qEl) { e.preventDefault(); qEl.focus(); qEl.select(); }
  if (e.key === "Escape") {
    if (document.activeElement === qEl && qEl.value) { qEl.value = ""; applyFilter(); }
    else if (panel.classList.contains("open") && window.innerWidth < 1060) {
      document.getElementById("closepanel").click();
    }
  }
});

// stripe links
document.getElementById("give25").href = STRIPE.a25;
document.getElementById("give50").href = STRIPE.a50;
document.getElementById("give100").href = STRIPE.a100;
document.getElementById("giveAny").href = STRIPE.any;

// deep link support: tools.html#live-search
if (location.hash.length > 1) {
  var target = location.hash.slice(1);
  if (TOOLS.some(function (t) { return t.id === target; })) selectTool(target);
}

applyFilter();
</script>
</body>
</html>
