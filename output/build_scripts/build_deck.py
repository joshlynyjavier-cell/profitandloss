import json, re
V = json.load(open('deck/values.json'))
S, P, D, I, MK, ES, PE = V['seg'], V['pe'], V['dd'], V['inc'], V['mk'], V['es'], V['pe']


def svg(n):
    s = open(f'deck/{n}.svg').read()
    s = s[s.index('<svg'):]
    return re.sub(r'(<svg[^>]*?)\swidth="[^"]*"\sheight="[^"]*"', r'\1 width="100%" height="100%"', s, count=1)


m = lambda x, d=1: ("−" if x < 0 else "") + f"${abs(x)/1e6:,.{d}f}M"
p = lambda x, d=1: f"{x*100:.{d}f}%"

css = """
@page { size: 13.333in 7.5in; margin: 0 }
* { box-sizing: border-box; margin: 0; padding: 0 }
body { font-family: Inter, Arial, sans-serif; color: #1E2530; background: #fff; }
.slide { width: 13.333in; height: 7.5in; padding: 0.42in 0.55in 0.35in; position: relative; page-break-after: always; overflow: hidden; }
.kicker { font-size: 10.5pt; letter-spacing: .08em; text-transform: uppercase; color: #6B7684; font-weight: 600; }
h1 { font-size: 21pt; line-height: 1.18; font-weight: 700; color: #1F3A5F; margin: 6px 0 14px; max-width: 12in; }
.grid { display: grid; gap: 22px; }
.g2 { grid-template-columns: 1fr 1fr } .g55 { grid-template-columns: 1.15fr 1fr } .g64 { grid-template-columns: 1.35fr 1fr }
.card { border-top: 2px solid #1F3A5F; padding-top: 8px; }
h3 { font-size: 11pt; font-weight: 700; color: #1F3A5F; margin-bottom: 6px; }
p, li { font-size: 10.5pt; line-height: 1.42; color: #2B3440; }
ul { padding-left: 16px } li { margin-bottom: 5px }
.big { font-size: 26pt; font-weight: 700; color: #1F3A5F; line-height: 1 }
.big.red { color: #B5482F } .big.grn { color: #2E7D5B }
.lbl { font-size: 9pt; color: #6B7684; margin-top: 3px; line-height: 1.3 }
.tiles { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 14px }
.tile { border-left: 3px solid #1F3A5F; padding: 2px 0 2px 10px }
.tile.r { border-color: #B5482F } .tile.g { border-color: #2E7D5B }
table { border-collapse: collapse; width: 100%; font-size: 9.6pt }
th { text-align: right; font-weight: 600; color: #6B7684; border-bottom: 1px solid #9AA3AE; padding: 4px 6px; font-size: 8.8pt }
td { text-align: right; padding: 4px 6px; border-bottom: 1px solid #E6E9ED }
th:first-child, td:first-child { text-align: left }
tr.hl td { font-weight: 700; background: #F3F5F8 }
.neg { color: #B5482F } .pos { color: #2E7D5B }
.chart { width: 100% }
.foot { position: absolute; bottom: 0.18in; left: 0.55in; right: 0.55in; font-size: 7.8pt; color: #8A94A1; display: flex; justify-content: space-between; border-top: 1px solid #E6E9ED; padding-top: 5px }
.pill { display: inline-block; background: #1F3A5F; color: #fff; font-weight: 700; padding: 3px 10px; border-radius: 3px; font-size: 10pt; letter-spacing: .04em }
.note { font-size: 8.8pt; color: #6B7684; line-height: 1.35 }
.act td, .act th { text-align: left; font-size: 8.8pt; padding: 3px 6px }
.callout { background: #F6F7F9; border-left: 3px solid #B5482F; padding: 8px 12px; margin-top: 10px }
"""


def foot(n, src):
    return f'<div class="foot"><span>Sola Kitchen · Board review · Marketplace delivery P&amp;L · Fictional company, synthetic data</span><span>Source: Sola_Kitchen_Financial_Model.xlsx — {src} · {n}/5</span></div>'


seg_rows = [("DoorDash", "DoorDash"), ("UberEats", "UberEats"), ("Grubhub", "Grubhub"),
            ("Established markets (5 regions)", "Established"), ("East Coast (12 new stores)", "East Coast (New)"),
            ("Total 3PD", "Total")]


def segtable():
    h = "<table><tr><th>Segment</th><th>FY gross</th><th>All-in take</th><th>CM2 % FY</th><th>CM2 % exit</th><th>CM2 / order</th><th>CM2 mix</th></tr>"
    for lab, k in seg_rows:
        s = S[k]
        cls = ' class="hl"' if k == 'Total' else ''
        h += (f"<tr{cls}><td>{lab}</td><td>{m(s['gross_ann'])}</td><td>{p(s['take'])}</td><td>{p(s['cm2p'])}</td>"
              f"<td class='neg'>{p(s['exit_cm2p'])}</td><td>${s['cm2_order']:.2f}</td><td>{p(s['cm_mix'],0)}</td></tr>")
    return h + "</table>"


s1 = f"""
<section class="slide">
 <div class="kicker">Board review · Third-party marketplace delivery · Recommendation</div>
 <h1>HOLD: keep all three platforms, stop paying for growth, and reprice — about {m(ES['tot'],0)} a year of contribution is recoverable without exiting anything</h1>
 <div class="tiles">
  <div class="tile"><div class="big">{m(V['cm2_fy'])}</div><div class="lbl">FY CM2 on {m(V['gross_fy'],0)} gross ({p(V['cm2p_fy'])}) — vs ~48% for a dine-in dollar</div></div>
  <div class="tile r"><div class="big red">{p(V['cm2p_exit'])}</div><div class="lbl">Exit run-rate CM2 % (June period): restaurant-funded promo has quadrupled since Feb</div></div>
  <div class="tile"><div class="big">{p(V['be_fy'],0)}</div><div class="lbl">Share of orders that must be truly incremental to beat dine-in. Evidence points to ~{p(I['incr'],0)}</div></div>
  <div class="tile g"><div class="big grn">+{m(ES['tot'])}</div><div class="lbl">Annual CM2 from the plan → ~{m(ES['cm2_after'])} run-rate (~{p(ES['cm2p_after'],0)} of gross)</div></div>
 </div>
 <div class="grid g55">
  <div>
   <div class="card"><h3>Objective: margin (CM2 dollars / EBITDA quality ahead of the raise)</h3>
   <ul>
    <li><b>Is it a business?</b> Yes, on its own terms: every store × platform earns positive CM2. But it keeps only {p(V['share_dinein'],0)} of the contribution a dine-in dollar earns, and its value depends on incrementality we cannot prove.</li>
    <li><b>Actions:</b> ① roll promo back to ~1.5% of gross <b>+{m(ES['p1'])}</b>; ② 5% 3PD menu markup <b>+{m(ES['p2'])}</b> (half the test-implied gain); ③ counter DoorDash at 28.5% / ≤40% ads / 12 months <b>−{m(-ES['p3'])}</b> vs today, {m(D['sav_sn'])} better than accepting over 24 months.</li>
    <li><b>Confidence: medium.</b> High on the P&amp;L, medium-high on the promo roll-back, low-medium on markup response and incrementality.</li>
    <li><b>Trade-off:</b> roughly 1–3% of 3PD order volume, and less fuel for the East Coast growth story. If the board optimizes for <b>growth</b> instead, keep promo where it retains customers (SoCal/PNW/Northeast) and accept DoorDash's 24-month lock — this costs about <b>{m(-ES['gv_gap'])}</b> a year of CM2 vs this plan.</li>
   </ul></div>
  </div>
  <div><h3>CM2 run-rate bridge, $M per year</h3><div class="chart" style="height:3.3in">{svg('c1_bridge')}</div></div>
 </div>
 {foot(1, 'Exec_Summary, PL_Waterfall, Promo_Escalation, Markup, DoorDash')}
</section>"""

s2 = f"""
<section class="slide">
 <div class="kicker">Q1 · Is marketplace delivery a business?</div>
 <h1>Delivery makes money, but thinly: $100 of gross leaves ${V['cm2p_fy']*100:.0f} of contribution, against $48 in the dining room</h1>
 <div class="grid g64">
  <div><h3>Gross to CM2, FY annualized (51 weeks × 52/51), $M</h3><div class="chart" style="height:3.4in">{svg('c2_waterfall')}</div>
   <p class="note">Tax collected and remitted by the platforms (~$6.3M) is excluded. The bank deposit is not contribution: COGS (30%) and variable labor (16% of gross, both assignment assumptions) come out after it.</p></div>
  <div>
   <h3>Profitability by platform and market</h3>
   {segtable()}
   <ul style="margin-top:10px">
    <li><b>The all-in take is {p(V['take_fy'])}, not the 27.7% headline commission.</b> Promo, ads and fees add about 8 points.</li>
    <li><b>The margin problem is UberEats, not DoorDash:</b> {p(S['UberEats']['cm2p'])} FY, {p(S['UberEats']['exit_cm2p'])} at exit. Grubhub is the best platform at {p(S['Grubhub']['cm2p'])}.</li>
    <li><b>Geography does not matter; maturity does.</b> All five established regions sit at 18.7%. The East Coast runs a {p(S['East Coast (New)']['take'])} take and {p(S['East Coast (New)']['promo_p'])} promo, for {p(S['East Coast (New)']['cm2p'])} CM2.</li>
   </ul>
  </div>
 </div>
 {foot(2, 'PL_Waterfall, Platform_Market')}
</section>"""

s3 = f"""
<section class="slide">
 <div class="kicker">Q2 · Growth vs profitability</div>
 <h1>Delivery growth is real, but about half of it appears to come out of the dining room — below the {p(V['be_fy'],0)} needed to create value</h1>
 <div class="grid g2">
  <div><h3>Same-store revenue per store-week, indexed (Jul-25 = 100)</h3><div class="chart" style="height:2.3in">{svg('c3_trend')}</div>
   <p class="note">Months contain 4 or 5 payout weeks, so all figures are per store-week. Only one year of data, so seasonality cannot be separated from trend.</p></div>
  <div><h3>Break-even incrementality vs the evidence</h3><div class="chart" style="height:2.3in">{svg('c3_breakeven')}</div>
   <p class="note">Break-even = 1 − CM2% ÷ 48%. A cannibalized delivery dollar replaces one that would have earned ~48% in-store.</p></div>
 </div>
 <div class="grid" style="grid-template-columns:repeat(4,1fr);margin-top:10px;gap:18px">
  <div class="card"><h3>Trend</h3><p>3PD grew <b>+{p(I['g_3pd'],0)}</b> (Q1→Q4) against in-store <b>+{p(I['g_in'],0)}</b> and other own channels <b>+{p(I['g_ctrl'],0)}</b>. The dine-in shortfall equals ~{p(I['disp'],0)} of 3PD's above-trend growth, implying <b>~{p(I['incr'],0)} incrementality</b>. Indicative only: the store-level correlation is +{I['corr']:.2f}, not negative.</p></div>
  <div class="card"><h3>Capacity</h3><p><b>{p(I['peak'],0)}</b> of 3PD orders land at lunch or dinner, when <b>{p(I['peak_util'],0)}</b> of seats are already full. There, delivery competes with dine-in for the line. Off-peak seat utilization is {p(I['off_util'],0)}, which is where delivery is plausibly additive.</p></div>
  <div class="card"><h3>Retention (90-day)</h3><p>Dine-in <b>{p(I['r90_dine'],0)}</b>, 1P delivery {p(I['r90_1p'],0)}, DoorDash <b>{p(I['r90_dd'],0)}</b>, UberEats {p(I['r90_ue'],0)}. Promo-acquired customers are lower still ({p(I['r90_ddp'],0)}–{p(I['r90_uep'],0)}). The platforms keep the customer.</p></div>
  <div class="card"><h3>Markup test</h3><p>A 5% markup cost {p(-MK['v5'])} of volume and raised net revenue per store by +{p(MK['n5'])} (t≈{MK['t5']:.1f}). The 10% and 15% arms are not significant. Elasticity is ≈{MK['el']:.1f}, but CM2 still rises because food and labor costs don't rise with price.</p></div>
 </div>
 <div class="callout"><p><b>So what:</b> at 55% incrementality the channel's economic value is about <b>{m(I['ev']['base'][0])}</b> FY ({m(I['ev']['base_exit'][0])} at the exit rate). At 80% it is <b>+{m(I['ev']['bull'][0])}</b>. The evidence does not justify leaning in or exiting. Reprice first, then settle the question with a geo-holdout test.</p></div>
 {foot(3, 'Incrementality, Markup, Sensitivity')}
</section>"""

s4 = f"""
<section class="slide">
 <div class="kicker">Q3 · The DoorDash terms (45-day clock)</div>
 <h1>Counter, don't accept or walk: 31% costs {m(D['acc24'])} over the 24-month lock, but walking costs more under base assumptions</h1>
 <div class="grid g55">
  <div>
   <table>
    <tr><th>DoorDash economics (Y1 gross {m(D['g1'])})</th><th>Current</th><th>Proposed</th><th>Our counter</th></tr>
    <tr><td>Headline commission</td><td>27.0%</td><td class="neg">31.0%</td><td>28.5%</td></tr>
    <tr><td>Marketing commission</td><td>1.5%</td><td>1.5%</td><td>1.5%</td></tr>
    <tr><td>Ads share of DoorDash marketing</td><td>{p(D['ad_share'],0)}</td><td class="neg">≥60%</td><td>≤40%, spend-neutral</td></tr>
    <tr><td>Term</td><td>Annual</td><td class="neg">24 months, no exit</td><td>12 months (or 24 with a 12-mo exit)</td></tr>
    <tr><td>DoorDash all-in take (FY mix)</td><td>34.0%</td><td class="neg">{p(D['take_acc'])} – {p(D['take_top'])}</td><td>35.5%</td></tr>
    <tr><td>DoorDash CM2 %</td><td>{p(D['cm2p'])}</td><td class="neg">{p(D['cm2_acc'])} – {p(D['cm2_top'])}</td><td>{p(D['cm2_ctr'])}</td></tr>
    <tr class="hl"><td>CM2 cost vs today: 12 months / 24 months</td><td>—</td><td class="neg">{m(D['acc12'])} / {m(D['acc24'])}<br>to {m(D['top12'])} / {m(D['top24'])}</td><td>{m(D['ctr12'])} / {m(D['ctr24'])}</td></tr>
   </table>
   <ul style="margin-top:12px">
    <li><b>Walk-away number:</b> walking (30% of DoorDash sales switch to UberEats/Grubhub, 15% return to our own channels) costs {m(D['walk12'])} a year. It only beats accepting above <b>~{p(D['be_tu'],0)}</b> with a forced ad top-up ({p(D['be_sn'],0)} if the ad mix is spend-neutral). So sign at <b>≤30%</b> with a spend-neutral mix, and never at 31% with a 24-month lock and an ad top-up.</li>
    <li><b>Leverage:</b> DoorDash is 55% of our 3PD sales and growing ~8% a year. We can steer volume to Grubhub (22.5% commission, best CM2) and to our own 1P channel.</li>
    <li><b>The ads mandate is cheap once promo is fixed:</b> with promo back at baseline, reaching 60% ads needs only ~{m(D['post_topup'],2)}.</li>
    <li>A {p(D['mk_off'])} DoorDash markup would neutralize +4 points at flat volume. But that lever is ours either way, so it should not be used to justify accepting.</li>
   </ul>
  </div>
  <div><h3>CM2 cost vs today's terms, $M</h3><div class="chart" style="height:3.2in">{svg('c4_dd')}</div>
   <p class="note">Year 2 assumes DoorDash gross grows 8% (Contract_Terms). The counter's 24-month figure assumes the same terms are renewed. The walk result is highly sensitive to recapture: across the recapture grid (DoorDash sheet), the Y1 walk cost ranges from −$4.7M to +$8.1M.</p></div>
 </div>
 {foot(4, 'DoorDash, Sensitivity')}
</section>"""

s5 = f"""
<section class="slide">
 <div class="kicker">Q4 · The second-order question</div>
 <h1>The bigger leak is self-inflicted: Sola's own promo spend has quadrupled since February and bought no volume</h1>
 <div class="grid g55">
  <div><h3>Weekly, all platforms: promo as % of gross (top) vs established-store gross per store-week (bottom)</h3><div class="chart" style="height:2.55in">{svg('c5_promo')}</div></div>
  <div>
   <ul>
    <li>Restaurant-funded promo went from <b>{p(P['pre_promo'])} → {p(P['exit_promo'])}</b> of gross, across every platform and region at once (UberEats {p(P['ue_exit'])}, East Coast {p(P['ec_exit'])}).</li>
    <li>Volume did not respond. Established gross per store-week is <b>{p(P['gpsw_chg'])}</b> vs late January, and the 3PD-to-dine-in ratio grew more slowly after the ramp ({p(P['ratio_post'])}/mo) than before it ({p(P['ratio_pre'])}/mo).</li>
    <li>It costs <b>{m(P['esc_cost'])} a year</b> at exit run-rate, <b>{P['ratio_dd']:.1f}×</b> DoorDash's ask. The board is negotiating price with DoorDash while its own promo governance sets the margin, and full-year reporting ({p(V['cm2p_fy'])}) hides a {p(V['cm2p_exit'])} exit rate.</li>
   </ul>
   <p class="note" style="margin-top:6px">Verify before the meeting: is this a deliberate campaign, platform auto-enrolled promos, or a data/classification change?</p>
  </div>
 </div>
 <table class="act" style="margin-top:8px">
  <tr><th>Action</th><th>Owner</th><th>When</th><th>Metric → decision threshold</th></tr>
  <tr><td>Freeze and roll back restaurant-funded promo to ≤2% of 3PD gross; diagnose the February ramp</td><td>CMO + FP&amp;A</td><td>Next period</td><td>Promo ≤2%. If gross/store-wk falls >5% vs the dine-in trend, restore selectively (SoCal, PNW, NE first)</td></tr>
  <tr><td>Counter DoorDash: 28.5%, ≤40% ads spend-neutral, 12-month term</td><td>CFO + VP Off-Premise</td><td>Day 30 of 45</td><td>Sign at ≤30% spend-neutral. Walk-away at ≥32% with a forced ad top-up</td></tr>
  <tr><td>5% 3PD menu markup on all platforms (UberEats/Grubhub as a staged pilot)</td><td>VP Off-Premise + Pricing</td><td>2 periods</td><td>Order loss ≤3%. CM2/order up ≥$0.75</td></tr>
  <tr><td>Geo-holdout incrementality test (pause 3PD in matched stores/dayparts)</td><td>FP&amp;A + Ops</td><td>2–3 periods</td><td>&lt;55% → pull back, UberEats first (break-even 71%). &gt;70% → lean in</td></tr>
  <tr><td>Report 3PD monthly on period run-rate CM2, not FY gross</td><td>CFO</td><td>Monthly</td><td>Run-rate CM2 ≥20%. Escalate if &lt;16% for 2 periods</td></tr>
 </table>
 {foot(5, 'Promo_Escalation, Weeks, Incrementality, Exec_Summary')}
</section>"""

html = f"<!doctype html><html><head><meta charset='utf-8'><title>Sola Kitchen Board Recommendation</title><style>{css}</style></head><body>{s1}{s2}{s3}{s4}{s5}</body></html>"
open('deck/deck.html', 'w').write(html)
print('ok')
