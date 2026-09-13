# What makes a residential or townhouse project feasible

Background notes gathered while building this model. The benchmarks are
Australian and UK practice, which is where this vocabulary comes from. Figures
vary by city and by year — treat them as orders of magnitude, not quotes.

---

## 1. The question the model exists to answer

A feasibility model is not a forecast. It answers one question:

> **Is there enough margin between what this will sell for and what it will cost
> to justify the risk and the capital?**

Everything else is detail feeding that sentence. The Upwork brief reduces it
further — to a single green or red verdict against a 25% profit-on-cost
threshold — which is a reasonable simplification for a lead magnet.

## 2. The calculation spine

Every model in this family runs the same chain, whatever it calls the steps:

```
Gross realisation (GDV)        avg retail value  x  number of dwellings
  less selling costs
= Net realisation

Land purchase                  avg land value  x  number of lots
  plus stamp duty / acquisition costs
= Total purchase

Build cost
  plus statutory contributions
  plus professional fees
  plus contingency
  plus finance / holding costs
= Total development cost

Total cost = Total purchase + Total development cost
Net profit = Net realisation - Total cost
Profit on cost = Net profit / Total cost      <- the verdict
```

The same inputs can be run backwards instead, solving for the land price that
leaves the required margin. That is **residual land value** — "the most I can
pay for this site" — and it is the other standard output of these models. The
brief does not ask for it, but a developer will expect the model to imply it.

## 3. The thresholds people actually use

| Measure | Benchmark | Notes |
|---|---|---|
| **Profit on cost** | **25% minimum** | The brief's threshold. Industry norm |
| Profit on GDV | 20% (RICS viability benchmark) | Roughly equivalent to 25% on cost |
| Lender minimum, profit on GDV | 15–20% | Below 15% is hard to fund at all |
| First-time developers | Upper end of that range | Less track record, more margin demanded |
| IRR hurdle | ~25% | Where models run a timeline; this one does not |

The brief's 25% on cost is therefore not arbitrary — it is the standard bar, and
it maps to the 20%-on-GDV benchmark lenders use.

## 4. The variables

### Revenue side

| Variable | Why it matters |
|---|---|
| Retail comparables per dwelling | Drives GDV. The single most influential input |
| Number of dwellings / lots | Multiplies everything. Set by planning, not by choice |
| Product mix and size | Bigger dwellings cost more per unit but not proportionally more per m² |
| Sales rate / absorption | How fast the market absorbs the stock. Drives holding cost |

### Land side

| Variable | Typical treatment |
|---|---|
| Raw land comparables | Averaged from comparable sales |
| Stamp duty / transfer duty | % of purchase, varies by state |
| Acquisition legals, due diligence | Often folded into a single % |
| Demolition and site clearing | Separate line where relevant |

### Construction

| Variable | Australian range (2026) |
|---|---|
| Townhouse build, national | ~$3,170–$4,320 /m² ex GST |
| Sydney townhouse | ~$2,700–$4,500+ /m² |
| Melbourne townhouse | ~$2,800–$3,300 /m² incl GST |
| Brisbane townhouse | ~$1,900–$3,300 /m² |

Spread that wide is not sloppiness — it reflects specification, site conditions,
dwelling count, structural complexity and parking arrangement.

### Statutory

Developer contributions run **8–11% of total development cost**, and in Sydney
growth areas can reach **$85,000 per dwelling** — enough to compress margins by
about 4 percentage points and cut residual land value by around a third.

A useful rule from the same source: every additional **$10,000 per dwelling** in
contributions reduces residual land value dollar-for-dollar and compresses IRR
by roughly 2.5 percentage points over a 24-month project.

GST is the other statutory swing. Under the **margin scheme**, GST is paid on
the margin rather than the full sale price. On a $3M development where land cost
$800k, that is $200k of GST instead of $272,727 — about $73k saved.

### Finance and holding

The most commonly underestimated cost in Australian development is **holding
cost during delays**. On $1M of land at 7.5%, every month of delay costs $6,250
in interest alone, before rates, insurance and land tax.

### Selling

Agent commission, marketing, legals — usually a single percentage of gross
realisation, commonly 2–3%.

## 5. Residential subdivision vs townhouse

This is the distinction the brief builds two separate tabs for, and the tabs are
not just different numbers — the cost structure genuinely differs.

| | House-and-land / subdivision | Townhouse |
|---|---|---|
| **Zoning** | Permitted broadly; duplex allowed in R2 low-density | Multi-dwelling housing generally **not** permitted in R2 |
| **Build cost** | Lower per m²; detached, simpler | Higher — party walls, fire separation |
| **Common property** | Minimal | Driveways, visitor parking, shared landscaping and services |
| **Site** | Usually easier access | Tighter sites, harder access, preliminaries spread over a slower build |
| **Title** | Torrens lots | Strata or community title, with body corporate setup |
| **Capital** | Lower | Higher — more land, more build, longer |

A worked comparison from the research: a duplex of 2 × 150 m² on a 600 m² block
returns roughly **$229k–$512k** profit. A townhouse scheme of 4 × 120 m² on
1,200 m² returns **$382k–$782k** — but on twice the land cost and materially
more capital at risk.

The conclusion matters for how the model gets used: **a well-designed duplex can
beat a denser townhouse scheme** where it is easier to approve, cheaper to build
and better matched to the local market. Density is not automatically better.

## 6. The gates that kill a deal before the maths runs

A financial model assumes the project is legal and buildable. These decide that,
and none of them appear in the workbook:

- **Zoning and permitted use** — whether the dwelling type is allowed at all
- **Setbacks** — each one carves away raw lot area, leaving the buildable envelope
- **Maximum site coverage** and minimum private open space
- **Car parking** — count, dimensions, visitor spaces, manoeuvring areas
- **Slope, soil, rock, flooding** — cost without adding saleable floor area
- **Stormwater and servicing** — can require infrastructure disproportionate to the site
- **Access** — frontage width, crossover location, construction access
- **Title and easements**

## 7. Which variables actually move the answer

Sensitivity analysis is standard practice, and the hierarchy is consistent:

1. **GDV is the most influential input.** A 5% drop in GDV produces a much
   larger fall in residual land value, because profit and surplus shrink
   together. A swing in GDV has roughly **double the impact** of the same swing
   in construction cost.
2. **Construction cost is second.** Modest inflation eats the contingency, then
   the profit.
3. Contributions, finance rate and holding period follow.

Standard test range is **±10%**, sometimes ±5% increments on GDV and build cost.
A 5% GDV reduction or a 10% construction cost increase can wipe out the entire
margin.

**Implication for this job:** a lead-magnet workbook that produces one number
from one set of assumptions is, strictly, less useful than one showing what
happens when retail values fall 5%. Worth raising — not as scope creep, but as
the obvious version two.

## 8. What the demo model covers, and what it does not

Covered by `build_workbook.py` today:

land comparables · retail comparables · dwelling count · build cost per dwelling ·
council contributions per dwelling · stamp duty % · professional fees % ·
contingency % · selling costs % · finance % · editable profit threshold →
market averages, total purchase, gross and net realisation, total development
costs, total cost, net profit, profit on cost, verdict.

Not covered, and each one is a question worth asking the client rather than
assuming:

| Gap | The question |
|---|---|
| GST / margin scheme | Is GST handled in the reference model, or deliberately left out? |
| Residual land value | Should the model also solve backwards for the maximum land price? |
| Timeline and holding costs | Finance is a flat % here. Does the reference model use a project duration? |
| Demolition, site works, servicing | Separate lines, or absorbed into build cost? |
| Sensitivity table | Version two, or out of scope? |
| Strata / body corporate setup | Relevant to the Townhouse tab specifically |

That table is the most useful thing on this page. It turns "I read your brief"
into six specific questions that only someone who has built the thing would ask
— and the brief explicitly says the client values people who challenge bad
assumptions rather than saying yes to everything.

---

## Sources

- [Feasly — Property Development Feasibility Guide Australia](https://www.feasly.com.au/guides/property-development-feasibility-australia)
- [Feasly — Townhouse Development in Australia](https://www.feasly.com.au/guides/townhouse-development-australia-developer-guide)
- [Feasly — Land Subdivision in Australia](https://www.feasly.com.au/guides/land-subdivision-australia-developer-guide)
- [Feasly — Construction Cost per Square Metre in Australia (2026)](https://www.feasly.com.au/guides/construction-cost-per-square-metre-australia)
- [Feasly — Developer Contributions & Infrastructure Levies](https://www.feasly.com.au/guides/developer-contributions-infrastructure-levies-australia)
- [Feasly — GST Margin Scheme for Property Development](https://www.feasly.com.au/guides/gst-margin-scheme-property-development)
- [Feasly — Residual Land Value in Australia](https://www.feasly.com.au/guides/residual-land-value-australia)
- [Feasly — Sensitivity Analysis for Property Development](https://www.feasly.com.au/guides/sensitivity-analysis-property-development-guide)
- [Altus Group — Property Development Feasibility Study](https://www.altusgroup.com/featured-insights/property-development-feasibility/part-1-study-fundamentals/)
- [Savills / GLA — Residential Development Margin](https://www.london.gov.uk/sites/default/files/app17_savills_residential_development_margin.pdf)
- [Finbri — Profit on Cost in Property Development Finance](https://www.finbri.co.uk/glossary/financial-ratios/profitability-ratios/profit-on-cost)
- [Duotax — Cost to Build a Townhouse in Australia](https://duotax.com.au/insights/cost-to-build-a-townhouse/)
- [Buildana — Duplex vs Townhouse Development in Sydney](https://www.buildana.com.au/insights/duplex-vs-townhouse-sydney)
- [Aprao — What is a sensitivity analysis and why does it matter](https://www.aprao.com/blog/what-is-a-sensitivity-analysis-and-why-does-it-matter)
