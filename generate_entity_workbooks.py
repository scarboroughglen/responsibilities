#!/usr/bin/env python3
"""Generate individual reserve funding Excel workbooks for each Scarborough Glen entity."""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from copy import copy

# ---------------------------------------------------------------------------
# Data from becht_report.md — all allocations, components, and disbursements
# ---------------------------------------------------------------------------

ENTITIES = {
    "HOA": {
        "name": "Scarborough Glen HOA",
        "filename": "HOA_Reserve_Plan.xlsx",
        "units": 136,
        "starting_balance": 59717,
        "baseline_contribution": 64563,
        "total_replacement_cost": 945673,
        "per_unit_floor": 350,  # ~5% of $6,953 per-unit replacement cost
        "components": [
            # (Component Name, Becht Total, Share%, Entity Allocated, Replacement Year, Inflated Cost)
            ("Aerator", 5500, 100.0, 5500, 2036, 7392),
            ("Asphalt Crack Filling", 15000, 100.0, 15000, "recurring", None),
            ("Asphalt Paving (Roads)", 330000, 100.0, 330000, 2036, 443487),
            ("Bathroom Refurbishment (Clubhouse)", 10000, 100.0, 10000, 2033, 12299),
            ("Clubhouse Furniture", 15000, 100.0, 15000, 2036, 20159),
            ("Concrete Pool Apron", 78250, 100.0, 78250, 2036, 105162),
            ("Concrete Pool Coping", 9900, 100.0, 9900, 2031, 11477),
            ("Concrete Sidewalks (HOA share)", 20103, 57.11, 20103, 2033, 24723),
            ("Curbing, Concrete", 35000, 100.0, 35000, 2036, 47037),
            ("Entrance Sign - Large", 20000, 100.0, 20000, 2046, 36122),
            ("Entrance Sign - Small", 12000, 100.0, 12000, 2046, 21673),
            ("Fence, Vinyl Stockade", 42500, 100.0, 42500, 2043, 70246),
            ("Fence, Wood Split Rail", 122500, 100.0, 122500, 2041, 190852),
            ("Fire Alarm Control Panel", 8000, 100.0, 8000, 2033, 9839),
            ("Fitness Room Refurbishment", 5000, 100.0, 5000, 2033, 6149),
            ("Gutters - Clubhouse", 2500, 100.0, 2500, 2028, 2652),
            ("Hot Water Heater", 2500, 100.0, 2500, "recurring", None),
            ("Kitchen Refurbishment", 3500, 100.0, 3500, 2033, 4305),
            ("Leaders - Clubhouse", 500, 100.0, 500, 2032, 597),
            ("Lights, Entry Clubhouse", 2700, 100.0, 2700, 2036, 3629),
            ("Lights, Entrance Sign", 1500, 100.0, 1500, 2036, 2016),
            ("Lights, Recessed", 5500, 100.0, 5500, 2033, 6764),
            ("Lights, Street", 83200, 100.0, 83200, 2033, 102325),
            ("Mailboxes", 20400, 100.0, 20400, 2049, 40261),
            ("Pool Filter System", 10000, 100.0, 10000, 2036, 13439),
            ("Roof, Shingles - Clubhouse", 15620, 100.0, 15620, 2032, 18650),
            ("Seal Coating", 16500, 100.0, 16500, "recurring", None),
            ("Siding, Vinyl - Clubhouse", 30000, 100.0, 30000, 2048, 57482),
            ("Skylights - Clubhouse", 4500, 100.0, 4500, 2032, 5373),
            ("Stop Signs", 2500, 100.0, 2500, 2046, 4515),
            ("Street Signs", 2000, 100.0, 2000, 2046, 3612),
            ("Windows, Dbl Hung Double - Clubhouse", 4000, 100.0, 4000, 2053, None),
            ("Windows, Dbl Hung Single - Clubhouse", 2000, 100.0, 2000, 2053, None),
            ("Windows, Transom", 1500, 100.0, 1500, 2053, None),
        ],
        # Year -> total disbursement (inflated) from Appendix D
        "disbursements": {
            2028: 36067,
            2031: 11477,
            2032: 24620,
            2033: 208222,
            2036: 642321,
            2038: 44915,
            2041: 190852,
            2043: 126440,
            2046: 65922,
            2048: 117838,
            2049: 40261,
        },
        # Disbursement detail: year -> [(component, amount)]
        "disbursement_detail": {
            2028: [("Asphalt Crack Filling", 15915), ("Seal Coating", 17500), ("Gutters - Clubhouse", 2652)],
            2031: [("Concrete Pool Coping", 11477)],
            2032: [("Roof - Clubhouse", 18650), ("Skylights - Clubhouse", 5373), ("Leaders - Clubhouse", 597)],
            2033: [("Lights, Street", 102325), ("Bathroom Refurbishment", 12299), ("Concrete Sidewalks (HOA)", 24723),
                   ("Fire Alarm Control Panel", 9839), ("Lights, Recessed", 6764), ("Fitness Room Refurb", 6149),
                   ("Hot Water Heater", 3070), ("Kitchen Refurbishment", 4305), ("Asphalt Crack Filling", 18444),
                   ("Seal Coating", 20304)],
            2036: [("Asphalt Paving (Roads)", 443487), ("Concrete Pool Apron", 105162), ("Curbing, Concrete", 47037),
                   ("Clubhouse Furniture", 20159), ("Pool Filter System", 13439), ("Aerator", 7392),
                   ("Lights, Entry Clubhouse", 3629), ("Lights, Entrance Sign", 2016)],
            2038: [("Asphalt Crack Filling", 21380), ("Seal Coating", 23535)],
            2041: [("Fence, Wood Split Rail", 190852)],
            2043: [("Fence, Vinyl Stockade", 70246), ("Asphalt Crack Filling", 24769), ("Seal Coating", 27274),
                   ("Hot Water Heater", 4151)],
            2046: [("Entrance Sign - Large", 36122), ("Entrance Sign - Small", 21673),
                   ("Stop Signs", 4515), ("Street Signs", 3612)],
            2048: [("Siding, Vinyl - Clubhouse", 57482), ("Asphalt Crack Filling", 28702), ("Seal Coating", 31654)],
            2049: [("Mailboxes", 40261)],
        },
        "omitted_items": [
            ("Playground Equipment", 27000, "Will need replacement; Falcon estimated $27,000"),
            ("Pool Shell Resurfacing", 22620, "Recurring ~10-year cycle; Falcon estimated $22,620/cycle"),
            ("Pool Fence, Aluminum", 18425, "Not in Becht; Falcon estimated $18,425"),
            ("Board-on-Board Fence", 37350, "Property-line fencing; Falcon estimated $37,350"),
            ("Roadway Granite Block Entry", 37875, "Decorative entry feature; Falcon estimated $37,875"),
            ("Clubhouse HVAC Split System", 9000, "Critical building system; Falcon estimated $9,000"),
            ("Clubhouse Restroom/Lockers", 30000, "Large item; Falcon estimated $30,000"),
            ("Pool Cover", 4000, "Recurring; Falcon estimated $4,000"),
            ("Pool Pump", 2500, "Recurring; Falcon estimated $2,500"),
            ("Pool Furniture Fund", 12000, "Recurring; Falcon estimated $12,000"),
            ("Pool Chlorination Equipment", 1500, "Recurring; Falcon estimated $1,500"),
            ("Clubhouse Flooring (carpet + tile)", 8280, "Falcon estimated $8,280 combined"),
            ("Clubhouse Door, Main Entry", 3500, "Falcon estimated $3,500"),
            ("Guard Rail", 10880, "Falcon estimated $10,880"),
            ("Vinyl on Masonry Wall Fence", 13950, "Falcon estimated $13,950"),
            ("Fire Suppression System", 7500, "Falcon estimated $7,500"),
            ("Irrigation Repair Fund", 5000, "Recurring; Falcon estimated $5,000"),
            ("CCTV Security", 2000, "Falcon estimated $2,000"),
            ("Key Fob Entry System", 1500, "Falcon estimated $1,500"),
        ],
    },
    "Condo_I": {
        "name": "Scarborough Glen Condo Association I",
        "filename": "Condo_I_Reserve_Plan.xlsx",
        "units": 18,
        "starting_balance": 29418,
        "baseline_contribution": 14383,
        "total_replacement_cost": 455954,
        "per_unit_floor": 1250,  # ~5% of $25,331 per-unit replacement cost
        "components": [
            ("Siding, Vinyl (Condo I share)", 1800000, 8.23, 148140, 2049, 292367),
            ("Deck Replacement, Wood (Condo I share)", 440000, 24.97, 109868, 2033, 135123),
            ("Deck Replacement, Composite (Condo I share)", 360000, 24.97, 89892, 2053, None),
            ("Roof, Shingles - Townhouses (Condo I share)", 1150988, 7.28, 83792, 2050, None),
            ("Chimney Chase Covers (Condo I share)", 178800, 10.00, 17880, 2031, 20728),
            ("Concrete Sidewalks (Condo I share)", 35200, 18.13, 6382, 2033, 7849),
        ],
        "disbursements": {
            2031: 20728,
            2033: 142972,
            2049: 292367,
        },
        "disbursement_detail": {
            2031: [("Chimney Chase Covers", 20728)],
            2033: [("Deck Replacement, Wood", 135123), ("Concrete Sidewalks", 7849)],
            2049: [("Siding, Vinyl", 292367)],
        },
        "omitted_items": [
            ("Entry Stoops (8 units)", 12000, "Concrete steps/landings at unit entries; Falcon estimated $12,000"),
            ("Wood Trim Replacement (recurring 3-yr cycle)", 5000, "Recurring maintenance; Falcon estimated $5,000/cycle"),
            ("Mailbox Hut Restoration", 1500, "Falcon estimated $1,500"),
            ("Walkway Pavers at Mailbox", 2200, "Falcon estimated $2,200"),
        ],
    },
    "Condo_II": {
        "name": "Scarborough Glen Condo Association II",
        "filename": "Condo_II_Reserve_Plan.xlsx",
        "units": 11,
        "starting_balance": 27319,
        "baseline_contribution": 18686,
        "total_replacement_cost": 423432,
        "per_unit_floor": 2000,  # ~5% of $38,494 per-unit replacement cost
        "components": [
            ("Siding, Wood (Condo II share)", 222000, 49.31, 109468, 2038, 156074),
            ("Deck Replacement, Wood (Condo II share)", 440000, 18.48, 81312, 2033, 100003),
            ("Deck Replacement, Composite (Condo II share)", 360000, 18.48, 66528, 2053, None),
            ("Roof, Shingles - Townhouses (Condo II share)", 1150988, 9.28, 106812, 2050, None),
            ("Asphalt Driveways (Condo II share)", 308000, 12.22, 37638, 2036, 50582),
            ("Chimney Chase Covers (Condo II share)", 178800, 11.00, 19668, 2031, 22801),
            ("Concrete Sidewalks (Condo II share)", 35200, 5.70, 2006, 2033, 2468),
        ],
        "disbursements": {
            2031: 22801,
            2033: 102471,
            2036: 50582,
            2038: 156074,
        },
        "disbursement_detail": {
            2031: [("Chimney Chase Covers", 22801)],
            2033: [("Deck Replacement, Wood", 100003), ("Concrete Sidewalks", 2468)],
            2036: [("Asphalt Driveways", 50582)],
            2038: [("Siding, Wood", 156074)],
        },
        "omitted_items": [
            ("Entry Stoops (11 units)", 44000, "Concrete steps/landings at unit entries; Falcon estimated $44,000"),
            ("Driveway Seal Coat (recurring 5-yr cycle)", 1400, "Recurring maintenance; Falcon estimated $1,400/cycle"),
        ],
    },
    "Condo_III": {
        "name": "Scarborough Glen Condo Association III",
        "filename": "Condo_III_Reserve_Plan.xlsx",
        "units": 9,
        "starting_balance": 24800,
        "baseline_contribution": 17450,
        "total_replacement_cost": 384386,
        "per_unit_floor": 2000,  # ~5% of $42,710 per-unit replacement cost
        "components": [
            ("Siding, Wood (Condo III share)", 222000, 50.69, 112532, 2038, 160442),
            ("Deck Replacement, Wood (Condo III share)", 440000, 19.75, 86900, 2033, 106876),
            ("Deck Replacement, Composite (Condo III share)", 360000, 19.75, 71100, 2053, None),
            ("Roof, Shingles - Townhouses (Condo III share)", 1150988, 6.98, 80339, 2050, None),
            ("Asphalt Driveways (Condo III share)", 308000, 5.13, 15800, 2036, 21234),
            ("Chimney Chase Covers (Condo III share)", 178800, 9.00, 16092, 2031, 18655),
            ("Concrete Sidewalks (Condo III share)", 35200, 4.61, 1623, 2033, 1996),
        ],
        "disbursements": {
            2031: 18655,
            2033: 108871,
            2036: 21234,
            2038: 160442,
        },
        "disbursement_detail": {
            2031: [("Chimney Chase Covers", 18655)],
            2033: [("Deck Replacement, Wood", 106876), ("Concrete Sidewalks", 1996)],
            2036: [("Asphalt Driveways", 21234)],
            2038: [("Siding, Wood", 160442)],
        },
        "omitted_items": [
            ("Entry Stoops (9 units)", 31500, "Concrete steps/landings at unit entries; Falcon estimated $31,500"),
            ("Wood Trim Replacement (recurring 5-yr cycle)", 7500, "Recurring maintenance; Falcon estimated $7,500/cycle"),
            ("Driveway Seal Coat (recurring 5-yr cycle)", 588, "Recurring maintenance; Falcon estimated $588/cycle"),
        ],
    },
    "Condo_IV": {
        "name": "Scarborough Glen Condo Association IV",
        "filename": "Condo_IV_Reserve_Plan.xlsx",
        "units": 98,
        "starting_balance": 207172,
        "baseline_contribution": 99619,
        "total_replacement_cost": 3211033,
        "per_unit_floor": 1500,  # ~5% of $32,765 per-unit replacement cost
        "components": [
            ("Siding, Vinyl (Condo IV share)", 1800000, 91.77, 1651860, 2049, 3260083),
            ("Roof, Shingles - Townhouses (Condo IV share)", 1150988, 76.46, 880045, 2050, None),
            ("Asphalt Driveways (Condo IV share)", 308000, 82.66, 254562, 2036, 342150),
            ("Deck Replacement, Wood (Condo IV share)", 440000, 36.79, 161876, 2033, 199087),
            ("Deck Replacement, Composite (Condo IV share)", 360000, 36.79, 132444, 2053, None),
            ("Chimney Chase Covers (Condo IV share)", 178800, 70.00, 125160, 2031, 145095),
            ("Concrete Sidewalks (Condo IV share)", 35200, 14.45, 5086, 2033, 6256),
        ],
        "disbursements": {
            2031: 145095,
            2033: 205342,
            2036: 342150,
            2049: 3260083,
        },
        "disbursement_detail": {
            2031: [("Chimney Chase Covers", 145095)],
            2033: [("Deck Replacement, Wood", 199087), ("Concrete Sidewalks", 6256)],
            2036: [("Asphalt Driveways", 342150)],
            2049: [("Siding, Vinyl", 3260083)],
        },
        "omitted_items": [
            ("Entry Stoops/Porches (recurring)", 10000, "End unit porches on 5-yr cycle; Falcon estimated $10,000/cycle"),
            ("Privacy Fence, Vinyl", 7735, "Condo-specific fencing; Falcon estimated $7,735"),
            ("Retaining Wall, Wood Tie Wells", 15000, "Structural/safety; Falcon estimated $15,000"),
            ("Driveway Seal Coat (recurring 5-yr cycle)", 9473, "Recurring maintenance; Falcon estimated $9,473/cycle"),
        ],
    },
}

YEARS = list(range(2026, 2051))

# ---------------------------------------------------------------------------
# Component useful-life lookup (Becht 2026 typical/remaining life)
# Used to compute the Fully Funded Balance (FFB) for each component:
#   FFB share = allocated cost x (effective age / useful life)
#   effective age = useful life - remaining life
# Ordered most-specific first; first substring match wins.
# ---------------------------------------------------------------------------

COMPONENT_LIFE = [
    # (name substring, useful life, remaining life)
    ("Railings", 12, 1),   # deck rails/balusters: safety cycle, past service life now
    ("Leveling", 7, 1),    # stoop leveling / stair nosing: safety repair cycle
    ("Entry Stoops", 28, 6),   # concrete stoop/landing replacement (F.2)
    ("Pool Shell", 10, 1),
    ("Pool Fence", 28, 10),
    ("Board-on-Board", 25, 8),
    ("Deck Replacement, Composite", 50, 27),
    ("Deck Replacement, Wood", 30, 7),
    ("Deck Boards", 18, 7),   # decking surface: ~18-yr wear cycle
    ("Deck", 30, 7),          # framing & footings (and generic deck)
    ("Siding, Vinyl - Clubhouse", 45, 22),
    ("Siding, Vinyl", 45, 23),
    ("Siding, Wood", 40, 12),
    ("Roof, Shingles - Clubhouse", 25, 6),
    ("Roof - Clubhouse", 25, 6),
    ("Roof", 25, 24),
    ("Chimney", 25, 5),
    ("Asphalt Crack", 3, 2),
    ("Asphalt Driveways", 20, 10),
    ("Asphalt Paving", 20, 10),
    ("Driveways", 20, 10),
    ("Seal Coating", 5, 2),
    ("Concrete Pool Apron", 30, 10),
    ("Concrete Pool Coping", 25, 5),
    ("Concrete Sidewalks", 30, 7),
    ("Curbing", 20, 10),
    ("Fence, Vinyl Stockade", 40, 17),
    ("Fence, Wood Split Rail", 25, 15),
    ("Mailboxes", 25, 23),
    ("Aerator", 15, 10),
    ("Bathroom Refurb", 30, 7),
    ("Clubhouse Furniture", 15, 10),
    ("Fire Alarm", 30, 7),
    ("Fitness Room", 30, 7),
    ("Gutters", 25, 2),
    ("Hot Water Heater", 10, 7),
    ("Kitchen Refurb", 25, 7),
    ("Leaders", 25, 6),
    ("Lights, Entry Clubhouse", 30, 10),
    ("Lights, Entrance Sign", 20, 10),
    ("Lights, Recessed", 30, 7),
    ("Lights, Street", 30, 7),
    ("Pool Filter", 20, 10),
    ("Skylights", 25, 6),
    ("Stop Signs", 25, 20),
    ("Street Signs", 25, 20),
    ("Windows", 50, 27),
    ("Entrance Sign", 30, 20),
]


def get_life(component_name):
    """Return (useful_life, remaining_life) for a component name, or (None, None)."""
    for substr, life, rem in COMPONENT_LIFE:
        if substr in component_name:
            return life, rem
    return None, None


# ---------------------------------------------------------------------------
# Three-fund configuration
#   Fund 1  Emergency Reserve      -> worst-case building replacement (target)
#   Fund 2  Long-Term Maintenance  -> Becht/Falcon capital schedule (never < $0)
#   Fund 3  Regular Reserve        -> decoupled general reserve (flat floor)
# emergency_target / regular_floor are placeholders the board replaces with real
# figures. total_cash is the entity's actual reserve cash to allocate at start.
# ---------------------------------------------------------------------------

# operating_multiplier x operating_monthly = the Operating Reserve floor (months of
# operating expenses). operating_monthly defaults to 0 (placeholder — enter actual).
FUND_DEFAULTS = {
    # HOA cash and operating figures are actuals from the Jan 1 - Jun 30 2026 GL and bank
    # register: reserve cash $364,333.66 (Webster Money Market 6412, rounded), operating
    # expense $287,561.51 over six months = $47,927/mo. NOTE: the half-year run rate is
    # seasonally high — it carries the full annual insurance premium, the whole winter snow
    # spend and both audits — so replace operating_monthly with the adopted annual operating
    # budget / 12 once available. The total contribution is insensitive to this: the floor
    # moves the answer by only ~$3,200/yr across a $35k-$48k/mo range.
    # start_split funds the Operating Reserve floor from cash on hand and seeds the
    # remainder to Long-Term Maintenance; Emergency is built from contributions.
    "HOA":      {"emergency_target": 200000, "operating_multiplier": 3, "operating_monthly": 47927,
                 "total_cash": 364334,
                 "start_split": (0, 220553, 143781),
                 "emergency_label": "largest common asset (clubhouse / pool)"},
    "Condo_I":  {"emergency_target": 150000, "operating_multiplier": 3, "operating_monthly": 0, "total_cash": 29418,
                 "emergency_label": "worst-case building replacement"},
    "Condo_II": {"emergency_target": 120000, "operating_multiplier": 3, "operating_monthly": 0, "total_cash": 27319,
                 "emergency_label": "worst-case building replacement"},
    "Condo_III": {"emergency_target": 120000, "operating_multiplier": 3, "operating_monthly": 0, "total_cash": 24800,
                  "emergency_label": "worst-case building replacement"},
    "Condo_IV": {"emergency_target": 300000, "operating_multiplier": 3, "operating_monthly": 36503, "total_cash": 145989,
                 "emergency_label": "worst-case building replacement",
                 # Recommended starting-balance split (Emergency, LTM, Operating); Emergency built over 10y
                 "start_split": (0, 95000, 50989), "emergency_fill_years": 10},
}
for _k, _v in FUND_DEFAULTS.items():
    ENTITIES[_k].update(_v)


def split_deck_rails(entity, rail_share=0.25, decking_share=0.35, rail_cycle=12,
                     inflation=0.03, horizon=2050, asap_year=2027):
    """Split the lumped wood deck into three components on their own cycles
    (per becht_report.md Appendix F.1):

      - Railings & Balusters (25%) — life-safety; must resist 200 lbs lateral.
        Past service life now, so replaced ASAP (asap_year), again with the deck
        rebuild, and every ~12 yrs thereafter.
      - Deck Boards / surface (35%) — ~18-yr wear cycle.
      - Framing & Footings (40%) — long-lived structure (survives 2-3 board cycles).

    Framing and boards are both renewed when the deck is rebuilt in Becht's due
    year (2033); their next stand-alone cycles (boards ~2051, framing ~2063) fall
    beyond the 2050 window, so only the railings recur in-horizon (2027/2033/2045)."""
    framing_share = 1 - rail_share - decking_share
    comps = entity["components"]
    wood = next((c for c in comps if "Deck Replacement, Wood" in c[0]), None)
    if wood is None:
        return
    name, becht, share, alloc, yr, infl = wood
    suffix = name[name.find(" ("):] if " (" in name else ""   # e.g. " (Condo IV share)"
    rail_cur = round(alloc * rail_share)
    deck_cur = round(alloc * decking_share)
    fram_cur = alloc - rail_cur - deck_cur

    def inf(cur, y):
        return round(cur * (1 + inflation) ** (y - 2026)) if isinstance(y, int) else infl

    fram_name = f"Deck Framing & Footings, Wood{suffix}"
    deck_name = f"Deck Boards (surface), Wood{suffix}"
    rail_name = f"Deck Railings & Balusters, Wood{suffix} (safety)"
    new_comps = []
    for c in comps:
        if c is wood:
            new_comps.append((fram_name, becht, share, fram_cur, yr, inf(fram_cur, yr)))
            new_comps.append((deck_name, becht, share, deck_cur, yr, inf(deck_cur, yr)))
            # rail line's headline date is the ASAP replacement (past service life)
            new_comps.append((rail_name, becht, share, rail_cur, asap_year, inf(rail_cur, asap_year)))
        else:
            new_comps.append(c)
    entity["components"] = new_comps

    # In the deck's replacement year, split the lumped deck into framing + boards + rails.
    detail = entity["disbursement_detail"]
    for y in list(detail):
        new_items = []
        for nm, amt in detail[y]:
            if "Deck Replacement, Wood" in nm:
                new_items.append(("Deck Framing & Footings", round(amt * framing_share)))
                new_items.append(("Deck Boards (surface)", round(amt * decking_share)))
                new_items.append(("Deck Railings & Balusters (safety)", round(amt * rail_share)))
            else:
                new_items.append((nm, amt))
        detail[y] = new_items

    # Rail-only replacement years: ASAP (overdue now) + mid-cycle after the deck job.
    rail_years = set()
    if asap_year and asap_year < yr:
        rail_years.add(asap_year)
    if isinstance(yr, int):
        ry = yr + rail_cycle
        while ry <= horizon:
            rail_years.add(ry)
            ry += rail_cycle
    for ry in rail_years:
        amt = round(rail_cur * (1 + inflation) ** (ry - 2026))
        detail.setdefault(ry, []).append(("Deck Railings & Balusters (safety)", amt))
    entity["disbursements"] = {y: sum(a for _, a in items) for y, items in detail.items()}


for _k in ("Condo_I", "Condo_II", "Condo_III", "Condo_IV"):
    split_deck_rails(ENTITIES[_k])


# Appendix F unbundling: recurring / safety items to fund as their own LTM lines.
# (name, current cost, cycle years, first year, basis note). Falcon $ where it
# exists; "placeholder" where no source figure is available.
APPENDIX_F_ADDITIONS = {
    "HOA": [
        ("Pool Shell Resurfacing", 22620, 10, 2027, "Falcon"),
        ("Pool Mechanical (pump/chlorinator/heater)", 4000, 8, 2028, "Falcon (pump+chlor)"),
        ("Pool Cover", 4000, 9, 2028, "Falcon"),
        ("Pool Furniture", 12000, 10, 2030, "Falcon"),
        ("Pool Fence / Safety Barrier", 18425, 28, 2035, "Falcon (aluminum)"),
        ("Split-Rail Fence — Repair Fund", 3750, 4, 2027, "Falcon"),
        ("Board-on-Board Fence (property line)", 37350, 25, 2034, "Falcon"),
    ],
    "Condo_I": [
        ("Entry Stoops — Concrete Replacement", 12000, 28, 2032, "Falcon"),
        ("Entry Stoops — Leveling & Stair Nosing (safety)", 1800, 7, 2028, "placeholder ~15% of replacement"),
        ("Wood / Window Trim Repair", 5000, 4, 2027, "Falcon"),
    ],
    "Condo_II": [
        ("Entry Stoops — Concrete Replacement", 44000, 28, 2032, "Falcon"),
        ("Entry Stoops — Leveling & Stair Nosing (safety)", 6600, 7, 2028, "placeholder ~15%"),
        ("Exterior Paint / Stain / Caulk", 8000, 6, 2029, "placeholder"),
    ],
    "Condo_III": [
        ("Entry Stoops — Concrete Replacement", 31500, 28, 2032, "Falcon"),
        ("Entry Stoops — Leveling & Stair Nosing (safety)", 4725, 7, 2028, "placeholder ~15%"),
        ("Wood / Window Trim Repair", 7500, 5, 2027, "Falcon"),
        ("Exterior Paint / Stain / Caulk", 8000, 6, 2029, "placeholder"),
    ],
    "Condo_IV": [
        ("Entry Stoops — Concrete Replacement", 196000, 28, 2032,
         "$2,000/unit x 98 — MAJOR item missed by Falcon AND Becht"),
        ("Entry Stoops — Leveling & Stair Nosing (safety)", 29400, 7, 2028, "placeholder ~15%"),
        ("Exterior Caulk/Sealant & Trim Maintenance", 15000, 6, 2027,
         "window-seal caulking (critical) + trim; protects the wall assembly behind the vinyl"),
    ],
}

# Omitted-list entries that are now funded in LTM (pruned to avoid double-listing).
_F_PRUNE_KEYS = ["Stoop", "Trim", "Pool Shell", "Pool Cover", "Pool Pump",
                 "Pool Furniture", "Chlorination", "Board-on-Board", "Pool Fence"]


def apply_appendix_f(entity, key, inflation=0.03, horizon=2050):
    """Fold Appendix F recurring/safety items into the funded LTM schedule as their
    own line items, each on its own cycle. Per becht_report.md Appendix F."""
    items = APPENDIX_F_ADDITIONS.get(key, [])
    if not items:
        return
    detail = entity["disbursement_detail"]
    for name, cost, cycle, first, _note in items:
        first_infl = round(cost * (1 + inflation) ** (first - 2026))
        entity["components"].append((name, None, 100.0, cost, first, first_infl))
        y = first
        while y <= horizon:
            detail.setdefault(y, []).append((name, round(cost * (1 + inflation) ** (y - 2026))))
            y += cycle
    entity["disbursements"] = {yr: sum(a for _, a in its) for yr, its in detail.items()}
    entity["total_replacement_cost"] = sum(c[3] for c in entity["components"])
    entity["omitted_items"] = [it for it in entity.get("omitted_items", [])
                               if not any(k in it[0] for k in _F_PRUNE_KEYS)]


for _k in ENTITIES:
    apply_appendix_f(ENTITIES[_k], _k)


def compute_ffb(entity):
    """Fully Funded Balance = sum(allocated cost x effective_age/useful_life)."""
    ffb = 0.0
    for name, _becht, _share, allocated, _yr, _infl in entity["components"]:
        life, rem = get_life(name)
        if life is None:
            continue
        frac = max(0.0, min(1.0, (life - rem) / life))
        ffb += allocated * frac
    return ffb


FUND_FILL_YEARS = 5   # years to build Emergency / Regular up to their targets


def _roundup(x, step=100):
    import math
    return int(math.ceil(x / step)) * step


def solve_ltm_contribution(start, disb, growth=0.03, interest=0.01):
    """Min year-1 contribution (growing) so LTM closing never drops below $0."""
    lo, hi = 0.0, 2_000_000.0
    for _ in range(60):
        c1 = (lo + hi) / 2
        bal, c, ok = start, c1, True
        for i, y in enumerate(YEARS):
            if i > 0:
                c *= (1 + growth)
            bal = bal + c - disb.get(y, 0) + bal * interest
            if bal < -0.5:
                ok = False
                break
        if ok:
            hi = c1
        else:
            lo = c1
    return hi


def solve_fund_to(start, target, years, growth=0.03, interest=0.01):
    """Year-1 contribution (growing) to reach `target` in `years` years."""
    if start >= target:
        return 0.0
    lo, hi = 0.0, float(target)
    for _ in range(60):
        c1 = (lo + hi) / 2
        bal, c = start, c1
        for i in range(years):
            if i > 0:
                c *= (1 + growth)
            bal = bal + c + bal * interest
        if bal >= target:
            hi = c1
        else:
            lo = c1
    return hi


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
LABEL_FONT = Font(name="Calibri", bold=True, size=11)
TITLE_FONT = Font(name="Calibri", bold=True, size=14)
SECTION_FONT = Font(name="Calibri", bold=True, size=12, color="2F5496")
NORMAL_FONT = Font(name="Calibri", size=11)
SMALL_FONT = Font(name="Calibri", size=10, italic=True, color="666666")
WARN_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
WARN_FONT = Font(name="Calibri", color="9C0006", bold=True)
GOOD_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
GOOD_FONT = Font(name="Calibri", color="006100", bold=True)
NEUTRAL_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
CURRENCY_FMT = '#,##0'
PCT_FMT = '0.00%'
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def style_header_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN_BORDER


def style_data_cell(cell, fmt=None):
    cell.font = NORMAL_FONT
    cell.border = THIN_BORDER
    if fmt:
        cell.number_format = fmt


def auto_width(ws, min_width=10, max_width=40):
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def build_instructions_sheet(ws, entity):
    ws.sheet_properties.tabColor = "2F5496"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 45
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 20

    row = 1
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
    ws.cell(row=row, column=2, value=entity["name"]).font = TITLE_FONT
    row += 1
    ws.cell(row=row, column=2, value="Capital Reserve Funding Plan").font = SECTION_FONT
    row += 2

    # How to read this workbook (orientation for board members opening the file)
    ws.cell(row=row, column=2, value="How to read this workbook").font = SECTION_FONT
    row += 1
    howto = [
        "START on the 'Fund Summary' tab — it shows the three funds, their targets/floors, the",
        "per-unit monthly funding, and the 'Reduction Readiness' gates all on one page.",
        "The three fund tabs — Emergency Reserve, Long-Term Maintenance, Operating Reserve — each",
        "show that fund's year-by-year balance from 2026 to 2050.",
        "'Components' lists every capital item and its cost; 'Expenditure Chart' plots spending by year.",
        "Yellow cells are inputs you can edit; every other cell is calculated automatically.",
    ]
    for line in howto:
        ws.cell(row=row, column=2, value=line).font = NORMAL_FONT
        row += 1
    row += 1

    ws.cell(row=row, column=2, value="Data Sources:").font = LABEL_FONT
    row += 1
    ws.cell(row=row, column=2,
            value="Based on Becht Engineering Capital Reserve Study (Project 25-1140, February 2026)").font = NORMAL_FONT
    row += 1
    ws.cell(row=row, column=2,
            value="Entity allocations derived from Falcon Group Report (2022) per-entity quantity ratios").font = NORMAL_FONT
    row += 2

    # Overview
    ws.cell(row=row, column=2, value="Overview").font = SECTION_FONT
    row += 1
    info = [
        f"Number of units: {entity['units']}",
        f"Total replacement cost (current dollars): ${entity['total_replacement_cost']:,.0f}",
        f"Total reserve cash to allocate (2026): ${entity.get('total_cash', entity['starting_balance']):,.0f}",
        "Fund contributions are solved automatically (see Fund Summary) and are fully editable.",
        "Planning horizon: 2026-2050 (25 years)",
    ]
    for line in info:
        ws.cell(row=row, column=2, value=line).font = NORMAL_FONT
        row += 1
    row += 1

    # Funding model
    ws.cell(row=row, column=2, value="Funding Model Assumptions").font = SECTION_FONT
    row += 1
    assumptions = [
        "Annual contributions grow at a configurable rate (default 3%) to keep pace with inflation.",
        "Reserve funds earn interest at a configurable rate (default 1%) annually.",
        "Disbursements (expenditures) occur in the year projected by the Becht engineering study.",
        "All disbursement amounts include 3% annual inflation applied to current replacement costs.",
    ]
    for line in assumptions:
        ws.cell(row=row, column=2, value=line).font = NORMAL_FONT
        row += 1
    row += 1

    # Three-fund model
    ws.cell(row=row, column=2, value="Three-Fund Model").font = SECTION_FONT
    row += 1
    fund_lines = [
        "This plan splits reserves into THREE decoupled funds, each with its own balance,",
        "annual contribution, floor, and monthly cost per unit. Each has its own tab.",
        "",
        "1. Emergency Reserve — self-funded buffer to replace the worst-case building.",
        "   Floor = a target you set (Fund Summary, yellow cell). Not drawn for routine work.",
        "2. Long-Term Maintenance — the Becht/Falcon capital replacement schedule.",
        "   Floor = must NEVER drop below $0 in any year.",
        "3. Operating Reserve — decoupled buffer of operating expenses.",
        "   Floor = multiplier (1-3 months) x monthly operating expense, both set on Fund Summary.",
        "",
        "Fund Summary sheet: set all inputs here (each fund's starting balance, target/floor and",
        "contribution; unit count; growth; interest). Allocate current cash across the three funds'",
        "starting balances (Allocated Start column). Recommended: fund near-term Long-Term Maintenance",
        "needs and Operating liquidity first, and build the Emergency Reserve from contributions.",
        "",
        "Percent Funded (Long-Term Maintenance) = reserve on hand / Fully Funded Balance. It is",
        "the industry (CAI) metric: under 30% = weak (high special-assessment risk), 30-70% = fair,",
        "70%+ = strong.",
        "",
        "Reduction Readiness — common charges should NOT be reduced until all three gates PASS:",
        "  Gate 1 — Emergency Reserve is funded to its target.",
        "  Gate 2 — Long-Term Maintenance stays at/above $0 every year at the proposed contribution.",
        "  Gate 3 — Operating Reserve stays at/above its floor (multiplier x monthly operating expense).",
        "The 'OVERALL' cell reads ELIGIBLE only when all three pass. To model a reduction, lower a",
        "fund's Year-1 Contribution on the Fund Summary sheet and watch the gates.",
        "Replace every yellow placeholder (targets, floors, total cash) with the entity's ACTUALS.",
    ]
    for line in fund_lines:
        ws.cell(row=row, column=2, value=line).font = NORMAL_FONT
        row += 1
    row += 1

    # Components summary
    ws.cell(row=row, column=2, value="Components Covered by This Plan").font = SECTION_FONT
    row += 1
    ws.cell(row=row, column=2, value="See the 'Components' sheet for full detail. Summary:").font = NORMAL_FONT
    row += 2
    # Header row
    ws.cell(row=row, column=2, value="Component").font = LABEL_FONT
    ws.cell(row=row, column=3, value="Allocated Cost").font = LABEL_FONT
    ws.cell(row=row, column=4, value="Replacement Year").font = LABEL_FONT
    row += 1
    for comp in entity["components"]:
        name, _, share_pct, allocated, repl_year, _ = comp
        ws.cell(row=row, column=2, value=name).font = NORMAL_FONT
        ws.cell(row=row, column=3, value=allocated).font = NORMAL_FONT
        ws.cell(row=row, column=3).number_format = CURRENCY_FMT
        year_val = repl_year if isinstance(repl_year, int) else None
        if year_val:
            ws.cell(row=row, column=4, value=year_val).font = NORMAL_FONT
        row += 1
    row += 1

    # Omitted items
    if entity.get("omitted_items"):
        ws.cell(row=row, column=2,
                value="Items NOT Included (Budget Separately)").font = SECTION_FONT
        row += 1
        ws.cell(row=row, column=2,
                value="The following items from the Falcon 2022 report are NOT in the Becht 2026 study "
                      "and are therefore NOT included in this funding schedule. Consider budgeting "
                      "for these items separately.").font = NORMAL_FONT
        row += 2
        # Header
        ws.cell(row=row, column=2, value="Item").font = LABEL_FONT
        ws.cell(row=row, column=3, value="Falcon Est. Cost").font = LABEL_FONT
        ws.cell(row=row, column=4, value="Notes").font = LABEL_FONT
        row += 1
        total_omitted = 0
        for item_name, cost, notes in entity["omitted_items"]:
            ws.cell(row=row, column=2, value=item_name).font = NORMAL_FONT
            ws.cell(row=row, column=3, value=cost).font = NORMAL_FONT
            ws.cell(row=row, column=3).number_format = CURRENCY_FMT
            ws.cell(row=row, column=4, value=notes).font = SMALL_FONT
            total_omitted += cost
            row += 1
        ws.cell(row=row, column=2, value="TOTAL OMITTED").font = LABEL_FONT
        ws.cell(row=row, column=3, value=total_omitted).font = LABEL_FONT
        ws.cell(row=row, column=3).number_format = CURRENCY_FMT
        row += 2

    ws.cell(row=row, column=2,
            value="Prepared June 2026. Review and update annually.").font = SMALL_FONT


def build_components_sheet(ws, entity):
    ws.sheet_properties.tabColor = "BF8F00"

    ws.cell(row=1, column=1, value=f"{entity['name']} — Component Detail").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="All costs from Becht Engineering Capital Reserve Study (2026)").font = SMALL_FONT

    header_row = 4
    headers = [
        "Component", "Becht Community Total", "Entity Share %",
        "Entity Allocated Cost", "Replacement Year", "Inflated Cost at Replacement"
    ]
    for col, hdr in enumerate(headers, 1):
        ws.cell(row=header_row, column=col, value=hdr)
    style_header_row(ws, header_row, len(headers))

    for i, comp in enumerate(entity["components"]):
        r = header_row + 1 + i
        name, becht_total, share_pct, allocated, repl_year, inflated = comp

        ws.cell(row=r, column=1, value=name)
        style_data_cell(ws.cell(row=r, column=1))

        ws.cell(row=r, column=2, value=becht_total)
        style_data_cell(ws.cell(row=r, column=2), CURRENCY_FMT)

        ws.cell(row=r, column=3, value=share_pct / 100)
        style_data_cell(ws.cell(row=r, column=3), PCT_FMT)

        ws.cell(row=r, column=4, value=allocated)
        style_data_cell(ws.cell(row=r, column=4), CURRENCY_FMT)

        if isinstance(repl_year, int):
            year_display = repl_year if repl_year <= 2050 else f"{repl_year} (beyond plan)"
        else:
            year_display = str(repl_year)
        ws.cell(row=r, column=5, value=year_display)
        style_data_cell(ws.cell(row=r, column=5))

        if inflated is not None:
            ws.cell(row=r, column=6, value=inflated)
            style_data_cell(ws.cell(row=r, column=6), CURRENCY_FMT)
        else:
            ws.cell(row=r, column=6, value="Beyond plan horizon")
            style_data_cell(ws.cell(row=r, column=6))

    # Totals
    total_row = header_row + 1 + len(entity["components"])
    ws.cell(row=total_row, column=1, value="TOTAL").font = LABEL_FONT
    ws.cell(row=total_row, column=1).border = THIN_BORDER
    ws.cell(row=total_row, column=4, value=entity["total_replacement_cost"])
    ws.cell(row=total_row, column=4).font = LABEL_FONT
    ws.cell(row=total_row, column=4).number_format = CURRENCY_FMT
    ws.cell(row=total_row, column=4).border = THIN_BORDER

    col_widths = [45, 22, 14, 22, 18, 26]
    for i, w in enumerate(col_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    ws.freeze_panes = f"A{header_row + 1}"


COMPONENT_COLORS = [
    "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5",
    "70AD47", "264478", "9B57A0", "636363", "EB7E30",
    "44546A", "BF4B28", "00B0F0", "92D050", "7030A0",
    "C55A11", "2E75B6", "AFABAB", "43682B", "D63384",
    "F4B183", "8FAADC", "A9D18E", "FFD966", "B4C7E7",
    "F8CBAD", "C9C9C9", "E2F0D9", "D6DCE4", "FBE5D6",
]


def build_chart_sheet(ws, entity):
    ws.sheet_properties.tabColor = "C00000"

    detail_data = entity.get("disbursement_detail", {})
    all_components = set()
    for year_details in detail_data.values():
        for comp_name, _ in year_details:
            all_components.add(comp_name)
    all_components = sorted(all_components)

    if not all_components:
        ws.cell(row=1, column=1, value="No disbursement detail available.").font = NORMAL_FONT
        return

    # --- Color key table (top-left) ---
    ws.cell(row=1, column=1, value="COLOR KEY").font = SECTION_FONT
    ws.cell(row=2, column=1, value="Color").font = LABEL_FONT
    ws.cell(row=2, column=2, value="Component").font = LABEL_FONT
    ws.cell(row=2, column=3, value="25-Year Total").font = LABEL_FONT
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 16

    # Sum each component's total across all years for the key table
    comp_totals = {}
    for year_items in detail_data.values():
        for comp_name, amt in year_items:
            comp_totals[comp_name] = comp_totals.get(comp_name, 0) + amt

    for j, comp in enumerate(all_components):
        r = 3 + j
        color = COMPONENT_COLORS[j % len(COMPONENT_COLORS)]
        cell_color = ws.cell(row=r, column=1)
        cell_color.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell_color.border = THIN_BORDER
        ws.cell(row=r, column=2, value=comp).font = NORMAL_FONT
        ws.cell(row=r, column=2).border = THIN_BORDER
        total_val = comp_totals.get(comp, 0)
        ws.cell(row=r, column=3, value=total_val).font = NORMAL_FONT
        ws.cell(row=r, column=3).number_format = CURRENCY_FMT
        ws.cell(row=r, column=3).border = THIN_BORDER

    # --- Data table for chart (to the right of the color key) ---
    data_start_col = 5  # Column E
    years_with_data = sorted(y for y in YEARS if y in detail_data)
    num_data_rows = len(years_with_data)
    num_components = len(all_components)

    ws.cell(row=1, column=data_start_col, value="CHART DATA").font = SECTION_FONT
    ws.cell(row=2, column=data_start_col, value="Year").font = LABEL_FONT
    for j, comp in enumerate(all_components):
        ws.cell(row=2, column=data_start_col + 1 + j, value=comp).font = LABEL_FONT
        ws.column_dimensions[get_column_letter(data_start_col + 1 + j)].width = 16

    for i, year in enumerate(years_with_data):
        r = 3 + i
        ws.cell(row=r, column=data_start_col, value=year)
        year_items = {name: amt for name, amt in detail_data.get(year, [])}
        for j, comp in enumerate(all_components):
            val = year_items.get(comp)
            if val:
                ws.cell(row=r, column=data_start_col + 1 + j, value=val)
                ws.cell(row=r, column=data_start_col + 1 + j).number_format = CURRENCY_FMT

    # --- Stacked bar chart (kept to standard settings so Excel opens it cleanly;
    # avoids hand-built axis text-rotation XML and non-standard number formats that
    # trigger Excel's "recover content" repair prompt) ---
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = entity["name"] + " — Projected Expenditures by Year"
    chart.style = 10
    chart.width = 36
    chart.height = 20
    chart.legend = None
    chart.y_axis.title = "Cost ($)"
    chart.y_axis.numFmt = '$#,##0'
    chart.y_axis.axPos = 'l'
    chart.x_axis.title = "Year"
    chart.x_axis.numFmt = '0'
    chart.x_axis.axPos = 'b'

    cats = Reference(ws, min_col=data_start_col, min_row=3, max_row=2 + num_data_rows)
    for j in range(num_components):
        col = data_start_col + 1 + j
        data = Reference(ws, min_col=col, min_row=2, max_row=2 + num_data_rows)
        chart.add_data(data, titles_from_data=True)
        series = chart.series[j]
        color = COMPONENT_COLORS[j % len(COMPONENT_COLORS)]
        series.graphicalProperties.solidFill = color

    chart.set_categories(cats)

    # Place chart below both the color key and data table
    chart_start_row = max(3 + num_components, 3 + num_data_rows) + 2
    ws.add_chart(chart, f"A{chart_start_row}")

    ws.column_dimensions[get_column_letter(data_start_col)].width = 8


# ---------------------------------------------------------------------------
# Three-fund builders (supersede build_schedule_sheet / build_health_sheet)
# ---------------------------------------------------------------------------

# Fund-schedule geometry (shared by all three fund tabs)
FUND_HEADER_ROW = 4
FUND_FIRST_ROW = 5
FUND_LAST_ROW = FUND_FIRST_ROW + len(YEARS) - 1   # 29
FUND_TOTALS_ROW = FUND_LAST_ROW + 1               # 30


def build_fund_schedule(ws, entity, fund):
    """A single fund's 25-year schedule. All inputs live on the Fund Summary
    sheet, so this tab is purely calculated.

    fund keys: title, tab, start_ref, contrib_ref, growth_ref, interest_ref,
               units_ref, floor_ref, disb (dict), floor_kind ('zero'|'flat'|'target')
    """
    ws.sheet_properties.tabColor = fund["tab"]
    ws.cell(row=1, column=1, value=fund["title"]).font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="All inputs are set on the Fund Summary sheet. This tab is calculated.").font = SMALL_FONT

    header_row, first, last, totals = FUND_HEADER_ROW, FUND_FIRST_ROW, FUND_LAST_ROW, FUND_TOTALS_ROW
    headers = ["Year", "Opening Balance", "Annual Contribution", "Disbursement",
               "Interest Earned", "Closing Balance", "Monthly Cost/Unit", "Status"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=header_row, column=c, value=h)
    style_header_row(ws, header_row, len(headers))

    growth, interest, units = fund["growth_ref"], fund["interest_ref"], fund["units_ref"]
    floor, kind, disb = fund["floor_ref"], fund["floor_kind"], fund["disb"]

    for i, year in enumerate(YEARS):
        r = first + i
        ws.cell(row=r, column=1, value=year)
        style_data_cell(ws.cell(row=r, column=1))
        ws.cell(row=r, column=2).value = f"={fund['start_ref']}" if i == 0 else f"=F{r - 1}"
        # Target-based funds (Emergency, Operating) stop contributing once the fund has
        # reached its target — otherwise the contribution compounds forever and the fund
        # overshoots by many multiples of the target. Long-Term Maintenance (kind 'zero')
        # contributes every year by design: it is funding a continuing disbursement
        # schedule, not filling a fixed bucket.
        # The growing Year-1 contribution is re-derived from contrib_ref rather than the
        # prior year's cell, so that contributions resume correctly if a draw ever takes
        # a target fund back below its floor.
        if kind == "zero":
            ws.cell(row=r, column=3).value = (
                f"={fund['contrib_ref']}" if i == 0 else f"=C{r - 1}*(1+{growth})")
        elif i == 0:
            ws.cell(row=r, column=3).value = (
                f"=IF({fund['start_ref']}>={floor},0,{fund['contrib_ref']})")
        else:
            ws.cell(row=r, column=3).value = (
                f"=IF(F{r - 1}>={floor},0,{fund['contrib_ref']}*(1+{growth})^{i})")
        ws.cell(row=r, column=4, value=disb.get(year, 0))
        ws.cell(row=r, column=5).value = f"=B{r}*{interest}"
        ws.cell(row=r, column=6).value = f"=B{r}+C{r}-D{r}+E{r}"
        ws.cell(row=r, column=7).value = f"=C{r}/{units}/12"
        if kind == "zero":
            ws.cell(row=r, column=8).value = f'=IF(F{r}<0,"BELOW $0","")'
        elif kind == "flat":
            ws.cell(row=r, column=8).value = f'=IF(F{r}<{floor},"BELOW FLOOR","")'
        else:  # target
            ws.cell(row=r, column=8).value = f'=IF(F{r}>={floor},"FUNDED","BUILDING")'
        for col, fmt in [(2, CURRENCY_FMT), (3, CURRENCY_FMT), (4, CURRENCY_FMT),
                         (5, CURRENCY_FMT), (6, CURRENCY_FMT), (7, '$#,##0')]:
            cc = ws.cell(row=r, column=col)
            cc.number_format = fmt
            cc.border = THIN_BORDER
        sc = ws.cell(row=r, column=8)
        sc.border = THIN_BORDER
        sc.alignment = Alignment(horizontal="center")

    ws.cell(row=totals, column=1, value="TOTAL").font = LABEL_FONT
    ws.cell(row=totals, column=1).border = THIN_BORDER
    for col in (3, 4, 5):
        cl = get_column_letter(col)
        cell = ws.cell(row=totals, column=col, value=f"=SUM({cl}{first}:{cl}{last})")
        cell.font = LABEL_FONT
        cell.number_format = CURRENCY_FMT
        cell.border = THIN_BORDER

    from openpyxl.formatting.rule import FormulaRule
    rng = f"F{first}:F{last}"
    if kind == "zero":
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f"F{first}<0"], fill=WARN_FILL, font=WARN_FONT))
    else:
        # Excel rejects cross-sheet references INSIDE conditional-formatting formulas, so
        # mirror the floor (which lives on Fund Summary) into a local cell (J2) and let the
        # conditional formatting compare against that same-sheet cell.
        ws.cell(row=1, column=10, value="Floor (from Fund Summary):").font = SMALL_FONT
        fc = ws.cell(row=2, column=10, value=f"={floor}")
        fc.number_format = CURRENCY_FMT
        fc.font = SMALL_FONT
        ws.column_dimensions["J"].width = 26
        floor_local = "$J$2"
        if kind == "flat":
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"F{first}<{floor_local}"], fill=WARN_FILL, font=WARN_FONT))
        else:  # target
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"F{first}>={floor_local}"], fill=GOOD_FILL, font=GOOD_FONT))
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f"F{first}<{floor_local}"], fill=NEUTRAL_FILL))

    for i, w in enumerate([8, 18, 20, 18, 16, 18, 18, 16]):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    ws.freeze_panes = f"A{first}"


def build_summary_sheet(ws, entity):
    """Fund Summary & Reduction Readiness — shared inputs, the three funds with
    fill-by-priority allocation, LTM funding health, and the three gates."""
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    ws.sheet_properties.tabColor = "2F5496"

    units = entity["units"]
    total_cash = entity.get("total_cash", entity["starting_balance"])
    et = entity["emergency_target"]
    op_mult = entity.get("operating_multiplier", 3)
    op_monthly = entity.get("operating_monthly", 0)
    rf = op_monthly * op_mult   # Operating Reserve floor = months x monthly op expense
    ffb = compute_ffb(entity)
    elabel = entity.get("emergency_label", "worst-case building replacement")

    # Manual starting-balance allocation across the three funds (editable on the sheet).
    split = entity.get("start_split")
    if split:
        emerg_start, ltm_start, reg_start = split
    else:
        emerg_start, ltm_start, reg_start = 0, total_cash, 0   # default: seed LTM (near-term needs)
    em_years = entity.get("emergency_fill_years", FUND_FILL_YEARS)
    # Solve for contributions that actually fund each fund (LTM never below $0;
    # Emergency to target over em_years; Operating to its floor over FUND_FILL_YEARS).
    ltm_contrib = _roundup(solve_ltm_contribution(ltm_start, entity["disbursements"]))
    emerg_contrib = _roundup(solve_fund_to(emerg_start, et, em_years))
    reg_contrib = _roundup(solve_fund_to(reg_start, rf, FUND_FILL_YEARS))

    ws.cell(row=1, column=1, value=f"{entity['name']} — Fund Summary & Reduction Readiness").font = TITLE_FONT
    ws.cell(row=2, column=1,
            value="Three decoupled funds. Yellow cells are inputs — replace placeholders with actual figures.").font = SMALL_FONT

    # --- Shared inputs ---
    ws.cell(row=4, column=1, value="SHARED INPUTS").font = SECTION_FONT

    def inp(r, label, val, fmt):
        ws.cell(row=r, column=1, value=label).font = LABEL_FONT
        c = ws.cell(row=r, column=2, value=val)
        c.font = LABEL_FONT
        c.number_format = fmt
        c.fill = INPUT_FILL
        c.border = THIN_BORDER

    # Total Reserve Cash is computed = sum of the three funds' starting balances (D12:D14)
    ws.cell(row=5, column=1, value="Total Reserve Cash (all accounts):").font = LABEL_FONT
    c = ws.cell(row=5, column=2, value="=SUM(D12:D14)")
    c.font = LABEL_FONT
    c.number_format = CURRENCY_FMT
    c.border = THIN_BORDER
    ws.cell(row=5, column=3, value="Auto-total of the three fund starting balances (Allocated Start column).").font = SMALL_FONT
    inp(6, "Number of Units:", units, "#,##0")
    inp(7, "Contribution Growth Rate:", 0.03, PCT_FMT)
    inp(8, "Interest Rate on Reserves:", 0.01, PCT_FMT)

    # --- Three funds table ---
    ws.cell(row=10, column=1, value="THE THREE FUNDS").font = SECTION_FONT
    for c, h in enumerate(["Fund", "Target / Floor", "Year-1 Contribution", "Allocated Start", "Floor Rule"], 1):
        ws.cell(row=11, column=c, value=h)
    style_header_row(ws, 11, 5)

    def fund_row(r, name, target_val, target_is_input, contrib_val, alloc_value, rule):
        ws.cell(row=r, column=1, value=name).font = NORMAL_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        c = ws.cell(row=r, column=2, value=target_val)
        c.font = LABEL_FONT
        c.number_format = CURRENCY_FMT
        c.border = THIN_BORDER
        if target_is_input:
            c.fill = INPUT_FILL
        c = ws.cell(row=r, column=3, value=contrib_val)
        c.font = LABEL_FONT
        c.number_format = CURRENCY_FMT
        c.fill = INPUT_FILL
        c.border = THIN_BORDER
        c = ws.cell(row=r, column=4, value=alloc_value)   # starting balance (editable input)
        c.font = LABEL_FONT
        c.number_format = CURRENCY_FMT
        c.fill = INPUT_FILL
        c.border = THIN_BORDER
        ws.cell(row=r, column=5, value=rule).font = SMALL_FONT
        ws.cell(row=r, column=5).border = THIN_BORDER

    fund_row(12, "1. Emergency Reserve", et, True, emerg_contrib,
             emerg_start, f"Reach & hold target ({elabel})")
    fund_row(13, "2. Long-Term Maintenance", 0, False, ltm_contrib,
             ltm_start, "Never below $0 (Becht capital schedule)")
    # Operating Reserve floor is computed = multiplier x monthly operating expense (B38*B37)
    fund_row(14, "3. Operating Reserve", "=B38*B37", False, reg_contrib,
             reg_start, "≥ multiplier x monthly op. expense (set below)")

    ws.cell(row=15, column=1, value="TOTAL").font = LABEL_FONT
    ws.cell(row=15, column=1).border = THIN_BORDER
    for col, formula in [(3, "=SUM(C12:C14)"), (4, "=SUM(D12:D14)")]:
        c = ws.cell(row=15, column=col, value=formula)
        c.font = LABEL_FONT
        c.number_format = CURRENCY_FMT
        c.border = THIN_BORDER
    ws.cell(row=16, column=1,
            value="Allocated Start = each fund's current balance (yellow, editable); Total Reserve Cash (B5) "
                  "is their sum. Recommended: fund near-term Long-Term Maintenance and Operating liquidity "
                  "first, and build the Emergency Reserve from contributions over time.").font = SMALL_FONT

    # --- LTM funding health ---
    ws.cell(row=18, column=1, value="LONG-TERM MAINTENANCE — FUNDING HEALTH").font = SECTION_FONT
    ws.cell(row=19, column=1, value="Fully Funded Balance (FFB — what LTM should hold today):").font = LABEL_FONT
    c = ws.cell(row=19, column=2, value=round(ffb))
    c.font = LABEL_FONT
    c.number_format = CURRENCY_FMT
    c.border = THIN_BORDER
    ws.cell(row=20, column=1, value="LTM Reserve (allocated start):").font = LABEL_FONT
    c = ws.cell(row=20, column=2, value="=D13")
    c.number_format = CURRENCY_FMT
    c.border = THIN_BORDER
    ws.cell(row=21, column=1, value="Percent Funded:").font = LABEL_FONT
    c = ws.cell(row=21, column=2, value="=IF(B19=0,0,B20/B19)")
    c.font = LABEL_FONT
    c.number_format = PCT_FMT
    c.border = THIN_BORDER
    ws.cell(row=21, column=3,
            value='=IF(B21<0.3,"WEAK (<30% - high special-assessment risk)",'
                  'IF(B21<0.7,"FAIR (30-70%)","STRONG (70%+)"))').font = NORMAL_FONT
    ws.conditional_formatting.add("B21", CellIsRule(operator="lessThan", formula=["0.3"], fill=WARN_FILL, font=WARN_FONT))
    ws.conditional_formatting.add("B21", CellIsRule(operator="between", formula=["0.3", "0.6999"], fill=NEUTRAL_FILL))
    ws.conditional_formatting.add("B21", CellIsRule(operator="greaterThanOrEqual", formula=["0.7"], fill=GOOD_FILL, font=GOOD_FONT))
    ws.cell(row=22, column=1, value="CAI bands: under 30% = weak; 30-70% = fair; 70%+ = strong.").font = SMALL_FONT

    # --- Reduction readiness gates ---
    ws.cell(row=24, column=1, value="REDUCTION READINESS — ALL THREE GATES MUST PASS").font = SECTION_FONT
    for c, h in enumerate(["Gate", "Target", "Actual", "Status"], 1):
        ws.cell(row=25, column=c, value=h)
    style_header_row(ws, 25, 4)

    def gate_row(r, label, target_formula, actual_formula, status_formula):
        ws.cell(row=r, column=1, value=label).font = NORMAL_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        c = ws.cell(row=r, column=2, value=target_formula)
        c.number_format = CURRENCY_FMT
        c.border = THIN_BORDER
        c = ws.cell(row=r, column=3, value=actual_formula)
        c.number_format = CURRENCY_FMT
        c.border = THIN_BORDER
        c = ws.cell(row=r, column=4, value=status_formula)
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal="center")

    gate_row(26, "1. Emergency Reserve funded to target", "=B12", "=D12", '=IF(D12>=B12,"PASS","FAIL")')
    gate_row(27, "2. Long-Term Maintenance never below $0", 0,
             "=MIN('Long-Term Maintenance'!F5:F29)", '=IF(C27>=B27,"PASS","FAIL")')
    gate_row(28, "3. Operating Reserve at/above its floor", "=B14",
             "=MIN('Operating Reserve'!F5:F29)", '=IF(C28>=B28,"PASS","FAIL")')

    ws.cell(row=29, column=1, value="OVERALL").font = LABEL_FONT
    ws.cell(row=29, column=1).border = THIN_BORDER
    c = ws.cell(row=29, column=4,
                value='=IF(AND(D26="PASS",D27="PASS",D28="PASS"),'
                      '"ELIGIBLE TO CONSIDER REDUCTION","NOT ELIGIBLE — HOLD DUES")')
    c.font = LABEL_FONT
    c.border = THIN_BORDER
    c.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=29, start_column=2, end_row=29, end_column=3)

    for rng in ("D26", "D27", "D28"):
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng}="FAIL"'], fill=WARN_FILL, font=WARN_FONT))
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{rng}="PASS"'], fill=GOOD_FILL, font=GOOD_FONT))
    ws.conditional_formatting.add("D29", FormulaRule(formula=['ISNUMBER(SEARCH("NOT",D29))'], fill=WARN_FILL, font=WARN_FONT))
    ws.conditional_formatting.add("D29", FormulaRule(formula=['ISNUMBER(SEARCH("CONSIDER",D29))'], fill=GOOD_FILL, font=GOOD_FONT))

    # --- Combined cost ---
    ws.cell(row=31, column=1, value="Combined Monthly Cost per Unit (Year 1, all three funds):").font = LABEL_FONT
    c = ws.cell(row=31, column=2, value="=(C12+C13+C14)/B6/12")
    c.font = LABEL_FONT
    c.number_format = '$#,##0'
    c.border = THIN_BORDER
    ws.cell(row=32, column=1,
            value="Grows at the Growth Rate each year; this is Year-1. Add the HOA per-unit amount for a "
                  "unit owner's all-in monthly.").font = SMALL_FONT
    ws.cell(row=34, column=1,
            value="To model a dues reduction: lower a fund's Year-1 Contribution above and watch the gates. "
                  "Any gate turning FAIL means the reduction is not supported.").font = SMALL_FONT

    # Operating Reserve basis — drives the Operating Reserve floor (B14 = B38 x B37)
    ws.cell(row=36, column=1, value="OPERATING RESERVE BASIS").font = SECTION_FONT
    ws.cell(row=37, column=1, value="Monthly Operating Expense:").font = LABEL_FONT
    c = ws.cell(row=37, column=2, value=op_monthly)
    c.font = LABEL_FONT
    c.number_format = CURRENCY_FMT
    c.fill = INPUT_FILL
    c.border = THIN_BORDER
    ws.cell(row=37, column=3,
            value="The association's total monthly operating spend (excludes reserve contributions).").font = SMALL_FONT
    ws.cell(row=38, column=1, value="Operating Reserve Multiplier (months, 1-3):").font = LABEL_FONT
    c = ws.cell(row=38, column=2, value=op_mult)
    c.font = LABEL_FONT
    c.number_format = "0"
    c.fill = INPUT_FILL
    c.border = THIN_BORDER
    ws.cell(row=38, column=3,
            value="Operating Reserve floor (row 14) = this multiplier x monthly operating expense.").font = SMALL_FONT

    # --- Monthly funding per unit (Year 1) ---
    ws.cell(row=40, column=1, value="MONTHLY FUNDING PER UNIT (Year 1)").font = SECTION_FONT
    for c, h in enumerate(["Fund", "$ / unit / month"], 1):
        ws.cell(row=41, column=c, value=h)
    style_header_row(ws, 41, 2)
    for i, (nm, formula) in enumerate([
        ("Emergency Reserve", "=C12/$B$6/12"),
        ("Long-Term Maintenance", "=C13/$B$6/12"),
        ("Operating Reserve", "=C14/$B$6/12"),
    ]):
        r = 42 + i
        ws.cell(row=r, column=1, value=nm).font = NORMAL_FONT
        ws.cell(row=r, column=1).border = THIN_BORDER
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = '$#,##0.00'
        c.border = THIN_BORDER
    ws.cell(row=45, column=1, value="TOTAL reserve funding / unit / month").font = LABEL_FONT
    ws.cell(row=45, column=1).border = THIN_BORDER
    c = ws.cell(row=45, column=2, value="=(C12+C13+C14)/$B$6/12")
    c.font = LABEL_FONT
    c.number_format = '$#,##0.00'
    c.border = THIN_BORDER
    ws.cell(row=46, column=1,
            value="Equal per-unit split (fund contribution / units / 12); grows ~3%/yr. Excludes the separate "
                  "$365/unit HOA dues. If common charges are set by percentage interest, larger units pay "
                  "proportionally more than this average.").font = SMALL_FONT

    for col, w in [("A", 46), ("B", 18), ("C", 20), ("D", 18), ("E", 34)]:
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def generate_workbook(key, entity):
    wb = Workbook()

    # Sheet 1: Instructions
    ws_instr = wb.active
    ws_instr.title = "Instructions"
    build_instructions_sheet(ws_instr, entity)

    # Sheet 2: Fund Summary & Reduction Readiness
    build_summary_sheet(wb.create_sheet("Fund Summary"), entity)

    # Shared cross-sheet input references (all live on Fund Summary)
    growth_ref = "'Fund Summary'!$B$7"
    interest_ref = "'Fund Summary'!$B$8"
    units_ref = "'Fund Summary'!$B$6"
    common = {"growth_ref": growth_ref, "interest_ref": interest_ref, "units_ref": units_ref}

    # Sheet 3: Emergency Reserve
    build_fund_schedule(wb.create_sheet("Emergency Reserve"), entity, {
        "title": f"{entity['name']} — Emergency Reserve", "tab": "C00000",
        "start_ref": "'Fund Summary'!$D$12", "contrib_ref": "'Fund Summary'!$C$12",
        "floor_ref": "'Fund Summary'!$B$12", "floor_kind": "target", "disb": {}, **common})

    # Sheet 4: Long-Term Maintenance (Becht capital schedule)
    build_fund_schedule(wb.create_sheet("Long-Term Maintenance"), entity, {
        "title": f"{entity['name']} — Long-Term Maintenance", "tab": "548235",
        "start_ref": "'Fund Summary'!$D$13", "contrib_ref": "'Fund Summary'!$C$13",
        "floor_ref": "'Fund Summary'!$B$13", "floor_kind": "zero",
        "disb": entity["disbursements"], **common})

    # Sheet 5: Operating Reserve
    build_fund_schedule(wb.create_sheet("Operating Reserve"), entity, {
        "title": f"{entity['name']} — Operating Reserve", "tab": "7030A0",
        "start_ref": "'Fund Summary'!$D$14", "contrib_ref": "'Fund Summary'!$C$14",
        "floor_ref": "'Fund Summary'!$B$14", "floor_kind": "flat", "disb": {}, **common})

    # Sheet 6: Components
    build_components_sheet(wb.create_sheet("Components"), entity)

    # Sheet 7: Expenditure Chart (Long-Term Maintenance disbursements)
    build_chart_sheet(wb.create_sheet("Expenditure Chart"), entity)

    filename = entity["filename"]
    wb.save(filename)
    print(f"  Created: {filename}")


def main():
    print("Generating entity reserve funding workbooks...")
    for key, entity in ENTITIES.items():
        generate_workbook(key, entity)
    print(f"\nDone. {len(ENTITIES)} workbooks created.")


if __name__ == "__main__":
    main()
