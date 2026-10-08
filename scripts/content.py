"""Prose for the Ironbridge workshop pack: the four saved emails, the Teams
channel export, the risk policy memo and the pricing schedule text. Imported
by generate.py. Everything here is fiction. Numbers that must tie out are
passed in from the model so the prose and the data come from one place."""

MAYA = "Maya Okafor <maya.okafor@ironbridgepay.example>"
TOM = "Tom Hallstrom <tom.hallstrom@ironbridgepay.example>"
JENNA = "Jenna Okoye <jenna.okoye@ironbridgepay.example>"
PRIYA = "Priya Raman <priya.raman@ironbridgepay.example>"
GREG = "Greg Tanaka <greg.tanaka@ironbridgepay.example>"
BEN = "Ben Castellano <ben.castellano@ironbridgepay.example>"

# ---------------------------------------------------------------- pricing schedule (PDF)
PRICING_ROWS = [
    ("A", "$0.12", "8 bps", "Over $25,000,000"),
    ("B", "$0.15", "11 bps", "$5,000,000 to $25,000,000"),
    ("C", "$0.18", "15 bps", "Under $5,000,000"),
]
PRICING_OTHER = [
    "ACH credits (payroll, disbursements, refunds): $0.10 per item, all tiers. No basis-point charge on credit volume.",
    "Card acceptance (credit and debit cards, all brands): 22 basis points on gross card volume, all tiers, plus interchange and assessments passed through at cost.",
    "Remote check capture: $0.22 per item, billed separately and not included in the ACH tiers above.",
    "Returns: no per-return fee. Return handling is included in the per-item price.",
]
PRICING_FORMULA = (
    "A merchant's monthly processing fee is the sum of four lines: debit items multiplied by the per-item price for its tier, "
    "debit volume multiplied by the basis points for its tier, credit items multiplied by $0.10, and card volume multiplied by 22 basis points. "
    "One basis point is one hundredth of one percent, so 11 bps on $1,000,000 is $1,100."
)
PRICING_TIERING = (
    "Each merchant is assigned one tier in the merchant master and billed at that tier until the master is changed. "
    "Merchant Contracts reviews tier assignments each quarter against the trailing three-month average of monthly debit volume. "
    "A move between tiers takes effect only through a signed pricing amendment; the effective date in the amendment governs billing. "
    "Merchant Services is responsible for recording the new tier in the merchant master before the first billing cycle it applies to."
)

# ---------------------------------------------------------------- risk policy memo (docx)
MEMO_HEADER = [
    ("To:", "Merchant Services, Merchant Risk, Merchant Contracts"),
    ("From:", "Helen Varga, Chief Risk Officer"),
    ("Date:", "January 12, 2026"),
    ("Re:", "Merchant risk policy for ACH origination, 2026 revision"),
]
MEMO_SECTIONS = [
    {
        "title": "1. Purpose",
        "paras": [
            "Ironbridge originates ACH debits on behalf of its merchants and is accountable to the network and to our ODFI for the quality of those entries. "
            "This memo sets the return-rate thresholds we monitor, the internal warning lines we act on before a network threshold is reached, what happens at each line, and how often the portfolio is reviewed. "
            "It replaces the June 2024 memo. Nothing in it changes a merchant's contract; it governs how we manage risk inside those contracts.",
        ],
    },
    {
        "title": "2. Network thresholds",
        "paras": [
            "The Nacha Operating Rules set three return-rate levels that an ODFI must monitor for each originator. Each rate is the number of returned debit entries in the month divided by the number of debit entries originated in the same month.",
        ],
        "table": (
            ["Measure", "Return reason codes", "Network threshold"],
            [
                ["Unauthorized return rate", "R05, R07, R10, R11, R29", "0.5%"],
                ["Administrative return rate", "R02, R03, R04", "3.0%"],
                ["Overall return rate", "All return reason codes", "15.0%"],
            ],
        ),
    },
    {
        "title": "3. Ironbridge internal warning lines",
        "paras": [
            "We act before a merchant reaches a network threshold. The internal lines below apply to each merchant individually, measured monthly on the same basis as the network rates.",
        ],
        "table": (
            ["Measure", "Internal warning line", "Network threshold"],
            [
                ["Unauthorized return rate", "0.40%", "0.5%"],
                ["Administrative return rate", "2.5%", "3.0%"],
                ["Overall return rate", "12%", "15.0%"],
            ],
        ),
        "after": [
            "The portfolio-wide rate for each measure is reported to the board quarterly against the network thresholds. A portfolio rate that moves toward a line over several months is noted in the board report as a trend even when no single merchant has crossed it.",
        ],
    },
    {
        "title": "4. What happens at each line",
        "table": (
            ["Event", "Action", "Owner", "Timing"],
            [
                ["Merchant crosses an internal warning line in one month", "Warning letter to the merchant's principal stating the rate, the entries behind it and the network threshold. Account manager opens a remediation file.", "Merchant Services", "Within 5 business days of month-end close"],
                ["Merchant crosses an internal warning line in two consecutive months", "Written remediation plan required from the merchant: root cause, authorization and timing changes, expected date back under the line.", "Merchant Services with Merchant Risk", "Plan due 10 business days after the letter"],
                ["Merchant crosses a network threshold", "Formal notice. Rolling reserve of 5% of monthly debit volume is established and held for 90 days. Volume caps may be set.", "Merchant Risk", "Immediately on close of the month"],
                ["Merchant stays over a network threshold for two consecutive months", "Termination review by the Chief Risk Officer. Default outcome is termination with 30 days notice unless the remediation plan shows the rate falling.", "Chief Risk Officer", "Decision within 10 business days of the second month's close"],
            ],
        ),
        "after": [
            "Unauthorized returns are treated more seriously than administrative or NSF returns because they indicate a consumer did not agree to the debit. A merchant over the unauthorized line is also required to produce authorization records for a sample of the returned entries.",
        ],
    },
    {
        "title": "5. Monthly review",
        "paras": [
            "Merchant Services calculates return rates for every merchant from the monthly ACH activity export within five business days of month end and circulates a portfolio review to Merchant Risk. "
            "The review lists every merchant over an internal line, the dollars at stake, and the action taken. Rates are calculated from the export, not from merchant-reported figures.",
        ],
    },
    {
        "title": "6. Concentration",
        "paras": [
            "Any single merchant that accounts for more than 25% of trailing-twelve-month debit volume is reported to the board with the account's contract renewal date, its return rates and the steps being taken to retain it or to reduce reliance on it. "
            "Concentration above 25% is not a breach of policy. It is a risk the board has asked to see.",
        ],
    },
    {
        "title": "7. Definitions",
        "paras": [
            "Unauthorized returns: R05 (unauthorized debit to consumer account using corporate SEC code), R07 (authorization revoked by customer), R10 (customer advises not authorized), R11 (customer advises entry not in accordance with the terms of the authorization), R29 (corporate customer advises not authorized).",
            "Administrative returns: R02 (account closed), R03 (no account or unable to locate account), R04 (invalid account number).",
            "NSF returns: R01 (insufficient funds), R09 (uncollected funds). NSF returns count toward the overall rate only.",
        ],
    },
]
MEMO_SIGNOFF = "Questions about this policy go to Helen Varga. Questions about a specific merchant go to that merchant's account manager first."

# ---------------------------------------------------------------- emails (plain text, Outlook save-as-text)


def emails(model):
    rows = model["rows"]
    summit_aug = model["summit"][10]
    ng_aug_m = model["ng_aug_vol"] / 1e6
    return [
        {
            "filename": "RE_ return rate notice.txt",
            "from": "Rachel Lindqvist <rlindqvist@summitridgelending.example>",
            "to": MAYA,
            "cc": TOM,
            "date": "Thursday, September 24, 2026 4:12 PM",
            "subject": "RE: Return rate notice - Summit Ridge Lending - August 2026",
            "body": f"""Maya,

I have your letter from Monday in front of me and I want to push back before this goes any further.

Every one of our borrowers authorizes every debit inside the app. They see the amount, they see the date, they tap to confirm, and we keep the record. Nobody at Summit Ridge is pulling money from an account without consent, and a letter that puts the word "unauthorized" next to our name is not something I can take to my board without the detail behind it.

So I need the underlying entries. For every return you are counting as unauthorized in August, send me the trace number, the return reason code, the settlement date and the amount. Not a total, the entries. Our servicing team will match each one to its authorization record and we will tell you what we find.

One thing that may be relevant: in June we changed our debit timing. We used to pull on each borrower's pay date. Starting in June we moved everyone to a fixed cycle on the 1st and the 15th so that our collections team could work exceptions in two batches instead of all month. The borrowers were notified in the app and by email. If some of them are disputing debits that landed on a date they did not expect, that is a timing issue, not an authorization issue, and I would like to understand it before anyone calls it a pattern.

I am available Monday or Tuesday for a call. Please include Tom.

Rachel Lindqvist
Chief Operating Officer
Summit Ridge Lending

-----Original Message-----
From: {MAYA}
Sent: Monday, September 21, 2026 9:40 AM
To: Rachel Lindqvist <rlindqvist@summitridgelending.example>
Cc: {TOM}
Subject: Return rate notice - Summit Ridge Lending - August 2026

Rachel,

This is the formal notice I mentioned on the phone. Summit Ridge's unauthorized return rate for August 2026 was {summit_aug:.2f}% of debit entries originated, which is above the Nacha network threshold of 0.5% and above Ironbridge's internal warning line of 0.40% for the second month in a row. Under our merchant risk policy that means we need a written remediation plan from you within ten business days, and I have to tell you plainly that a second month over the network threshold starts a termination review on our side whether or not either of us wants it.

I know this is not the direction either of us expected the account to go. Tom will send the August return detail by code tomorrow, and I would rather spend the call on what changed than on the letter.

Maya Okafor
VP, Merchant Services
Ironbridge Payments
""",
        },
        {
            "filename": "FW_ pricing amendment - Harrow Finch.txt",
            "from": GREG,
            "to": MAYA,
            "cc": BEN + "; " + JENNA,
            "date": "Friday, March 27, 2026 2:51 PM",
            "subject": "FW: Pricing amendment - Harrow & Finch Receivables - countersigned",
            "body": f"""Maya,

Countersigned copy below. Harrow & Finch moves from Tier B to Tier C effective April 1. Their trailing three months came in under $5 million a month and Karen asked for the review rather than wait for us.

Ben, please update the pricing tier in the merchant master before the April close so billing picks it up. Jenna, FYI since it is your account.

Greg Tanaka
Director, Merchant Contracts
Ironbridge Payments

-----Original Message-----
From: Karen Oyelaran <k.oyelaran@harrowfinch.example>
Sent: Friday, March 27, 2026 1:18 PM
To: {GREG}
Subject: RE: Pricing amendment - Harrow & Finch Receivables

Greg,

Signed by both of us. Text of the amendment is pasted below for the file since our scanner is still out. The original is in the mail.

Thank you for turning this around quickly.

Karen Oyelaran
Chief Financial Officer
Harrow & Finch Receivables, LLC


AMENDMENT NO. 2 TO MERCHANT PROCESSING AGREEMENT

This Amendment No. 2 (the "Amendment") is made as of March 26, 2026 between Ironbridge Payments, Inc. ("Ironbridge") and Harrow & Finch Receivables, LLC ("Merchant") and amends the Merchant Processing Agreement dated November 2, 2020 (the "Agreement").

1. Pricing tier. Effective April 1, 2026, Merchant's pricing tier under Schedule A of the Agreement is changed from Tier B to Tier C. From that date Merchant's ACH debit pricing is $0.18 per debit item and 15 basis points on debit volume, as set out in the Ironbridge Merchant Pricing Schedule 2026.

2. Basis. The parties acknowledge that Merchant's average monthly ACH debit volume for the three months ended February 28, 2026 was below $5,000,000, and that under the Pricing Schedule Tier C applies to merchants under $5,000,000 in monthly debit volume.

3. Other pricing. Credit items, card acceptance and remote check capture pricing are unchanged.

4. Review. Ironbridge will review Merchant's tier at the next quarterly review and the parties will execute a further amendment if Merchant's volume returns to the Tier B range for three consecutive months.

5. Effect. Except as changed by this Amendment the Agreement remains in full force.

Ironbridge Payments, Inc.                        Harrow & Finch Receivables, LLC
By: Greg Tanaka                                  By: Karen Oyelaran
Title: Director, Merchant Contracts              Title: Chief Financial Officer
Date: March 26, 2026                             Date: March 27, 2026
""",
        },
        {
            "filename": "Northgate - processor review.txt",
            "from": "Daniel Achebe <dachebe@northgatebnpl.example>",
            "to": MAYA,
            "cc": PRIYA,
            "date": "Friday, September 18, 2026 11:06 AM",
            "subject": "Northgate - processor review",
            "body": f"""Maya,

I wanted you to hear this from me rather than from a procurement email.

Our board has asked us to run a processor review before the contract renews at year end. We are talking to two other providers alongside Ironbridge. The decision will be made by November 15 so that whoever we choose has time to be in place for the January cycle.

To be direct about where we stand: the service has been good, Priya is responsive, and nobody here is looking to leave for the sake of it. But we sent you a little over ${int(ng_aug_m)} million in debits in August alone and the volume keeps growing, and at that size the board wants to know the pricing reflects it. The other two have both come in under our current per-item rate.

Two things I need from Ironbridge by October 16:

1. Your best and final on pricing. Per item and basis points, for the volume we are running now, not the volume we were running when the agreement was signed.

2. A return-handling SLA in writing. Specifically: return files delivered by 7:00 AM Eastern the banking day after settlement, R10 and R11 returns flagged to us within 24 hours with the trace number and consumer name, and a named contact who can answer a return question the same day. Today this works because Priya makes it work. We need it in the agreement.

If you want to talk before you put anything in writing, I am free most of next week.

Daniel Achebe
VP Finance
Northgate BNPL
""",
        },
    ]


LUNCH_ORDER = """Thursday lunch - Riverside Deli - someone call it in by 11

Priya - turkey club on wheat, no tomato, chips
Tom - italian sub, extra peppers, no onion
Jenna - greek salad with chicken, dressing on the side
Luis - roast beef on rye, horseradish, pickle
Dev - veggie wrap, add avocado
Ben - meatball sub (if they have it, otherwise chicken parm)
Maya - cup of tomato soup and half a tuna sandwich
Greg - whatever the special is, he said he doesn't care

Two bags of the kettle chips for the table. Receipt to Maya for the card.
"""

# ---------------------------------------------------------------- Teams channel export


def teams_chat(model):
    rows = model["rows"]
    summit_jul = model["summit"][9]
    summit_aug = model["summit"][10]
    cedar_aug = rows[(10, "M-1005")]["vol"] / 1e6
    lines = [
        ("2026-09-01 08:14", "Jenna Okoye", "Morning. August export is up in the shared folder. Ben is out as of today so I pulled it myself, let me know if anything looks off."),
        ("2026-09-01 08:22", "Tom Hallstrom", "thanks. did the May file ever get cleaned up? there are still two of them in there"),
        ("2026-09-01 08:31", "Jenna Okoye", "Ben pulled May twice when the report timed out, one of them is just a copy, ignore it. He said he'd delete it and then he went on leave."),
        ("2026-09-01 08:33", "Tom Hallstrom", "ok. leaving it alone then, not my folder"),
        ("2026-09-02 10:05", "Dev Mehta", "August return rates are out. Summit Ridge unauthorized is climbing again, " + ("%.2f" % summit_jul) + " in July and " + ("%.2f" % summit_aug) + " in August. That is two months over our line and now over the Nacha line too."),
        ("2026-09-02 10:07", "Tom Hallstrom", "I know. Rachel is going to lose it when the letter goes out. they are adamant everything is authorized in the app"),
        ("2026-09-02 10:12", "Dev Mehta", "Reminder for everyone since we have a few newer merchants: Ironbridge's internal warning line is 0.40% unauthorized, tighter than Nacha's 0.50%. We act at 0.40, the network acts at 0.50. Admin is 2.5 internal, 3.0 network. Overall 12 and 15."),
        ("2026-09-02 10:15", "Priya Raman", "is the Nacha letter risk real for Summit or is that a Dev worst case"),
        ("2026-09-02 10:19", "Dev Mehta", "It is real. Two consecutive months over 0.5 and the ODFI can require a reduction plan and Nacha can open an inquiry. Maya's letter has to go out this month under the policy. I am not making this up to scare Tom."),
        ("2026-09-02 10:20", "Tom Hallstrom", "noted. I'll draft it for Maya today"),
        ("2026-09-03 14:40", "Luis Ferreira", "Pinecrest renewal is Oct 31. anyone know if Greg has the paper started or do I need to chase"),
        ("2026-09-03 14:52", "Jenna Okoye", "chase. he is buried in the Tri-County card addendum"),
        ("2026-09-04 09:02", "Priya Raman", "Redfern Apparel is live as of Aug 11, first partial month in the August file. expect them to look tiny until October"),
        ("2026-09-04 09:30", "Jenna Okoye", "Has anyone heard from Cedar Falls? They went quiet since July. Volume is off a cliff and nobody there answers me. Their renewal is November and I have a bad feeling they are not renewing."),
        ("2026-09-04 09:33", "Luis Ferreira", "I saw their August number, " + ("%.1f" % cedar_aug) + " million? they were doing three and change in the spring"),
        ("2026-09-04 09:35", "Jenna Okoye", "yes. and the status in the master says Notice given, which I did not put there, so somebody heard something. Ben maybe. I'll ask Greg."),
        ("2026-09-08 11:18", "Tom Hallstrom", "Summit letter went to Maya for signature. Rachel's going to ask for every entry, I already know it"),
        ("2026-09-08 11:25", "Dev Mehta", "Give her the entries. That is the right ask. Have her look at what changed in June, the jump starts there."),
        ("2026-09-09 15:47", "Jenna Okoye", "Did Harrow & Finch's new pricing ever make it into the master? Karen signed it back in March and I honestly do not remember seeing the tier change."),
        ("2026-09-09 16:30", "Luis Ferreira", "Westbrook asked for a Q4 volume forecast template. do we have one or do I make one"),
        ("2026-09-09 16:41", "Priya Raman", "make one. there is nothing"),
        ("2026-09-11 08:50", "Priya Raman", "heads up, Daniel at Northgate wants a call with Maya next week. he would not say what about. not a return thing, their rates are fine"),
        ("2026-09-11 08:55", "Tom Hallstrom", "that is never good when they will not say"),
        ("2026-09-14 13:02", "Dev Mehta", "Admin returns across the book keep creeping up. Nothing over a line yet but it has gone up a little every quarter for a year. Mostly R03 and R04. I think a few of the ARM shops are loading old account numbers. Will put it in the quarterly."),
        ("2026-09-14 13:10", "Jenna Okoye", "Rivermark admitted they bought a portfolio in the spring with bad account data. that is probably some of it"),
        ("2026-09-16 10:22", "Luis Ferreira", "Stonebridge renewal also Oct 31. two for Greg. adding to his list"),
        ("2026-09-18 11:40", "Priya Raman", "Northgate call happened. It is a processor review. Two other providers. Decision by Nov 15. Maya has the email. I would like everyone to not panic in this channel please."),
        ("2026-09-18 11:42", "Tom Hallstrom", "that is a quarter of the book"),
        ("2026-09-18 11:43", "Priya Raman", "I am aware"),
        ("2026-09-22 09:15", "Jenna Okoye", "Maya's Summit letter went out yesterday. Tom you are cc'd."),
        ("2026-09-24 16:30", "Tom Hallstrom", "Rachel replied. she wants every entry and she says they changed debit timing in June. Dev was right. forwarding to Maya"),
        ("2026-09-25 08:05", "Dev Mehta", "September close is Oct 5. I will have rates out Oct 7. If Summit is over 0.5 again it goes to Helen for termination review, that is not my call to soften."),
        ("2026-09-29 12:12", "Luis Ferreira", "lunch order for thursday is in the shared folder, add yourself"),
        ("2026-09-30 17:01", "Jenna Okoye", "September export will be in the folder once the batch finishes tonight. Same file name mess as always, sorry, Ben's naming is Ben's naming."),
    ]
    out = ["Merchant Ops - Microsoft Teams channel export", "Exported by Maya Okafor on 2026-10-06, messages 2026-09-01 to 2026-09-30", ""]
    for when, who, text in lines:
        out.append("[%s] %s: %s" % (when, who, text))
    return "\n".join(out) + "\n"
