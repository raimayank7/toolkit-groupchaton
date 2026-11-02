"""Contracts & Agreements toolkit: Word files (checklist before signing, clause library, mutual NDA, MOU)."""
from make_files import *

GREY = '64748B'

def note(d, text):
    """grey guidance line the user deletes before signing"""
    return para(d, 'Guidance: ' + text, size=9, italic=True, color=GREY, space=6)

def heading(d, text, size=13):
    return para(d, text, bold=True, size=size, color=BLUE, space=4)

def clause(d, num, title_, items):
    p = para(d, '%s. %s' % (num, title_), bold=True, size=11.5, space=3)
    p.paragraph_format.keep_with_next = True
    for i, t in enumerate(items, 1):
        q = para(d, '%s.%d  %s' % (num, i, t), space=4)
        q.paragraph_format.left_indent = Cm(0.9); q.paragraph_format.first_line_indent = Cm(-0.9)

def sign_block(d, left, right, witnesses=2):
    t = d.add_table(rows=1, cols=2)
    for c, who in zip(t.rows[0].cells, (left, right)):
        c.text = ''
        r = c.paragraphs[0].add_run(who); r.bold = True
        for line in ('Name: [ ]', 'Title: [ ]', '', 'Signature: ______________________', 'Date: [ ]'):
            c.add_paragraph(line)
    para(d, space=6)
    if witnesses:
        t = d.add_table(rows=1, cols=witnesses)
        for i, c in enumerate(t.rows[0].cells, 1):
            c.text = ''; c.paragraphs[0].add_run('Witness %d' % i).bold = True
            for line in ('Signature: ______________________', 'Name: [ ]', 'Address: [ ]'):
                c.add_paragraph(line)

# ------------------------------------------------------------------ 1. checklist before signing
CHECKS = [
    ('Parties and authority', [
        'Full legal names match ID documents, company records or the trade licence (no nicknames or brand names alone).',
        'Addresses are complete, with PIN or postcode, and an email for notices is given for each party.',
        'For a company, LLP or firm: the person signing has authority (board resolution, partner consent or power of attorney seen).',
        'Each party is legally able to contract (adult, of sound mind, not barred by law).']),
    ('What is being exchanged', [
        'The goods, services, property or rights are described so a stranger could tell exactly what is included and what is not.',
        'Quantities, specifications, standards and service levels are measurable (numbers, dates, named standards).',
        'Schedules, drawings, price lists or annexures referred to are actually attached and signed.']),
    ('Money and payment', [
        'Price or fee is stated in figures and words, and both match.',
        'It says whether tax (GST, VAT, sales tax) is included or extra.',
        'Payment dates, method and bank details are stated; advances and deposits say when and how they are refunded.',
        'Late payment interest, if any, is reasonable and stated as a rate, not a vague "penalty".',
        'Who pays expenses, stamp duty, registration and other costs is clear.']),
    ('Dates and term', [
        'Start date, end date or term, and any renewal (automatic or by agreement) are clear.',
        'Deadlines for delivery or completion are stated, with what happens if they are missed.',
        'Lock-in periods and notice periods work both ways, or the difference is intended.']),
    ('Performance and quality', [
        'How work is checked, accepted or rejected is described (inspection period, acceptance tests, revisions).',
        'Warranties or guarantees say what is covered, for how long and the remedy (repair, replace, refund).',
        'Changes to scope need a written change order with price and time agreed.']),
    ('Ending the agreement', [
        'Either side can end for serious breach after a reasonable notice to fix it.',
        'Ending for convenience (without a reason) is allowed or excluded on purpose, with the notice period.',
        'What happens at the end is clear: final payment, handover, return of property and data, refund of deposits.']),
    ('Risk clauses', [
        'Indemnities: who covers which losses, and whether they are capped. One-sided indemnities are understood.',
        'Limitation of liability: the cap amount is acceptable and excludes only what the law allows.',
        'Insurance required of either party is affordable and actually in place.',
        'Personal guarantees, security or post-dated cheques are understood by the person giving them.']),
    ('Confidentiality, IP and data', [
        'Confidential information is defined and the duty lasts a sensible time.',
        'Who owns work, designs, code or content created under the agreement is stated, and when ownership passes.',
        'Use of names, logos and trade marks is limited to what is agreed.',
        'Personal data is handled under the data protection law that applies.']),
    ('Disputes and law', [
        'Governing law and the courts or arbitration seat are stated and practical for you to reach.',
        'There is a step to talk or mediate before going to court or arbitration.',
        'Non-compete and non-solicitation clauses are no wider than the law allows where you are.']),
    ('Signing and records', [
        'Stamp duty, registration, notarisation and witnesses checked for the country and state (see the table below).',
        'Every page is initialled, blanks are filled or struck out, and changes are initialled by both sides.',
        'Each party keeps a signed original (or a full scanned copy), with all schedules.',
        'Key dates (renewal, notice deadlines, payment dates) are put in a calendar with reminders.']),
]
RED = [
    ('Blank spaces or "to be decided later" for price, dates or scope', 'Fill every blank before signing, or strike it out and initial.'),
    ('"Penalty" of a large fixed amount for any breach', 'A reasonable pre-estimate of loss, or actual losses proved.'),
    ('Only one side can end the agreement, or notice periods differ greatly', 'Equal rights to end, or a clear business reason for the difference.'),
    ('Automatic renewal with a short window to cancel', 'Renewal only by written agreement, or a reminder before renewal.'),
    ('Unlimited indemnity from you, none from the other side', 'Mutual indemnities limited to each party\'s own fault, with a cap.'),
    ('Ownership of your existing tools or know-how passes to the client', 'Only work made specially for the client passes; you keep background material.'),
    ('Deposit or advance "non-refundable in all cases"', 'Refundable if the other side fails, and partly refundable on early cancellation.'),
    ('Non-compete that lasts years after the end, everywhere', 'A narrow restriction during the term, or as the law allows.'),
    ('Disputes only in a far-away court or foreign arbitration', 'A seat or courts near both parties, or online hearings.'),
    ('Amounts in figures and words do not match', 'Correct the mismatch before signing; courts often prefer the words.'),
    ('Schedules or annexures "as agreed" but not attached', 'Attach and sign every schedule.'),
    ('Signatory is not the owner, director or an authorised person', 'Ask for the authority document (resolution, POA, partner consent).'),
    ('Cash payments of large amounts', 'Bank transfer or cheque, with receipts; some countries ban large cash deals.'),
    ('Unclear who pays stamp duty, tax or registration', 'Name the party that pays each cost.'),
    ('Copied text naming a different country\'s law or another business', 'Remove every leftover from the source document.'),
]
FORMAL = [
    ('India', 'Most agreements are stamped under the Indian or state Stamp Act (stamp paper, e-stamp or franking); amount depends on the state and document.',
     'Leases of 12 months or more, sale of immovable property, and some powers of attorney must be registered (Registration Act, 1908). Some states register all rent agreements.',
     'Two witnesses is common practice; notary for affidavits and powers of attorney.', 'Any language; English widely used.'),
    ('United Kingdom', 'No stamp duty on most contracts; SDLT, LBTT or LTT on land deals; stamp duty on some share transfers.',
     'Leases over 7 years and property transfers at HM Land Registry.', 'Deeds need one witness; most contracts need none.', 'English (Welsh also valid in Wales).'),
    ('United States', 'No general stamp duty; some states and counties charge transfer or recording taxes on real estate.',
     'Real-estate deeds and some leases recorded with the county.', 'Notary for powers of attorney and recorded documents; witnesses for some state forms.', 'English; some states require a translation for consumers.'),
    ('UAE', 'No general stamp duty; registration fees for tenancy and property.',
     'Tenancies on Ejari (Dubai) or Tawtheeq (Abu Dhabi); property sales at the land department; employment on the official MOHRE contract.',
     'Notary public for powers of attorney; legalisation for documents signed abroad.', 'Courts work in Arabic: a certified Arabic translation may be needed.'),
]

def checklist():
    d = doc_new('ct-contract-checklist-word')
    para(d, 'CONTRACT CHECKLIST BEFORE YOU SIGN', bold=True, size=17, color=BLUE, space=2)
    para(d, 'Read the agreement once from start to end, then go through these checks with the agreement open beside you. Tick each box, and write the clause number and your concern in Notes.', size=10, space=8)
    note(d, 'This checklist is a reminder of what to read carefully. It is not legal advice; for property, loans, large amounts or long terms, take your notes to a lawyer.')
    kv(d, [('Agreement', '[name / type]'), ('Other party', '[ ]'), ('Value', '[amount and currency]'), ('Start and end', '[ ]'),
           ('Received on', '[ ]'), ('Sign by (deadline)', '[ ]'), ('Reviewed by', '[ ]'), ('Lawyer consulted', '[Yes / No]')])
    for n, (sect, items) in enumerate(CHECKS, 1):
        heading(d, '%d. %s' % (n, sect), 12)
        table(d, ['✓', 'Check', 'Notes (clause number, concern)'], [['☐', s, ''] for s in items], [1, 10, 6])
    heading(d, 'Red flags: what to ask for instead', 12)
    table(d, ['☐', 'Red flag', 'Ask for instead'], [['☐', a, b] for a, b in RED], [1, 8, 8])
    heading(d, 'Signing formalities by country (check the current rule)', 12)
    table(d, ['Country', 'Stamp duty', 'Registration', 'Witnesses and notary', 'Language'], [list(r) for r in FORMAL], [2.2, 4, 4.5, 3.6, 2.7])
    note(d, 'Rules change and differ by state, emirate and city. Ask a local lawyer, notary or stamp vendor for the current amount and procedure.')
    heading(d, 'Sign-off', 12)
    kv(d, [('Changes asked for', '[ ]'), ('Changes agreed', '[ ]'), ('Open points', '[ ]'), ('Decision', '[Sign / Do not sign / Wait]'),
           ('Checked by', '[name, date]'), ('Signed copy kept at', '[file / folder]')], cols=1)
    dsave(d, 'Contract-Checklist-Before-Signing.docx')

# ------------------------------------------------------------------ 2. clause library
CL = [
    ('Definitions', '"[Term]" means [meaning]. Words in the singular include the plural, "including" means "including without limitation", and headings do not affect meaning.',
     'Fixes the meaning of words used again and again, so the rest of the contract can be short and consistent.', 'Every contract longer than a page.',
     'A defined word used with a different meaning elsewhere; definitions that contain obligations (put duties in the clauses, not definitions).', 'Same in every country.'),
    ('Term', 'This Agreement starts on [date] and continues for [period], unless ended earlier under its terms. It may be renewed only by written agreement.',
     'Says how long the contract lasts.', 'Every contract with ongoing duties.', 'Automatic renewal you might forget; a term that ends before the work does.',
     'India: leases of 12 months or more need registration. UK/US: auto-renewal in consumer contracts has notice rules.'),
    ('Payment', 'The [Client] will pay the fees in Schedule [ ] within [15] days of receiving a correct invoice, by bank transfer to the account named in it.',
     'When and how money is paid.', 'Any contract with money.', 'No due date; payment "on completion" when completion is not defined.',
     'India: TDS may apply; MSME suppliers must be paid within 45 days. UK: Late Payment Act gives statutory interest between businesses.'),
    ('Late payment', 'Overdue amounts carry simple interest at [1.5]% per month from the due date until paid, which the Parties agree is a reasonable estimate of loss.',
     'Compensates for late payment and encourages paying on time.', 'Contracts with credit periods.', 'A rate so high it looks like a penalty and may not be enforced.',
     'India: section 74 Contract Act allows only reasonable compensation. US: state usury limits.'),
    ('Taxes', 'All amounts are exclusive of [GST / VAT / sales tax], which will be added at the applicable rate and paid against a valid tax invoice.',
     'Says whether the price includes tax.', 'Any priced contract.', 'Silence: then the price may be read as tax-inclusive.', 'Name the tax: GST (India), VAT (UK, UAE), sales tax (US).'),
    ('Confidentiality', 'Each Party will keep the other\'s Confidential Information secret, use it only for this Agreement and disclose it only to people who need it and are bound to keep it secret, for [3] years after the end.',
     'Protects business secrets shared during the deal.', 'Most business contracts.', 'No exceptions for public information or legal requirements; no time limit.', 'Similar everywhere; trade secrets can be protected indefinitely.'),
    ('Data protection', 'Each Party will process personal data received under this Agreement only for its purpose and in line with [applicable data protection law], keep it secure and report breaches without undue delay.',
     'Meets privacy law when customer or staff data is shared.', 'Any contract involving personal data.', 'A processor with no instructions or deletion duty.',
     'India: DPDP Act, 2023. UK: UK GDPR and DPA 2018. UAE: Federal Decree-Law 45 of 2021. US: state privacy laws.'),
    ('Intellectual property (ownership)', 'On full payment, all rights in the Deliverables pass to the [Client]. The [Provider] keeps its pre-existing materials and grants the [Client] a permanent licence to use them in the Deliverables.',
     'Says who owns work created under the contract.', 'Design, writing, software, photography, consulting.', 'Assigning your own tools and templates by mistake.',
     'India: copyright assignment should state period and territory (section 19 Copyright Act) or defaults apply. US: "work made for hire" has a narrow meaning.'),
    ('Licence', 'The [Licensor] grants the [Licensee] a [non-exclusive] licence to use [the Work] in [media] in [territory] for [period]. All other rights are reserved.',
     'Permission to use something without owning it.', 'Content, software, brands, photos.', 'Vague scope ("any use"); missing territory or period.', 'Similar everywhere.'),
    ('Warranties', 'The [Seller] warrants that the [goods/services] will match their description, be of satisfactory quality and be free from defects for [12] months.',
     'Promises about quality, with a remedy if broken.', 'Sale and service contracts.', 'Exclusions that remove all protection; consumers often cannot lose statutory rights.',
     'UK: Consumer Rights Act 2015. India: Sale of Goods Act 1930 and Consumer Protection Act 2019. US: UCC implied warranties can be excluded only in a specific way.'),
    ('Indemnity', 'Each Party will indemnify the other against losses and third-party claims arising from its breach, negligence or breach of law, provided it is told promptly and can control the defence.',
     'A promise to cover certain losses, including claims by outsiders.', 'Services, supply, licensing, events.', 'One-sided or unlimited indemnities; indemnities for the other side\'s own fault.',
     'India: section 124 Contract Act. Insurers may need to approve indemnities you give.'),
    ('Limitation of liability', 'Neither Party is liable for indirect loss or loss of profit. Each Party\'s total liability is capped at [the fees paid in the last 12 months], except for fraud, death or injury by negligence and other liability that cannot be limited by law.',
     'Caps how much one side can be made to pay.', 'Most business contracts.', 'A cap so low it makes the contract meaningless; caps that also cover unpaid fees.',
     'UK: Unfair Contract Terms Act 1977 reasonableness test. India: courts generally respect negotiated caps between businesses.'),
    ('Insurance', 'The [Provider] will keep public liability insurance of at least [amount] and professional indemnity insurance of at least [amount] during this Agreement, and show certificates on request.',
     'Makes sure money exists to pay claims.', 'Work on site, professional advice, events.', 'Insurance demanded that is not available or affordable.', 'Similar everywhere; workers\' compensation rules vary.'),
    ('Force majeure', 'Neither Party is liable for delay caused by events beyond its reasonable control (such as flood, epidemic, war or government order), provided it gives prompt notice and limits the effect. Either may end the Agreement if the event lasts over [60] days.',
     'Excuses delay from events nobody could control.', 'Contracts with deadlines or long terms.', 'Lists that include ordinary business risks (price rises, staff shortage).',
     'India: section 56 Contract Act (frustration) applies without a clause too. UK/US: frustration is narrow, so the clause matters.'),
    ('Termination for breach', 'Either Party may end this Agreement by written notice if the other commits a serious breach that is not put right within [15] days of notice, or becomes insolvent.',
     'A way out when the other side fails.', 'Every ongoing contract.', 'No cure period; "any breach" however small.', 'Insolvency termination may be restricted in UK supply contracts (Corporate Insolvency and Governance Act 2020).'),
    ('Termination for convenience', 'Either Party may end this Agreement for any reason by giving [30] days\' written notice. The [Client] will pay for work done up to the end date.',
     'A way out without a reason.', 'Services and long arrangements.', 'Only one side has the right; no payment for work done.', 'Similar everywhere.'),
    ('Consequences of ending', 'On ending, each Party returns the other\'s property and Confidential Information, amounts due up to the end are paid, and clauses meant to survive (confidentiality, IP, liability, disputes) continue.',
     'Ties up loose ends.', 'Every ongoing contract.', 'Forgetting deposits, data and handover.', 'Similar everywhere.'),
    ('Non-solicitation', 'During this Agreement and for [12] months after, neither Party will persuade the other\'s employees involved in this Agreement to leave. General advertisements are allowed.',
     'Stops poaching of staff (and sometimes customers).', 'Agencies, outsourcing, consultants.', 'Very long periods; covering all staff of a big company.', 'India: more likely upheld than non-competes but still tested for reasonableness.'),
    ('Non-compete', 'During this Agreement [and for [6] months after], the [Party] will not carry on a competing business [in the territory], to the extent permitted by law.',
     'Stops competition.', 'Franchises, distribution, sale of a business, senior employees.', 'Clauses wider than the law allows are often void and can taint trust.',
     'India: post-term restraints are generally void (section 27 Contract Act) except on sale of goodwill. UK: must be reasonable. US: banned or limited in several states, e.g. California. UAE: allowed within limits.'),
    ('Assignment', 'Neither Party may transfer its rights or obligations under this Agreement without the other\'s written consent, which will not be unreasonably withheld.',
     'Controls who you end up dealing with.', 'Most contracts.', 'Allowing transfer to a competitor; blocking transfer in a sale of your business.', 'Similar everywhere.'),
    ('Notices', 'Notices must be in writing and delivered by hand, courier, registered post or email to the addresses above. Email notices are received when sent unless an error message is received.',
     'Makes sure formal notices count.', 'Every contract.', 'Out-of-date addresses; email not allowed for important notices.', 'Similar everywhere.'),
    ('Governing law and courts', 'This Agreement is governed by [the laws of India / the law of England and Wales / the laws of the State of [ ] / the laws of the UAE]. The courts at [city] have exclusive jurisdiction.',
     'Decides which law applies and where disputes are heard.', 'Every contract, especially cross-border.', 'A court far from you; law that does not match the place of performance.', 'Pick one law and one place; consumers may keep rights under their home law.'),
    ('Arbitration', 'Disputes not settled by negotiation within [30] days will be decided by a sole arbitrator under [the Arbitration and Conciliation Act, 1996 / the Arbitration Act 1996 / AAA Commercial Rules / DIAC Rules], seated at [city], in English.',
     'Private decision by an arbitrator instead of a court.', 'Business contracts, especially cross-border.', 'Cost for small disputes; seat far away; three arbitrators for small amounts.',
     'India: name the seat clearly. UAE: DIAC replaced DIFC-LCIA in 2021. US: Federal Arbitration Act; consumer arbitration has extra rules.'),
    ('Entire agreement and severability', 'This Agreement is the whole agreement on its subject and replaces earlier discussions. If any part is invalid, the rest stays in force.',
     'Stops arguments about earlier promises, and saves the contract if one clause fails.', 'Every contract.', 'Important promises made only by email or in a proposal: put them in the contract.', 'Similar everywhere; fraud is never excluded.'),
]

def clause_library():
    d = doc_new('ct-clause-library-word')
    para(d, 'CONTRACT CLAUSE LIBRARY', bold=True, size=17, color=BLUE, space=2)
    para(d, '24 standard clauses with sample wording and plain-English notes. Copy the wording into your own agreement, fill every [square bracket] and make the defined words match.', size=10, space=6)
    note(d, 'Sample wording is general and not legal advice. Laws differ by country and state; have important contracts checked by a lawyer.')
    table(d, ['#', 'Clause', 'Use it for'], [[str(i), c[0], c[3]] for i, c in enumerate(CL, 1)], [1, 6, 10])
    for i, (name, wording, means, when, watch, country) in enumerate(CL, 1):
        h = heading(d, '%d. %s' % (i, name), 12.5); h.paragraph_format.keep_with_next = True
        t = d.add_table(rows=1, cols=1); t.style = 'Table Grid'; c = t.rows[0].cells[0]; c.text = ''
        r = c.paragraphs[0].add_run(wording); r.font.size = Pt(10.5); shade(c, 'F1F5F9')
        para(d, space=2)
        for lab, txt in (('What it means: ', means), ('When to use it: ', when), ('Watch for: ', watch), ('Country notes: ', country)):
            p = para(d, space=3); a = p.add_run(lab); a.bold = True; a.font.size = Pt(10); b = p.add_run(txt); b.font.size = Pt(10)
    para(d, space=8)
    para(d, 'Want these clauses fitted into a complete agreement? Use the free online agreement and contract builder at toolkit.groupchaton.com/contracts/ct-contract-builder/.', size=9.5, italic=True, color=GREY)
    dsave(d, 'Contract-Clause-Library.docx')

# ------------------------------------------------------------------ 3. mutual NDA
def mutual_nda():
    d = doc_new('ct-mutual-nda-word')
    para(d, 'MUTUAL NON-DISCLOSURE AGREEMENT', bold=True, size=16, align='c', space=10)
    note(d, 'Starting draft, not legal advice. Fill every [square bracket], delete these grey guidance lines, and check stamp duty or other formalities where you are.')
    para(d, 'This Agreement is made on [date] at [city] between:', space=6)
    for n, who in ((1, 'First Party'), (2, 'Second Party')):
        p = para(d, '(%d)  [Full legal name], [a company incorporated under [law] / an individual], having its [registered office / address] at [full address], acting through [name, title] (the "%s"); %s' % (n, who, 'and' if n == 1 else ''), space=4)
        p.paragraph_format.left_indent = Cm(0.9); p.paragraph_format.first_line_indent = Cm(-0.9)
    para(d, 'together the "Parties" and each a "Party".', space=8)
    heading(d, 'Background', 11.5)
    para(d, 'A.  The Parties wish to explore [purpose, for example a possible joint project to develop a clinic booking app] (the "Purpose").')
    para(d, 'B.  For the Purpose each Party (as "Discloser") may share confidential information with the other (as "Recipient"), and both agree to protect it on these terms.', space=8)
    clause(d, 1, 'Confidential Information', [
        '"Confidential Information" means all information, in any form, that the Discloser gives to the Recipient or the Recipient learns in connection with the Purpose, including business plans, customers, prices, finances, designs, software, source code, formulas and know-how, whether or not marked confidential, if a reasonable person would understand it to be confidential, and the information described in Schedule 1.',
        'It also includes the fact that the Parties are in discussions and the terms discussed.'])
    clause(d, 2, 'Duties of the Recipient', [
        'The Recipient will use Confidential Information only for the Purpose and will not disclose it to anyone except its directors, employees and professional advisers who need to know it for the Purpose and are bound by duties of confidentiality at least as strict as these.',
        'The Recipient will protect Confidential Information with at least the care it uses for its own, and never less than reasonable care, and will not copy, reverse engineer or analyse it except as needed for the Purpose.',
        'The Recipient will tell the Discloser at once of any unauthorised use or disclosure and help to stop it, and is responsible for breaches by people it shared information with.'])
    clause(d, 3, 'Exceptions', [
        'These duties do not apply to information that the Recipient can show (a) is or becomes public other than by its breach; (b) it knew lawfully before receiving it, free of any duty of confidence; (c) it received lawfully from a third party free to disclose it; or (d) it developed independently.',
        'The Recipient may disclose Confidential Information if required by law, a court or a regulator, only to the extent required and, where lawful, after telling the Discloser first.'])
    clause(d, 4, 'Return and destruction', ['On the Discloser\'s request or when the Purpose ends, the Recipient will return or destroy all Confidential Information and copies and confirm this in writing, except one archive copy where the law requires, which stays confidential.'])
    clause(d, 5, 'No licence, no warranty, no obligation', ['Confidential Information remains the Discloser\'s property. No licence is granted except to use it for the Purpose. It is provided "as is". Neither Party must enter into any further agreement.'])
    clause(d, 6, 'Term', ['Information may be shared under this Agreement for [24] months from its date. The duties of confidentiality last for [3] years after that period ends, and for trade secrets for as long as they remain secret.'])
    clause(d, 7, 'Non-solicitation (optional)', ['During this Agreement and for [12] months after, neither Party will persuade the other\'s employees involved in the Purpose to leave. General advertisements are not a breach. [Delete this clause if not wanted.]'])
    clause(d, 8, 'Remedies', ['Damages alone may not be an adequate remedy for breach. The Discloser may seek an injunction or specific performance in addition to any other remedy.'])
    clause(d, 9, 'General', [
        'This Agreement is the whole agreement on its subject. Changes must be in writing and signed by both Parties. If any part is invalid, the rest continues.',
        'Neither Party may transfer this Agreement without the other\'s written consent. It may be signed in counterparts and by electronic signature where the law allows.'])
    clause(d, 10, 'Governing law and disputes', ['This Agreement is governed by [the laws of India / the law of England and Wales / the laws of the State of [ ] / the laws of the UAE]. The Parties will first try to settle any dispute by discussion for 30 days; after that the [courts at [city] will have exclusive jurisdiction / dispute will be decided by a sole arbitrator under [rules], seated at [city]].'])
    note(d, 'India: Arbitration and Conciliation Act, 1996. England and Wales: Arbitration Act 1996. US: name the State, e.g. "the laws of the State of New York", and AAA rules for arbitration. UAE: DIAC rules for arbitration, or the courts of the Emirate.')
    heading(d, 'Schedule 1: Information to be shared', 11.5)
    para(d, '[Describe the kinds of information, for example: product roadmap, wireframes and source code of the booking app; list of partner clinics; pricing model.]', space=10)
    para(d, 'Signed by the Parties on the date written above.', space=8)
    sign_block(d, 'For the First Party', 'For the Second Party')
    dsave(d, 'Contract-Mutual-NDA-Template.docx')

# ------------------------------------------------------------------ 4. MOU
def mou():
    d = doc_new('ct-mou-word')
    para(d, 'MEMORANDUM OF UNDERSTANDING', bold=True, size=16, align='c', space=10)
    note(d, 'Starting draft, not legal advice. Fill every [square bracket] and delete these grey lines. Clause 9 says which parts are binding: read it carefully.')
    para(d, 'This Memorandum of Understanding ("MOU") is made on [date] at [city] between:', space=6)
    for n, who in ((1, 'First Party'), (2, 'Second Party')):
        p = para(d, '(%d)  [Full legal name of organisation], having its office at [address], represented by [name, title] (the "%s"); %s' % (n, who, 'and' if n == 1 else ''), space=4)
        p.paragraph_format.left_indent = Cm(0.9); p.paragraph_format.first_line_indent = Cm(-0.9)
    heading(d, 'Background', 11.5)
    para(d, 'A.  [Who the First Party is and what it does.]'); para(d, 'B.  [Who the Second Party is and what it does.]')
    para(d, 'C.  The Parties wish to work together on [purpose] and to record their understanding before any detailed agreement.', space=8)
    clause(d, 1, 'Purpose', ['The Parties will work together on [describe the project or programme] (the "Project"), aiming to [goal, for example train 60 final-year students in digital skills by March 2027].'])
    clause(d, 2, 'Roles and responsibilities', ['Each Party will carry out the tasks listed for it in the table below, by the target dates.'])
    table(d, ['#', 'Task', 'Responsible party', 'Target date'], [['1', '[task]', '[First / Second Party]', '[date]'], ['2', '[task]', '[ ]', '[ ]']], [1, 9, 4, 3], fill_rows=4)
    clause(d, 3, 'Contact persons and reviews', ['First Party contact: [name, title, phone, email]. Second Party contact: [name, title, phone, email].', 'The contact persons will meet at least [monthly] to review progress and will record decisions in writing.'])
    clause(d, 4, 'Costs and resources', ['[Each Party will bear its own costs. / Agreed costs will be shared equally. / Costs will be shared as follows: [ ].] No Party will incur a cost on the other\'s behalf without written approval.'])
    clause(d, 5, 'Confidentiality', ['Each Party will keep confidential the other\'s non-public information received under this MOU, use it only for the Project, and continue to do so for [2] years after this MOU ends.'])
    clause(d, 6, 'Names, logos and publicity', ['Neither Party will use the other\'s name or logo, or announce the Project publicly, without the other\'s prior written approval.'])
    clause(d, 7, 'Intellectual property', ['Each Party keeps ownership of what it brings to the Project. Ownership of anything created jointly will be agreed in writing before it is used commercially.'])
    clause(d, 8, 'Duration and ending', ['This MOU starts on [date] and lasts [12 months], unless replaced by a detailed agreement or ended earlier by either Party giving [30] days\' written notice.'])
    clause(d, 9, 'Binding and non-binding parts', [
        'Clauses 4 (Costs), 5 (Confidentiality), 6 (Publicity), 8 (Duration and ending), 9 and 10 are legally binding.',
        'All other clauses record the Parties\' present intentions only and are not legally binding. Neither Party is obliged to enter into any further agreement.'])
    clause(d, 10, 'Law and disputes', ['This MOU is governed by [the laws of India / other]. The Parties will try to settle any disagreement by discussion between senior representatives; if that fails within 30 days, the [courts at [city] / arbitration under [rules] at [city]] will decide.'])
    note(d, 'In India, an MOU you intend to rely on may need stamp duty as an agreement under your state\'s Stamp Act.')
    para(d, 'Signed by the authorised representatives of the Parties on the date written above.', space=8)
    sign_block(d, 'For the First Party', 'For the Second Party')
    dsave(d, 'Contract-MOU-Template.docx')

MAKERS = [checklist, clause_library, mutual_nda, mou]
