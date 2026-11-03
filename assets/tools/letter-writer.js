/* Letters & Applications: the letter writer (toolkit.groupchaton.com/letters/letters-writer/).
   Pick what the letter is about, fill the few details that letter needs, and get a complete formal letter in English or any of
   the 22 scheduled Indian languages, Indian, UK or US layout (English only), formal / polite / firm tone. Live preview, edit by hand,
   Copy, Print / PDF, .docx and .txt. Everything runs in the browser; details are remembered only on this device. No AI, no server.
   The English text lives in DATA below, written as a plain literal (quoted keys, strings, numbers, lists) so that
   scripts/makers/letters.py can read the same templates for the Word files. Every other language is a pack in lw-lang/<code>.js
   (same keys, loaded only when picked; scripts/makers/letters.py reads lw-lang/hi.js for the Hindi Word file). Template marks:
     {key}          a field (shown as [Label] while empty)
     [[ ... ]]      optional words: dropped when any {key} inside is empty
     ((a|b))        a for a male writer, b for a female writer;  ((key:a|b)) the same by another gender field (cg, eg, pg)
   The shared engine (assets/tools.js) is made for item tables, so this tool registers with TKTools and mounts its own form. */
(function () {
  'use strict';
  var TK = window.TKTools; if (!TK) return;
  var SRC = (document.currentScript && document.currentScript.src) || '';   // lw-lang/ packs load from next to this file
  var L = TK.lib, h = L.h, store = L.store, doc = document;
  var CREDIT = TK.CREDIT || 'Designed & Developed by: Nivesh Sahayak', MADE = TK.MADE || 'toolkit.groupchaton.com';
  var DATA = /*DATA*/{
"C": [["school", "School & college", "🎓", ["Govt. Senior Secondary School", "Model Town, Ludhiana"], 1], ["office", "Office & job", "💼", ["Gupta & Co.", "Ferozepur Road, Ludhiana"], 1], ["hr", "Employer & HR letters", "🏢", ["", "House 8, Civil Lines, Ludhiana"], 0], ["bank", "Bank", "🏦", ["State Bank of India", "Model Town branch, Ludhiana"], 1], ["govt", "Government & police", "🏛️", ["", "Ludhiana"], 0], ["housing", "Housing, rent & society", "🏠", ["Sunrise Apartments Welfare Society", "Ludhiana"], 0], ["business", "Business", "📦", ["Mehta Office Supplies", "Industrial Area-A, Ludhiana"], 1], ["personal", "Personal & certificates", "✉️", ["", "Ludhiana"], 0]],
"R": {"student": {"p": ["sname", "sg"], "a": ["cls", "roll"], "x": ["cls", "roll"], "cl": "obed", "resp": 1}, "parent": {"p": ["sname", "sg", "sphone"], "a": ["child", "cg", "cls", "roll"], "x": ["_parent", "sphone"], "cl": "app", "resp": 1}, "person": {"p": ["sname", "sg", "saddr", "sphone", "semail"], "a": [], "x": ["saddr", "sphone", "semail"], "cl": "app"}, "employee": {"p": ["sname", "sg", "sdesig", "dept", "empid", "sphone"], "a": [], "x": ["sdesig", "dept", "empid"], "cl": "app"}, "org": {"p": ["sorg", "saddr", "sphone", "semail", "sname", "sg", "sdesig", "ref"], "a": [], "x": ["sdesig"], "cl": "org", "head": 1}},
"F": {
 "sname": ["Your full name", "t", "Riya Sharma"],
 "sg": ["You are", "g", "f"],
 "saddr": ["Your address", "a", "House 24, Model Town\nLudhiana, Punjab 141002"],
 "sphone": ["Mobile number", "p", "98765 43210"],
 "semail": ["Email (optional)", "e", "riya.sharma@example.com"],
 "sorg": ["Business / organisation name", "t", "Sharma Traders Pvt. Ltd."],
 "sdesig": ["Your designation", "t", "Accounts Executive"],
 "dept": ["Department", "t", "Accounts"],
 "empid": ["Employee ID", "t", "E-1042"],
 "ref": ["Reference number (optional)", "t", "STPL/HR/2026/118"],
 "cls": ["Class / course and section", "t", "VIII-B"],
 "roll": ["Roll number (optional)", "t", "17"],
 "child": ["Child's name", "t", "Aarav Sharma"],
 "cg": ["Child is", "k", "m"],
 "acno": ["Account number", "t", "30214567890"],
 "actype": ["Account type", "s", "0", ["savings", "current"]],
 "flat": ["Flat / house number", "t", "B-304"],
 "rname": ["Recipient's name (with Mr. / Ms.)", "t", "Mr. Rajesh Gupta"],
 "rto": ["To", "t", ""],
 "rorg": ["Organisation / office", "t", ""],
 "raddr": ["Address / city", "a", ""],
 "subj": ["Subject", "t", ""],
 "enc": ["Enclosures", "a", ""],
 "cc": ["Copy to", "a", ""],
 "problem": ["What needs repair", "t", ""],
 "d1": ["From date", "d", ""],
 "d2": ["To date", "d", ""],
 "by": ["Needed by (date)", "d", ""],
 "reason": ["Reason", "t", "my grandmother is unwell and I must be with her"],
 "illness": ["Illness", "t", "viral fever"],
 "event": ["Event / occasion", "t", "my elder sister's wedding"],
 "place": ["Place", "t", "Jaipur"],
 "earner": ["Earning member (e.g. My father)", "t", "My father"],
 "job": ["Their occupation", "t", "a daily-wage carpenter"],
 "work": ["Your work", "t", "a senior accountant at Gupta & Co."],
 "income": ["Monthly income", "m", "12000"],
 "conc": ["Concession wanted", "s", "0", ["full", "half"]],
 "marks": ["Marks in last exam (optional)", "t", "86%"],
 "since": ["Since", "t", "2019"],
 "session": ["Session", "t", "2026-27"],
 "purpose": ["Needed for", "t", "my scholarship application"],
 "amt": ["Amount", "m", "500"],
 "month": ["Month", "t", "October 2026"],
 "now": ["Present section / subject", "t", "Section C"],
 "want": ["Section / subject wanted", "t", "Section A"],
 "doc": ["Document lost", "t", "identity card"],
 "scheme": ["Scholarship / scheme", "t", "Post-Matric Scholarship"],
 "aclass": ["Class sought", "t", "VI"],
 "prev": ["Present school (optional)", "t", "St. Mary's Convent School, Ambala"],
 "regno": ["University registration number", "t", "PU/2023/48215"],
 "newuni": ["New university / college", "t", "the University of Delhi"],
 "what": ["What happened", "t", "coming late to school three times this week"],
 "post": ["Post / position", "t", "Accounts Assistant"],
 "ad": ["Where you saw the vacancy (optional)", "t", "The Tribune of 1 November 2026"],
 "qual": ["Your qualification", "t", "a B.Com degree from Panjab University"],
 "exp": ["Experience (optional)", "t", "two years"],
 "skills": ["Key skills (optional)", "t", "Tally Prime, GST returns, Excel and bank reconciliation"],
 "cover": ["Colleague who will handle your work (optional)", "t", "Mr. Anil Verma"],
 "ldate": ["Last working day", "d", ""],
 "notice": ["Notice period", "t", "one-month"],
 "rdate": ["Resignation date", "d", ""],
 "months": ["Number of months", "n", "6"],
 "lastrev": ["Salary last revised (optional)", "t", "April 2025"],
 "ach": ["Your main contributions (optional)", "t", "I took over GST filing for 40 clients and cut month-end closing from 7 days to 3"],
 "offerdate": ["Appointment letter date", "d", ""],
 "jdate": ["Date of joining", "d", ""],
 "edd": ["Expected delivery date", "d", ""],
 "ename": ["Employee's name", "t", "Rahul Mehra"],
 "eg": ["Employee is", "g", "m"],
 "edesig": ["Employee's designation", "t", "Sales Executive"],
 "ecode": ["Employee ID (optional)", "t", "E-1077"],
 "edept": ["Employee's department (optional)", "t", "Sales"],
 "salary": ["Monthly gross salary", "m", "28000"],
 "newsal": ["New monthly gross salary", "m", "32000"],
 "prob": ["Probation (months)", "n", "6"],
 "hours": ["Working hours (optional)", "t", "10:00 am to 6:00 pm, Monday to Saturday"],
 "issue": ["What happened / the problem", "t", ""],
 "duties": ["Main duties (optional)", "t", "handling customer accounts, billing and GST returns"],
 "project": ["Work / project done (optional)", "t", "a customer survey and a sales dashboard"],
 "toacc": ["Move the balance to account no. (optional)", "t", ""],
 "newaddr": ["New address", "a", "Flat 12, Green Avenue\nSector 70, Mohali, Punjab 160071"],
 "leaves": ["Cheque leaves", "s", "0", ["25", "50", "100"]],
 "loantype": ["Loan type", "s", "0", ["personal", "home", "education", "vehicle", "business", "gold"]],
 "card": ["Last 4 digits of the card", "t", "4821"],
 "chq": ["Cheque number", "t", "004512"],
 "chqdate": ["Cheque date", "d", ""],
 "payee": ["Cheque in favour of", "t", "Kumar Electronics"],
 "oldmob": ["Old mobile number (optional)", "p", "98140 11111"],
 "newmob": ["New mobile number", "p", "98765 43210"],
 "nominee": ["Nominee's name", "t", "Sunita Sharma"],
 "nrel": ["Relationship", "t", "Mother"],
 "nage": ["Nominee's age (optional)", "t", "52"],
 "mode": ["Type of transaction", "s", "0", ["an ATM withdrawal", "a UPI payment", "an online card payment", "a fund transfer"]],
 "txnref": ["Transaction / UTR number (optional)", "t", "630512345678"],
 "loanno": ["Loan account number", "t", "HL-778812"],
 "area": ["Locality / ward", "t", "Ward 12, Shimlapuri, Ludhiana"],
 "details": ["More details (optional)", "t", ""],
 "rproblem": ["Problem", "s", "0", ["broken road and potholes", "street lights that do not work", "broken road and street lights that do not work"]],
 "wproblem": ["Problem", "s", "0", ["no water supply", "very low water pressure", "dirty, smelly water"]],
 "eproblem": ["Problem", "s", "0", ["frequent power cuts", "an unusually high bill", "a faulty meter", "low voltage"]],
 "conno": ["Consumer / connection number", "t", "3002145678"],
 "info": ["Information wanted (one point per line)", "l", "The number of applications for new water connections received in Ward 12\nThe action taken on each application and the present status\nCopies of the file notings on my application dated 12-08-2026"],
 "fee": ["How the ₹10 fee was paid", "t", "Indian Postal Order No. 12F 345678"],
 "period": ["Period (optional)", "t", "1 April 2025 to 31 March 2026"],
 "item": ["What was stolen / lost", "t", "motorcycle (Hero Splendor, PB10 AB 1234)"],
 "value": ["Approximate value (optional)", "m", "45000"],
 "time": ["Time (optional)", "t", "9:30 pm"],
 "arn": ["Passport application number (ARN)", "t", "26-1000123456"],
 "apptdate": ["Appointment date (optional)", "d", ""],
 "docs": ["Documents attached (one per line)", "l", "Aadhaar card\nElectricity bill for October 2026\nBank passbook (first page)"],
 "uid": ["Last 4 digits of Aadhaar", "t", "4321"],
 "prop": ["Rented property", "t", "House 24, Model Town, Ludhiana"],
 "deposit": ["Security deposit", "m", "30000"],
 "pay": ["Where to send the refund (optional)", "t", "my UPI ID riya@okaxis"],
 "rent": ["Present monthly rent", "m", "15000"],
 "newrent": ["New monthly rent", "m", "16500"],
 "mfor": ["Rent pending for", "t", "September and October 2026"],
 "items": ["Items (one per line)", "l", "A4 copier paper, 75 GSM: 100 reams\nPrinter toner HP 88A: 10 pieces\nBox files: 50 pieces"],
 "dplace": ["Delivery address", "t", "our warehouse at Focal Point, Ludhiana"],
 "pono": ["Order number", "t", "PO-2026-118"],
 "qref": ["Quotation reference (optional)", "t", "Q-341 dated 28-10-2026"],
 "terms": ["Payment terms", "t", "30 days from the date of invoice"],
 "invno": ["Invoice number", "t", "INV-0458"],
 "invdate": ["Invoice date", "d", ""],
 "due": ["Due date", "d", ""],
 "bankd": ["Your bank details (optional)", "t", "HDFC Bank, A/c 50200012345678, IFSC HDFC0001234"],
 "orderno": ["Order / invoice number (optional)", "t", "ORD-9921"],
 "fix": ["How you are putting it right", "t", "we have sent the goods today by express courier and waived the delivery charge"],
 "occasion": ["What you are thanking them for", "t", "your continued orders over the past year"],
 "offer": ["Thank-you offer (optional)", "t", "we are happy to give you a 5% discount on your next order"],
 "offer2": ["What you offer", "t", "office stationery, printer supplies and housekeeping material, delivered within 24 hours in Ludhiana"],
 "clients": ["Existing clients (optional)", "t", "Gupta & Co. and City Hospital"],
 "venue": ["Venue", "t", "our new home, 45 Green Park, Ludhiana"],
 "person": ["Person's name", "t", "Aman Verma"],
 "pg": ["Person is", "g", "m"],
 "prel": ["How you know them", "t", "my student in the B.Com programme"],
 "known": ["Known for", "t", "three years"],
 "quality": ["What they are like", "t", "a bright, hardworking student who topped the class in accounting"],
 "idp": ["Their ID proof", "t", "Aadhaar card ending 7788"],
 "task": ["What they will do for you", "t", "collect my new cheque book and debit card"],
 "deceased": ["Name of the person who passed away", "t", "Shri Mohan Lal ji"],
 "drel": ["Their relation to the reader (optional)", "t", "father"],
 "product": ["Product / service", "t", "a washing machine (model WW70T)"],
 "ask": ["What you want them to do", "t", "replace the machine or refund the full amount"],
 "policy": ["Policy number", "t", "MOT/2026/554321"],
 "help": ["What they helped with", "t", "helping me prepare for my job interview"]
},
"TONE": {"req": {"formal": "I shall be thankful for your kind consideration.", "polite": "I would be very grateful if you could kindly consider my request. Thank you for your time and understanding.", "firm": "I request you to act on this at the earliest and let me know the outcome in writing."}, "comp": {"formal": "I request you to look into the matter and take the necessary action at the earliest.", "polite": "I would be grateful if you could look into this at your earliest convenience. Thank you for your help.", "firm": "If the matter is not resolved within 15 days of this letter, I will be compelled to take it up with the higher authorities and the appropriate forum without further notice."}, "info": {"formal": "Kindly take this on record.", "polite": "Kindly take this on record. Thank you for your support.", "firm": "Kindly take this on record and acknowledge receipt in writing."}, "remind": {"formal": "If you have already made the payment, please ignore this letter and share the payment details for our records.", "polite": "We value our relationship with you. If there is any issue with the invoice, please let us know so that we can resolve it quickly.", "firm": "Please note that if the payment is not received by this date, we will be forced to stop further supplies, charge interest on the overdue amount as per our terms and take steps to recover it."}},
"T": [
{"id": "sick-leave", "c": "school", "n": "Sick leave application (student)", "kw": "medical fever ill absent", "r": "student", "k": "req", "to": "The Principal", "org": "[School name]", "f": ["d1", "d2", "illness"], "s": "Application for sick leave from {d1} to {d2}", "b": ["With due respect, I beg to state that I am a student of Class {cls} of your school. Since {d1}, I have been suffering from {illness}, and the doctor has advised me complete rest.", "Therefore, I am unable to attend school from {d1} to {d2}[[ ({ndays})]]. Kindly grant me leave for this period."], "o": [["cert", "Mention the medical certificate", "I have attached the medical certificate from my doctor.", 1], ["work", "Promise to cover missed work", "I will complete all the classwork and homework that I miss.", 1]], "e": "Medical certificate"},
{"id": "leave-parent", "c": "school", "n": "Leave application by a parent", "kw": "child son daughter absent mother father guardian", "r": "parent", "k": "req", "to": "The Principal", "org": "[School name]", "f": ["d1", "d2", "reason"], "sm": {"reason": "we have to travel to Jaipur for a family wedding"}, "s": "Leave application for {child}, Class {cls}", "b": ["With due respect, I am writing to inform you that my ((cg:son|daughter)) {child}, a student of Class {cls}[[ (Roll No. {roll})]], will not be able to attend school from {d1} to {d2}[[ ({ndays})]] because {reason}.", "Kindly grant ((cg:him|her)) leave for these days."], "o": [["catch", "Promise to cover missed work", "I will make sure that ((cg:he|she)) completes the classwork and homework missed during this period.", 1], ["call", "Give your phone number", "You may reach me on {sphone} if anything is needed.", 0]]},
{"id": "leave-family", "c": "school", "n": "Leave for a family function or travel (student)", "kw": "wedding marriage function out of station urgent work", "r": "student", "k": "req", "to": "The Principal", "org": "[School name]", "f": ["event", "place", "d1", "d2"], "s": "Application for leave to attend {event}", "b": ["Respectfully, I beg to state that I am a student of Class {cls} of your school. I have to go to {place} with my family to attend {event}.", "As a result, I will not be able to come to school from {d1} to {d2}[[ ({ndays})]]. Kindly grant me leave for these days."], "o": [["work", "Promise to cover missed work", "I will complete the work missed during my absence as soon as I return.", 1]]},
{"id": "fee-concession", "c": "school", "n": "Fee concession (fee maafi) application", "kw": "fees waiver poor financial help maafi", "r": "student", "k": "req", "to": "The Principal", "org": "[School name]", "f": ["earner", "job", "income", "conc", "marks"], "s": "Request for {conc} fee concession", "b": ["With due respect, I am a student of Class {cls} in your school. {earner} works as {job}, and our family's total monthly income is only {income}. With this income it has become very difficult to pay my school fee along with the other needs of the family.", "[[I have always worked hard at my studies and scored {marks} in my last examination. ]]I wish to continue my studies in this school without a break. I therefore request you to kindly grant me {conc} fee concession for the current session."], "o": [["proof", "Mention the income certificate", "A copy of our income certificate is attached for your reference.", 1], ["promise", "Promise to keep up good results", "I assure you that I will keep up my good results and conduct.", 1]], "e": "Income certificate"},
{"id": "tc-request", "c": "school", "n": "Transfer certificate (TC) application", "kw": "school leaving certificate slc transfer", "r": "student", "k": "req", "to": "The Principal", "org": "[School name]", "f": ["since", "reason"], "lb": {"since": "Studying here since (year)"}, "sm": {"reason": "my father has been transferred to Pune"}, "s": "Application for Transfer Certificate (TC)", "b": ["Respectfully, I beg to state that I have been studying in your school since {since} and am at present a student of Class {cls}. As {reason}, I have to leave this school and take admission in another school.", "Kindly issue my Transfer Certificate at the earliest so that I can complete my new admission in time."], "o": [["dues", "Say all dues are cleared", "I have paid all my fees and returned the library books and identity card.", 1], ["cc", "Also ask for a character certificate", "I also request you to issue my character certificate along with the TC.", 1]]},
{"id": "bonafide", "c": "school", "n": "Bonafide certificate request", "kw": "bona fide student certificate scholarship bus pass", "r": "student", "k": "req", "to": "The Principal", "org": "[School / college name]", "f": ["session", "purpose", "by"], "s": "Request for a bonafide certificate", "b": ["Respectfully, I am a regular student of Class {cls}[[ (Roll No. {roll})]] in your institution for the session {session}. I need a bonafide certificate for {purpose}.", "Kindly issue the certificate at the earliest[[, as it has to be submitted by {by}]]."], "o": [["fee", "Offer to pay any fee", "I am ready to pay the prescribed fee, if any.", 0]]},
{"id": "char-school", "c": "school", "n": "Character certificate from school / college", "kw": "conduct certificate", "r": "student", "k": "req", "to": "The Principal", "org": "[School / college name]", "f": ["session", "purpose"], "sm": {"purpose": "admission to college"}, "s": "Request for a character certificate", "b": ["Respectfully, I studied in Class {cls} of your institution in the session {session}. I need a character certificate for {purpose}.", "During my time here I took part in school activities and was never involved in any act of indiscipline. Kindly issue the certificate at the earliest."]},
{"id": "fine-waiver", "c": "school", "n": "Request to waive a late fee fine", "kw": "penalty fine late fee", "r": "student", "k": "req", "to": "The Principal", "org": "[School name]", "f": ["month", "amt", "reason"], "lb": {"amt": "Fine charged"}, "sm": {"reason": "my father was in hospital that week", "amt": "200"}, "s": "Request to waive the late fee fine for {month}", "b": ["With due respect, I am a student of Class {cls}. My fee for {month} was deposited late because {reason}, and a late fine of {amt} has been charged.", "This was the first time it happened and the delay was not intentional. I request you to kindly waive the fine. I will make sure that the fee is paid on time in future."]},
{"id": "section-change", "c": "school", "n": "Change of section or subject", "kw": "stream subject change section", "r": "student", "k": "req", "to": "The Principal", "org": "[School / college name]", "f": ["now", "want", "reason"], "sm": {"reason": "my younger brother is in that section and we travel to school together"}, "s": "Request for change from {now} to {want}", "b": ["Respectfully, I am a student of Class {cls}[[ (Roll No. {roll})]] and am at present in {now}. I request a change to {want} because {reason}.", "I will complete any work or formalities needed for the change. Kindly allow the change from the earliest possible date."]},
{"id": "duplicate-doc", "c": "school", "n": "Duplicate ID card or mark sheet", "kw": "lost id card marksheet admit card library card", "r": "student", "k": "req", "to": "The Principal", "org": "[School / college name]", "f": ["doc", "d1"], "lb": {"d1": "Date lost (optional)"}, "s": "Request for a duplicate {doc}", "b": ["Respectfully, I am a student of Class {cls}[[ (Roll No. {roll})]]. I regret to inform you that I have lost my {doc}[[ on {d1}]]. I have searched everywhere but could not find it.", "Kindly issue me a duplicate {doc}. I am ready to pay the fee charged for it."]},
{"id": "scholarship", "c": "school", "n": "Scholarship application", "kw": "stipend merit financial help", "r": "student", "k": "req", "to": "The Principal", "org": "[School / college name]", "f": ["scheme", "marks", "income"], "lb": {"income": "Family's yearly income (optional)"}, "sm": {"income": "180000"}, "s": "Application for the {scheme}", "b": ["Respectfully, I am a student of Class {cls}[[ (Roll No. {roll})]] and wish to apply for the {scheme}. I scored {marks} in my last examination[[, and my family's yearly income is {income}]].", "I am sincere about my studies, and this support will help me continue my education without financial worry. Kindly consider my application and forward it to the authority concerned."], "e": "Mark sheet\nIncome certificate"},
{"id": "admission", "c": "school", "n": "Admission request by a parent", "kw": "new admission school seat nursery", "r": "parent", "a": ["child", "cg"], "k": "req", "to": "The Principal", "org": "[School name]", "f": ["aclass", "prev", "reason"], "lb": {"reason": "Why a new school (optional)"}, "sm": {"reason": "we have recently moved to this area"}, "s": "Application for admission of {child} in Class {aclass}", "b": ["Respectfully, I wish to seek admission for my ((cg:son|daughter)) {child} in Class {aclass} of your school for the coming session.[[ ((cg:He|She)) is at present studying in {prev}.]]", "[[We are looking for a new school because {reason}. ]]We have heard very good things about the teaching and discipline of your school. Kindly let me know the admission procedure, the documents required and the date of the admission test or interaction, if any."]},
{"id": "migration", "c": "school", "n": "Migration certificate (college / university)", "kw": "university migration college registrar", "r": "student", "k": "req", "to": "The Registrar", "org": "[University / college name]", "f": ["regno", "session", "newuni"], "sm": {"cls": "B.Com (final year)"}, "s": "Request for a migration certificate", "b": ["Respectfully, I was a student of {cls} at your institution[[ in the session {session}]], with University Registration No. {regno}. I have been offered admission at {newuni}, which requires a migration certificate.", "Kindly issue my migration certificate at the earliest. I have attached the required documents and the fee receipt."], "e": "Copy of the final mark sheet\nFee receipt"},
{"id": "event-permission", "c": "school", "n": "Permission to take part in an event or competition", "kw": "sports competition tournament quiz debate trip tour", "r": "student", "k": "req", "to": "The Principal", "org": "[School / college name]", "f": ["event", "place", "d1"], "lb": {"d1": "Event date"}, "sm": {"event": "the inter-school debate competition", "place": "Chandigarh"}, "s": "Request for permission to take part in {event}", "b": ["Respectfully, I am a student of Class {cls}. I have been selected to take part in {event}, which will be held at {place} on {d1}.", "Kindly allow me to take part and grant me leave from classes for the event. I will make up for any classwork that I miss."]},
{"id": "apology-school", "c": "school", "n": "Apology letter to the principal or teacher", "kw": "sorry late coming misbehaviour mistake", "r": "student", "k": "apol", "to": "The Principal", "org": "[School name]", "f": ["what", "d1"], "lb": {"d1": "Date (optional)"}, "s": "Apology for {what}", "b": ["Respectfully, I am a student of Class {cls}. I sincerely apologise for {what}[[ on {d1}]]. I realise that it was wrong and that it disturbed the discipline of the class.", "I assure you that this will not happen again. Kindly forgive me this once."]},
{"id": "job-application", "c": "office", "n": "Job application (cover letter)", "kw": "cover letter vacancy apply resume cv", "r": "person", "k": "req", "to": "The HR Manager", "org": "[Company name]", "f": ["post", "ad", "qual", "exp", "skills"], "s": "Application for the post of {post}", "b": ["I am writing to apply for the post of {post}[[ advertised in {ad}]]. I hold {qual}[[ and have {exp} of experience]].", "[[My key skills include {skills}. ]]I am hardworking, quick to learn and comfortable working in a team, and I am confident that I can contribute to your organisation from the first day.", "I have attached my résumé for your consideration and would welcome the opportunity to meet you for an interview at your convenience."], "e": "Résumé\nCopies of certificates"},
{"id": "casual-leave", "c": "office", "n": "Leave application (office)", "kw": "casual leave earned leave personal work days off", "r": "employee", "k": "req", "to": "The Manager", "org": "[Company name]", "f": ["d1", "d2", "reason", "cover"], "sm": {"reason": "I have to attend my cousin's wedding in Jaipur"}, "s": "Leave application from {d1} to {d2}", "b": ["I would like to request leave from {d1} to {d2}[[ ({ndays})]] because {reason}.", "I will complete my urgent work before going on leave[[, and {cover} will take care of my duties in my absence]]."], "o": [["reach", "Say you can be reached", "I will be available on my mobile, {sphone}, for anything urgent.", 1]]},
{"id": "sick-leave-office", "c": "office", "n": "Sick leave application (office)", "kw": "medical leave unwell fever", "r": "employee", "k": "req", "to": "The Manager", "org": "[Company name]", "f": ["illness", "d1", "d2"], "s": "Sick leave application", "b": ["I am unwell with {illness}, and the doctor has advised me to rest. I am therefore unable to come to the office from {d1} to {d2}[[ ({ndays})]].", "Kindly grant me sick leave for this period. I will resume work as soon as I recover."], "o": [["cert", "Mention the medical certificate", "I will submit the medical certificate when I rejoin.", 1]]},
{"id": "wfh", "c": "office", "n": "Work from home request", "kw": "wfh remote work", "r": "employee", "k": "req", "to": "The Manager", "org": "[Company name]", "f": ["d1", "d2", "reason"], "sm": {"reason": "my mother has had surgery and needs care at home"}, "s": "Request to work from home from {d1} to {d2}", "b": ["I request your permission to work from home from {d1} to {d2}[[ ({ndays})]] because {reason}.", "I will be online during normal office hours, join all meetings by phone or video call and complete my work on time. My laptop and internet connection are ready for remote work."]},
{"id": "resignation", "c": "office", "n": "Resignation letter", "kw": "resign quit notice period leaving job", "r": "employee", "k": "none", "to": "The Manager", "org": "[Company name]", "f": ["notice", "ldate", "reason"], "lb": {"reason": "Reason (optional)"}, "sm": {"reason": "I have accepted a role closer to my family in Chandigarh"}, "s": "Resignation from the post of {sdesig}", "b": ["Please accept this letter as formal notice of my resignation from the post of {sdesig}[[ in the {dept} department]]. As per my {notice} notice period, my last working day will be {ldate}.", "[[I have decided to move on because {reason}. ]]I am grateful for the opportunities and support I have received here, and I have learnt a great deal from you and the team.", "I will complete a proper handover of my work and documents before my last day. Kindly arrange my full and final settlement and relieving letter."], "t": {"formal": "Kindly accept my resignation.", "polite": "Thank you once again for everything. I wish the company continued success.", "firm": "Kindly acknowledge this resignation in writing and confirm my last working day."}},
{"id": "relieving-request", "c": "office", "n": "Request for relieving letter and experience certificate", "kw": "relieving experience letter full and final", "r": "employee", "k": "req", "to": "The HR Manager", "org": "[Company name]", "f": ["rdate", "ldate"], "s": "Request for relieving letter and experience certificate", "b": ["I resigned from the post of {sdesig} on {rdate}, and {ldate} was my last working day. I have handed over all my work, files and company property.", "Kindly issue my relieving letter and experience certificate, and let me know the date of my full and final settlement. These documents are needed for my next job."]},
{"id": "salary-cert-request", "c": "office", "n": "Salary certificate request", "kw": "salary slip proof income loan visa", "r": "employee", "k": "req", "to": "The HR Manager", "org": "[Company name]", "f": ["purpose", "months"], "lb": {"months": "Also ask for payslips of last … months (optional)"}, "sm": {"purpose": "a home loan application", "months": "3"}, "s": "Request for a salary certificate", "b": ["I need a salary certificate for {purpose}. Kindly issue a certificate showing my designation, date of joining and present monthly salary[[, along with my salary slips for the last {months} months]].", "The papers have to be submitted soon, so I request you to issue it within this week if possible."]},
{"id": "increment", "c": "office", "n": "Salary increment request", "kw": "raise hike appraisal pay rise", "r": "employee", "k": "req", "to": "The Manager", "org": "[Company name]", "f": ["since", "ach", "lastrev"], "lb": {"since": "Working here since"}, "sm": {"since": "June 2023"}, "s": "Request for salary increment", "b": ["I have been working as {sdesig} since {since}. During this time I have taken on more responsibility and have consistently met my targets.", "[[Among other things, {ach}. ]]My salary has not been revised[[ since {lastrev}]], while my work and the cost of living have both grown. I request you to kindly review my performance and consider a suitable increment.", "I would be happy to discuss this with you at a convenient time."]},
{"id": "salary-advance", "c": "office", "n": "Salary advance request", "kw": "advance loan from employer", "r": "employee", "k": "req", "to": "The Manager", "org": "[Company name]", "f": ["amt", "reason", "months"], "lb": {"months": "Repay in how many months"}, "sm": {"amt": "30000", "reason": "I have to pay for my father's medical treatment", "months": "6"}, "s": "Request for a salary advance of {amt}", "b": ["I request a salary advance of {amt}[[ ({amtw})]] because {reason}.", "I request that the amount be recovered from my salary in {months} equal monthly instalments, starting next month."]},
{"id": "transfer-request", "c": "office", "n": "Transfer request", "kw": "transfer branch location city", "r": "employee", "k": "req", "to": "The Manager", "org": "[Company name]", "f": ["place", "reason"], "lb": {"place": "Branch / city wanted"}, "sm": {"place": "the Chandigarh branch", "reason": "my parents live there and need my support"}, "s": "Request for transfer to {place}", "b": ["I am working as {sdesig}[[ in the {dept} department]] at this branch. I request a transfer to {place} because {reason}.", "I will continue to work with the same commitment at the new location and will hand over my present work properly before moving."]},
{"id": "noc-employer", "c": "office", "n": "NOC request from employer", "kw": "no objection passport higher studies visa travel", "r": "employee", "k": "req", "to": "The HR Manager", "org": "[Company name]", "f": ["purpose", "d1", "d2"], "lb": {"d1": "Travel from (optional)", "d2": "Travel to (optional)"}, "sm": {"purpose": "my passport application"}, "s": "Request for a No Objection Certificate (NOC)", "b": ["I need a No Objection Certificate from the company for {purpose}[[ (travel from {d1} to {d2})]].", "This will not affect my work; I will plan my duties and leave in advance. Kindly issue the NOC on the company letterhead at the earliest."]},
{"id": "joining-report", "c": "office", "n": "Joining report", "kw": "joining letter first day report duty", "r": "employee", "k": "none", "to": "The HR Manager", "org": "[Company name]", "f": ["offerdate", "jdate"], "s": "Joining report", "b": ["With reference to your appointment letter dated {offerdate}, I am pleased to report that I have joined duty as {sdesig}[[ in the {dept} department]] today, {jdate}, in the forenoon.", "I have brought the documents asked for in the letter. Kindly take my joining on record. I look forward to working with you and the team."]},
{"id": "maternity-leave", "c": "office", "n": "Maternity leave application", "kw": "pregnancy baby delivery maternity benefit", "r": "employee", "k": "req", "to": "The HR Manager", "org": "[Company name]", "f": ["edd", "d1", "d2"], "s": "Application for maternity leave", "b": ["I am expecting a baby, and my expected date of delivery is {edd}. On my doctor's advice, I request maternity leave from {d1} to {d2}[[ ({ndays})]], as allowed under the company's leave policy and the Maternity Benefit Act.", "I will hand over my ongoing work before I go on leave and keep my reporting manager informed of any change in dates."], "e": "Medical certificate from the doctor"},
{"id": "appointment", "c": "hr", "n": "Appointment letter (employer)", "kw": "offer letter job letter hire new employee", "r": "org", "k": "none", "to": "{ename}", "org": "", "sal": "Dear {ename}", "f": ["ename", "eg", "post", "edept", "jdate", "place", "salary", "prob", "hours"], "lb": {"place": "Place of work (optional)", "edept": "Department (optional)"}, "sm": {"place": "our Ludhiana office", "edept": "Sales", "post": "Sales Executive"}, "s": "Appointment as {post}", "b": ["We are pleased to appoint you as {post}[[ in the {edept} department]] of {sorg}. Your appointment is effective from {jdate}[[, and your place of work will be {place}]].", "Your gross salary will be {salary} per month[[ ({amtw})]]. Statutory deductions such as PF, ESI, professional tax and income tax will be made as per law. You will be on probation for {prob} months, after which your confirmation will be communicated in writing.", "[[Your working hours will be {hours}. ]]During your employment you will follow the rules and policies of the company and keep all business information confidential. Either side may end this appointment by giving one month's notice in writing or one month's salary in lieu of notice.", "Please sign the duplicate copy of this letter as your acceptance. We welcome you to {sorg} and wish you a rewarding career with us."], "o": [["docs", "List documents to bring on joining", "Please bring your identity and address proof, educational certificates, the relieving letter from your last employer and two photographs on the day of joining.", 1]], "oa": 3},
{"id": "warning", "c": "hr", "n": "Warning letter to an employee", "kw": "misconduct late coming discipline show cause", "r": "org", "k": "none", "to": "{ename}", "org": "{edesig}", "sal": "Dear {ename}", "f": ["ename", "eg", "edesig", "issue", "d1"], "lb": {"issue": "What the employee did", "d1": "Date of the incident (optional)"}, "sm": {"issue": "you have been reaching the office late on most days this month without informing your manager"}, "s": "Warning letter", "b": ["It has come to the management's notice that[[ on {d1}]] {issue}. This is against the rules of the company and the standard of conduct expected from every employee.", "You are hereby warned to correct this immediately. Any repetition will lead to strict disciplinary action as per company rules, which may include termination of service.", "Please treat this as a final opportunity to improve, and sign the copy of this letter as acknowledgement."], "o": [["earlier", "Say it was discussed before", "This matter was discussed with you earlier, but there has been no improvement.", 0], ["reply", "Ask for a written explanation", "Please give your written explanation within three days of receiving this letter.", 1]], "oa": 1},
{"id": "experience-cert", "c": "hr", "n": "Experience certificate (employer)", "kw": "service certificate experience letter to whom it may concern", "r": "org", "k": "none", "tw": 1, "f": ["ename", "eg", "edesig", "edept", "jdate", "ldate", "duties"], "lb": {"edept": "Department (optional)"}, "s": "Experience certificate", "b": ["This is to certify that ((eg:Mr.|Ms.)) {ename} worked with {sorg} from {jdate} to {ldate}. At the time of leaving, ((eg:he|she)) held the position of {edesig}[[ in the {edept} department]].", "[[((eg:His|Her)) main responsibilities included {duties}. ]]During ((eg:his|her)) service we found ((eg:him|her)) sincere, hardworking and honest, and ((eg:his|her)) conduct was good throughout.", "((eg:He|She)) has been relieved of ((eg:his|her)) duties after completing the handover. We wish ((eg:him|her)) every success in the future."]},
{"id": "relieving", "c": "hr", "n": "Relieving letter (employer)", "kw": "release letter resignation accepted", "r": "org", "k": "none", "to": "{ename}", "org": "{edesig}", "sal": "Dear {ename}", "f": ["ename", "eg", "ecode", "edesig", "rdate", "ldate"], "s": "Relieving letter", "b": ["This refers to your resignation letter dated {rdate}. Your resignation has been accepted, and you are relieved from the services of {sorg} at the close of working hours on {ldate}.", "We confirm that you worked with us as {edesig}[[ (Employee ID {ecode})]] and have handed over your charge and company property. Your full and final settlement will be processed as per company policy.", "We thank you for your contribution and wish you success in your future career."]},
{"id": "internship-cert", "c": "hr", "n": "Internship certificate", "kw": "intern trainee training completion", "r": "org", "k": "none", "tw": 1, "f": ["ename", "eg", "edept", "d1", "d2", "project"], "lb": {"ename": "Intern's name", "eg": "Intern is", "edept": "Department (optional)"}, "s": "Internship certificate", "b": ["This is to certify that ((eg:Mr.|Ms.)) {ename} completed an internship with {sorg}[[ in the {edept} department]] from {d1} to {d2}.", "[[During the internship ((eg:he|she)) worked on {project}. ]]((eg:He|She)) was punctual, eager to learn and worked well with the team.", "We wish ((eg:him|her)) all the best for ((eg:his|her)) studies and career."]},
{"id": "salary-cert", "c": "hr", "n": "Salary certificate (employer)", "kw": "income certificate employee salary proof", "r": "org", "k": "none", "tw": 1, "f": ["ename", "eg", "edesig", "jdate", "salary", "purpose"], "sm": {"purpose": "a home loan application"}, "s": "Salary certificate", "b": ["This is to certify that ((eg:Mr.|Ms.)) {ename} has been working with {sorg} as {edesig} since {jdate}.", "((eg:His|Her)) present gross salary is {salary} per month[[ ({amtw})]].", "This certificate is issued at ((eg:his|her)) request for {purpose} and carries no financial liability on the part of the company."]},
{"id": "increment-letter", "c": "hr", "n": "Salary revision letter (employer)", "kw": "increment letter raise hike revised salary", "r": "org", "k": "none", "to": "{ename}", "org": "", "sal": "Dear {ename}", "f": ["ename", "eg", "salary", "newsal", "d1"], "lb": {"salary": "Present monthly gross salary", "d1": "Effective from"}, "s": "Revision of salary", "b": ["We are pleased to inform you that, in recognition of your performance, your gross monthly salary has been revised from {salary} to {newsal} with effect from {d1}.", "All other terms of your employment remain the same. We appreciate your hard work and look forward to your continued contribution."]},
{"id": "account-closure", "c": "bank", "n": "Bank account closure", "kw": "close account bank", "r": "person", "k": "req", "x": ["saddr", "sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "actype", "reason", "toacc"], "lb": {"reason": "Reason (optional)"}, "sm": {"reason": "I am moving to another city"}, "s": "Request to close my {actype} account No. {acno}", "b": ["I hold a {actype} account, No. {acno}, at your branch. I wish to close this account[[ as {reason}]].", "Kindly close the account and pay me the remaining balance[[ by transfer to my account No. {toacc}]]. I am returning the unused cheque leaves and my debit card with this letter."], "o": [["kyc", "Mention the ID proof copy", "A self-attested copy of my identity proof is attached.", 1]], "e": "Unused cheque leaves\nDebit card\nCopy of ID proof"},
{"id": "bank-address-change", "c": "bank", "n": "Change of address in bank account", "kw": "address update kyc bank new address", "r": "person", "k": "req", "x": ["sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "newaddr"], "s": "Change of address in account No. {acno}", "b": ["I hold account No. {acno} at your branch. I have moved house, and my new address is:\n{newaddr}", "Kindly update this address in your records for all letters, statements and cards. A self-attested copy of my new address proof is attached."], "e": "Address proof (self-attested)"},
{"id": "cheque-book", "c": "bank", "n": "New cheque book request", "kw": "cheque book chequebook", "r": "person", "k": "req", "x": ["saddr", "sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "actype", "leaves"], "s": "Request for a new cheque book", "b": ["I hold a {actype} account, No. {acno}, at your branch. My present cheque book is almost used up. Kindly issue a new cheque book of {leaves} leaves.", "Please send it to my registered address, or let me know when I can collect it from the branch."]},
{"id": "bank-statement", "c": "bank", "n": "Bank statement request", "kw": "statement passbook account details", "r": "person", "k": "req", "x": ["saddr", "sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "d1", "d2", "purpose"], "lb": {"purpose": "Needed for (optional)"}, "sm": {"purpose": "my visa application"}, "s": "Request for account statement from {d1} to {d2}", "b": ["I hold account No. {acno} at your branch. I need a statement of this account for the period {d1} to {d2}[[ for {purpose}]].", "Kindly issue the statement duly signed and stamped by the branch. You may debit any applicable charges from my account."]},
{"id": "loan-application", "c": "bank", "n": "Loan application", "kw": "personal loan home loan education loan vehicle business", "r": "person", "k": "req", "x": ["saddr", "sphone", "semail"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["loantype", "amt", "months", "purpose", "work", "income", "acno"], "lb": {"months": "Repay in how many months", "purpose": "What the loan is for", "acno": "Your account number here (optional)", "income": "Monthly income (optional)"}, "sm": {"amt": "300000", "months": "36", "purpose": "my daughter's college fees", "income": "65000"}, "s": "Application for a {loantype} loan of {amt}", "b": ["I wish to apply for a {loantype} loan of {amt}[[ ({amtw})]] for {purpose}, to be repaid in {months} monthly instalments.", "I work as {work}[[ with a monthly income of {income}]][[, and I hold account No. {acno} at your branch]]. My income is regular and I have a good repayment record.", "Kindly let me know the interest rate, processing charges and documents required, and process my application at the earliest."], "e": "Identity and address proof\nSalary slips / income proof for the last 3 months\nBank statement for the last 6 months"},
{"id": "card-block", "c": "bank", "n": "Block lost ATM / debit card", "kw": "atm card lost stolen block hotlist", "r": "person", "k": "req", "x": ["saddr", "sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "card", "d1", "place"], "lb": {"d1": "Date lost", "place": "Where (optional)"}, "sm": {"place": "the railway station"}, "s": "Request to block my debit card ending {card}", "b": ["I hold account No. {acno} at your branch. My debit card ending in {card} was lost on {d1}[[ at {place}]].", "Kindly block the card immediately so that it cannot be misused, and issue a new card to my registered address. I will not be responsible for any transaction made with this card after it was lost."], "o": [["called", "Say you already called customer care", "I reported the loss to customer care on the same day.", 1]]},
{"id": "stop-payment", "c": "bank", "n": "Stop payment of a cheque", "kw": "cheque stop lost cheque", "r": "person", "k": "req", "x": ["saddr", "sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "chq", "chqdate", "amt", "payee", "reason"], "lb": {"reason": "Reason (optional)"}, "sm": {"amt": "18500", "reason": "the cheque has been lost in transit"}, "s": "Stop payment of cheque No. {chq}", "b": ["I hold account No. {acno} at your branch. I request you to stop payment of the following cheque:\nCheque No.: {chq}\nDate: {chqdate}\nAmount: {amt}\nIn favour of: {payee}", "[[This is needed because {reason}. ]]Please do not honour the cheque if it is presented. You may debit the stop-payment charges from my account."]},
{"id": "mobile-update", "c": "bank", "n": "Change of registered mobile number", "kw": "mobile number update otp sms alerts", "r": "person", "k": "req", "x": ["saddr"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "oldmob", "newmob"], "s": "Change of registered mobile number in account No. {acno}", "b": ["I hold account No. {acno} at your branch.[[ The mobile number now registered with it, {oldmob}, is no longer in use.]]", "Kindly register my new mobile number, {newmob}, for SMS alerts, OTPs and internet and mobile banking. A copy of my identity proof is attached."], "e": "Copy of ID proof"},
{"id": "nominee", "c": "bank", "n": "Add or change nominee in bank account", "kw": "nomination nominee", "r": "person", "k": "req", "x": ["saddr", "sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "nominee", "nrel", "nage"], "s": "Registration of nominee in account No. {acno}", "b": ["I hold account No. {acno} at your branch. I wish to register or change the nominee for this account as follows:\nName: {nominee}\nRelationship: {nrel}[[\nAge: {nage}]]", "Kindly update the nomination in your records and give me an acknowledgement. I will fill in the bank's nomination form if required."]},
{"id": "failed-txn", "c": "bank", "n": "Complaint about money debited in a failed transaction", "kw": "atm upi failed transaction refund amount debited not received", "r": "person", "k": "comp", "x": ["saddr", "sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["acno", "mode", "d1", "amt", "txnref"], "lb": {"d1": "Transaction date"}, "sm": {"amt": "5000"}, "s": "Complaint about failed transaction of {amt} on {d1}", "b": ["I hold account No. {acno} at your branch. On {d1}, I attempted {mode} of {amt}[[ (reference No. {txnref})]]. The transaction failed, but the amount was debited from my account and has not been returned so far.", "Kindly look into this and credit the amount back to my account. Under the RBI rules on failed transactions, such amounts are to be reversed within a fixed time, failing which compensation is payable to the customer."], "t": {"firm": "If the matter is not resolved within 30 days, I will be compelled to file a complaint with the RBI Ombudsman."}},
{"id": "loan-noc", "c": "bank", "n": "NOC and documents after loan closure", "kw": "loan closed no dues certificate original documents hypothecation", "r": "person", "k": "req", "x": ["saddr", "sphone"], "to": "The Branch Manager", "org": "[Bank name]", "f": ["loanno", "d1"], "lb": {"d1": "Loan closed on"}, "s": "Request for NOC and documents after closure of loan account No. {loanno}", "b": ["I had taken a loan from your branch under loan account No. {loanno}. I have repaid the entire loan, and the account was closed on {d1}.", "Kindly issue the No Objection Certificate (NOC) and return all the original documents held by the bank. Please also remove any lien or charge on the security and report the closure to the credit bureaus."], "o": [["hyp", "Vehicle loan: ask for Form 35", "As this was a vehicle loan, please also issue Form 35 for removing the hypothecation from the registration certificate.", 0]]},
{"id": "sanitation-complaint", "c": "govt", "n": "Complaint about garbage and blocked drains", "kw": "municipal corporation sanitation garbage cleanliness drain sewer nagar nigam", "r": "person", "k": "comp", "to": "The Municipal Commissioner", "org": "[Municipal Corporation / Council]", "so": "Municipal Corporation", "f": ["area", "since", "details"], "lb": {"since": "For how long (optional)"}, "sm": {"since": "two weeks", "details": "The garbage point near the park has not been cleared for days."}, "s": "Complaint about garbage and blocked drains in {area}", "b": ["I am a resident of {area}, and I wish to bring to your notice the poor state of sanitation in our locality. Garbage is not being collected regularly[[ for the last {since}]], and the drains are blocked.", "[[{details} ]]The heaps of waste and dirty water are breeding mosquitoes and flies, and there is a real risk of disease, especially for children and elderly people.", "I request you to arrange regular garbage collection, cleaning of the drains and spraying of insecticide in the area."]},
{"id": "road-light", "c": "govt", "n": "Road repair or street light complaint", "kw": "pothole road street light municipal pwd", "r": "person", "k": "comp", "to": "The Executive Engineer", "org": "[Municipal Corporation / PWD]", "so": "Municipal Corporation", "f": ["area", "rproblem", "since"], "lb": {"since": "For how long (optional)"}, "sm": {"since": "over a month"}, "s": "Repair of {rproblem} in {area}", "b": ["I wish to draw your attention to the {rproblem} in {area}[[, which have not been repaired for {since}]].", "The road has become unsafe, especially after dark. Two-wheelers slip, pedestrians and children are at risk of accidents, and dark stretches encourage theft and harassment.", "I request you to get the repair work done at the earliest."]},
{"id": "water-complaint", "c": "govt", "n": "Water supply complaint", "kw": "water supply jal board no water dirty water pressure", "r": "person", "k": "comp", "to": "The Executive Engineer", "org": "[Water supply department / Jal Board]", "so": "Water Supply and Sewerage Board", "f": ["area", "conno", "wproblem", "since"], "lb": {"conno": "Connection number (optional)", "since": "For how long (optional)"}, "sm": {"since": "five days"}, "s": "Complaint about {wproblem} in {area}", "b": ["I am a resident of {area}[[ (connection No. {conno})]]. Our area has been facing {wproblem}[[ for the last {since}]].", "Families are having to buy water or carry it from far away, and unsafe water can cause serious illness.", "Kindly look into the matter and restore a regular, clean water supply at the earliest."]},
{"id": "electricity-complaint", "c": "govt", "n": "Electricity complaint (power cuts, bill, meter)", "kw": "electricity bill meter power cut voltage discom", "r": "person", "k": "comp", "to": "The Assistant Engineer / SDO", "org": "[Electricity distribution company]", "so": "Power Distribution Corporation", "f": ["conno", "eproblem", "since", "details"], "lb": {"since": "Since (optional)"}, "sm": {"since": "last month", "details": "The bill for October shows 1,240 units, about four times my usual use."}, "s": "Complaint about {eproblem}, consumer No. {conno}", "b": ["I am a consumer of your department with consumer No. {conno}. I am writing about {eproblem} at my premises[[ since {since}]].", "[[{details} ]]Kindly send your staff to inspect the problem and set it right. If the bill or meter reading is wrong, please revise the bill and adjust any excess amount in the next bill."], "e": "Copy of the latest bill"},
{"id": "rti", "c": "govt", "n": "RTI application (Right to Information)", "kw": "rti right to information pio public information officer", "r": "person", "k": "none", "to": "The Public Information Officer", "org": "[Name of department / office]", "so": "Office of the Municipal Corporation", "f": ["info", "period", "fee"], "s": "Application under Section 6(1) of the Right to Information Act, 2005", "b": ["I request the following information under the Right to Information Act, 2005[[, for the period {period}]]:\n{info}", "I have paid the application fee of ₹10 by {fee}.", "If any part of this information is held by another public authority, kindly transfer this application to it under Section 6(3) of the Act and inform me. Please send the information to my address given below."], "o": [["bpl", "I am below the poverty line (no fee)", "I belong to the Below Poverty Line (BPL) category, so no fee is payable; a copy of my BPL card is attached.", 0], ["cert", "Ask for certified copies", "Kindly provide certified copies of the documents asked for; I will pay the copying charges.", 1]]},
{"id": "police-theft", "c": "govt", "n": "Theft complaint to police (FIR request)", "kw": "fir police theft stolen bike mobile robbery", "r": "person", "k": "comp", "to": "The Station House Officer (SHO)", "org": "[Police station name]", "so": "Police Station Division No. 5", "f": ["item", "value", "d1", "time", "place", "details"], "lb": {"d1": "Date of theft", "place": "Where it happened"}, "sm": {"place": "outside Ghumar Mandi market", "details": "I had parked it and locked the handle; when I came back after twenty minutes, it was gone."}, "s": "Request to register an FIR for theft of my {item}", "b": ["I wish to report the theft of my {item}[[, worth about {value},]] on {d1}[[ at about {time}]] {place}.", "[[{details} ]]I searched the area and asked people nearby, but could not find it.", "I request you to register an FIR, investigate the matter and help me recover my property. Kindly give me a copy of the FIR, which I need for insurance and other formalities."]},
{"id": "lost-report", "c": "govt", "n": "Lost documents or mobile report to police", "kw": "lost report gum missing documents mobile phone sim passport", "r": "person", "k": "none", "to": "The Station House Officer (SHO)", "org": "[Police station name]", "so": "Police Station Division No. 5", "f": ["item", "d1", "place", "details"], "lb": {"item": "What was lost", "d1": "Date lost", "place": "Where (place)"}, "sm": {"item": "wallet with my Aadhaar card, PAN card and driving licence", "place": "the bus stand", "details": "I was travelling by city bus from the bus stand to Model Town."}, "s": "Report of loss of {item}", "b": ["I wish to report that I lost my {item} on {d1} at or near {place}. I have searched for it but could not find it.", "[[{details} ]]Kindly record this report and give me a copy, which I need to apply for duplicates and to prevent misuse. I will inform the police station if the items are found."]},
{"id": "passport-cover", "c": "govt", "n": "Passport address proof covering letter", "kw": "passport seva kendra arn address proof covering letter", "r": "person", "k": "none", "to": "The Passport Officer", "org": "Regional Passport Office", "f": ["arn", "apptdate", "docs"], "s": "Submission of address proof for passport application, ARN {arn}", "b": ["I have applied for a passport online under Application Reference Number (ARN) {arn}[[, and my appointment is on {apptdate}]].", "As proof of my present address, I am submitting self-attested copies of the following documents:\n{docs}", "I confirm that I live at the address given below and that the details in my application are correct. Kindly accept these documents with my application."]},
{"id": "aadhaar-cover", "c": "govt", "n": "Aadhaar address update covering letter", "kw": "aadhaar uidai address update enrolment centre", "r": "person", "k": "none", "to": "The Officer-in-charge", "org": "Aadhaar Seva Kendra", "f": ["uid", "newaddr", "docs"], "s": "Request to update address in Aadhaar ending {uid}", "b": ["I wish to update the address in my Aadhaar, whose last four digits are {uid}. My new address is:\n{newaddr}", "As proof of the new address, I am attaching self-attested copies of:\n{docs}", "Kindly update my address. I will complete biometric verification at the centre, if required."]},
{"id": "landlord-repair", "c": "housing", "n": "Repair request to landlord", "kw": "tenant landlord repair leakage maintenance", "r": "person", "k": "req", "rs": "n", "x": ["sphone"], "to": "{rname}", "org": "", "f": ["prop", "problem", "since"], "lb": {"problem": "What needs repair", "since": "Since (optional)"}, "sm": {"problem": "the bathroom pipe is leaking and the wall has become damp", "since": "last week"}, "s": "Request for repairs at {prop}", "b": ["I am your tenant at {prop}. I would like to bring to your notice that {problem}[[, and this has been the case since {since}]].", "This is making daily life difficult and may cause further damage to the property if it is left as it is. I request you to arrange the repairs at the earliest. Please let me know when the workman will come so that someone is at home."], "o": [["self", "Offer to get it done and adjust in rent", "If it is easier for you, I can get the work done myself and adjust the cost, with bills, in next month's rent.", 0]]},
{"id": "vacate-notice", "c": "housing", "n": "Notice to vacate (tenant to landlord)", "kw": "vacating notice leave house rent agreement", "r": "person", "k": "none", "rs": "n", "x": ["sphone"], "to": "{rname}", "org": "", "f": ["prop", "ldate", "deposit"], "lb": {"ldate": "Date you will move out"}, "s": "Notice to vacate {prop}", "b": ["As per our rent agreement, I am giving you notice that I will vacate {prop} on {ldate}.", "I will clear the rent and all utility bills up to that date and hand over the keys with the house in good condition. Kindly arrange a joint inspection on the last day and return my security deposit of {deposit} when I hand over the keys."], "t": {"polite": "Thank you for being so helpful during my stay.", "firm": "Kindly acknowledge this notice in writing."}},
{"id": "deposit-refund", "c": "housing", "n": "Security deposit refund request", "kw": "deposit refund landlord advance", "r": "person", "k": "req", "rs": "n", "x": ["saddr", "sphone"], "to": "{rname}", "org": "", "f": ["prop", "d1", "deposit", "pay"], "lb": {"d1": "Date you moved out"}, "s": "Refund of security deposit for {prop}", "b": ["I vacated {prop} on {d1} after clearing all rent and bills, and handed over the keys to you. My security deposit of {deposit} has not been returned yet.", "Kindly refund the deposit[[ to {pay}]] at the earliest. If you plan to deduct anything for damage, please send me the details with bills."], "t": {"firm": "If the deposit is not refunded within 15 days, I will have to take legal steps to recover it."}},
{"id": "rent-increase", "c": "housing", "n": "Rent increase notice (landlord to tenant)", "kw": "rent hike revision landlord tenant", "r": "person", "k": "none", "rs": "n", "x": ["sphone"], "to": "{rname}", "org": "", "f": ["prop", "rent", "newrent", "d1"], "lb": {"d1": "New rent from"}, "s": "Revision of rent for {prop}", "b": ["As per the terms of our rent agreement, I am writing to inform you that the monthly rent for {prop} will be revised from {rent} to {newrent} with effect from {d1}.", "This revision is in line with the yearly increase agreed in the agreement and the rents now charged in the area. All other terms remain the same. Kindly confirm your acceptance in writing."], "t": {"polite": "Thank you for being a good tenant. Please feel free to talk to me if you have any questions."}},
{"id": "rent-reminder", "c": "housing", "n": "Rent payment reminder (landlord to tenant)", "kw": "rent due reminder pending tenant", "r": "person", "k": "none", "rs": "n", "x": ["sphone"], "to": "{rname}", "org": "", "f": ["prop", "mfor", "amt", "by"], "lb": {"amt": "Total due", "by": "Pay by"}, "sm": {"amt": "30000"}, "s": "Reminder for pending rent of {prop}", "b": ["This is a reminder that the rent for {prop} for {mfor} is still pending. The total amount due is {amt}.", "Kindly pay the amount by {by}."], "t": {"formal": "If you have already paid, please ignore this letter and share the payment details.", "polite": "I understand that delays can happen. If you are facing any difficulty, please talk to me so that we can work out a solution.", "firm": "Please note that if the payment is not received by this date, I will have to act as per the rent agreement, which may include adjusting the security deposit and asking you to vacate."}},
{"id": "society-complaint", "c": "housing", "n": "Complaint to housing society / RWA", "kw": "rwa society secretary parking noise water leakage neighbour", "r": "person", "k": "comp", "x": ["flat", "sphone"], "to": "The Secretary", "org": "[Society / RWA name]", "f": ["flat", "issue", "since"], "lb": {"issue": "The problem", "since": "For how long (optional)"}, "sm": {"issue": "cars being parked in front of the gate and blocking the exit", "since": "a month"}, "s": "Complaint regarding {issue}", "b": ["I live in Flat {flat} and wish to bring to the committee's notice the problem of {issue}[[, which has continued for {since}]].", "Several residents are facing the same difficulty, and it is affecting our peace and safety. I request the committee to look into it, discuss it at the next meeting if needed, and take suitable action."]},
{"id": "society-noc", "c": "housing", "n": "NOC request from housing society", "kw": "society noc sale rent loan flat transfer", "r": "person", "k": "req", "x": ["flat", "sphone"], "to": "The Secretary", "org": "[Society name]", "f": ["flat", "purpose"], "sm": {"purpose": "the sale of the flat"}, "s": "Request for a No Objection Certificate for Flat {flat}", "b": ["I am the owner of Flat {flat} in our society. I need a No Objection Certificate from the society for {purpose}.", "I confirm that all maintenance charges and other dues have been paid up to date. Kindly issue the NOC on the society letterhead at the earliest."]},
{"id": "quotation-request", "c": "business", "n": "Request for quotation", "kw": "rfq quotation enquiry price rates supplier", "r": "org", "k": "none", "to": "The Sales Manager", "org": "[Supplier name]", "f": ["items", "dplace", "by"], "lb": {"by": "Quotation needed by", "dplace": "Delivery address (optional)"}, "s": "Request for quotation", "b": ["We are interested in buying the following and request you to send your best quotation:\n{items}", "Please mention the unit price, applicable GST, delivery time, payment terms, warranty and the validity of the quotation.[[ Delivery will be required at {dplace}.]]", "We would appreciate receiving your quotation by {by}. For any clarification, please contact {sname} on {sphone}."]},
{"id": "purchase-order", "c": "business", "n": "Purchase order letter", "kw": "po order placing order supplier", "r": "org", "k": "none", "to": "The Sales Manager", "org": "[Supplier name]", "f": ["pono", "qref", "items", "amt", "by", "dplace", "terms"], "lb": {"amt": "Total order value (with GST)", "by": "Deliver by"}, "sm": {"amt": "84960"}, "s": "Purchase order No. {pono}", "b": ["We are pleased to place an order for the following[[ as per your quotation {qref}]]:\n{items}", "The total order value is {amt}[[ ({amtw})]], including GST. Please deliver the goods to {dplace} by {by}, with the tax invoice and delivery challan quoting this order number.", "Payment terms: {terms}. Goods that do not match the agreed specifications will be returned at your cost. Kindly confirm acceptance of this order."]},
{"id": "payment-reminder", "c": "business", "n": "Payment reminder to a customer", "kw": "overdue invoice dues outstanding collection follow up", "r": "org", "k": "remind", "to": "The Accounts Manager", "org": "[Customer company name]", "so": "Gupta & Co.", "f": ["invno", "invdate", "amt", "due", "by", "bankd"], "lb": {"amt": "Amount due", "by": "Pay by"}, "sm": {"amt": "118000"}, "s": "Payment reminder: invoice No. {invno}", "b": ["This is to remind you that payment of {amt} against our invoice No. {invno} dated {invdate} was due on {due} and has not been received so far.", "Kindly arrange the payment by {by}[[ to our bank account: {bankd}]]."], "o": [["msme", "Mention the MSME 45-day rule", "As a registered MSME, we are entitled under the MSMED Act, 2006 to payment within the agreed period, which cannot exceed 45 days, failing which interest is payable.", 0]]},
{"id": "apology-customer", "c": "business", "n": "Apology letter to a customer", "kw": "sorry delay wrong delivery service complaint reply", "r": "org", "k": "none", "rs": "n", "to": "{rname}", "org": "", "f": ["orderno", "issue", "fix"], "lb": {"issue": "What went wrong"}, "sm": {"issue": "the late delivery"}, "s": "Apology for {issue}", "b": ["Thank you for bringing this to our attention. We sincerely apologise for {issue}[[ in your order No. {orderno}]]. This is not the standard of service we aim to give, and we understand the inconvenience it caused you.", "To put things right, {fix}. We have also reviewed our process so that this does not happen again.", "We value your trust and hope to serve you better in future. Please contact {sname} on {sphone} if there is anything more we can do."]},
{"id": "thanks-customer", "c": "business", "n": "Thank-you letter to a customer", "kw": "thank you client appreciation loyal customer", "r": "org", "k": "none", "rs": "n", "cl": "warm", "to": "{rname}", "org": "", "f": ["occasion", "offer"], "s": "Thank you for your business", "b": ["On behalf of everyone at {sorg}, thank you for {occasion}. Customers like you are the reason our business keeps growing.", "[[As a small token of our appreciation, {offer}. ]]We look forward to serving you again. If you ever need anything, we are just a phone call away."]},
{"id": "supplier-complaint", "c": "business", "n": "Complaint to supplier about defective goods", "kw": "defective damaged goods wrong supply replacement return", "r": "org", "k": "comp", "to": "The Sales Manager", "org": "[Supplier name]", "f": ["pono", "d1", "issue", "ask"], "lb": {"d1": "Goods received on", "issue": "What was wrong", "ask": "What you want them to do"}, "sm": {"issue": "12 of the 50 cartons were damaged and the bottles inside were broken", "ask": "replace the damaged cartons within 7 days"}, "s": "Complaint about defective goods against order No. {pono}", "b": ["We received the goods against our purchase order No. {pono} on {d1}. On inspection, we found that {issue}.", "This has affected our work and our commitments to our own customers. We request you to {ask}, at your cost.", "The defective goods have been kept aside for your inspection, and photographs are attached for your reference."], "t": {"firm": "If this is not resolved within 7 days, we will be forced to return the full consignment, withhold payment and look for another supplier."}},
{"id": "order-cancel", "c": "business", "n": "Order cancellation letter", "kw": "cancel order cancellation refund advance", "r": "org", "k": "none", "to": "The Sales Manager", "org": "[Supplier name]", "f": ["pono", "d1", "reason", "amt"], "lb": {"d1": "Order date", "reason": "Reason (optional)", "amt": "Advance paid (optional)"}, "sm": {"reason": "the project for which the material was ordered has been postponed", "amt": "20000"}, "s": "Cancellation of order No. {pono}", "b": ["We regret to inform you that we wish to cancel our order No. {pono} dated {d1}[[, as {reason}]].", "As the goods have not been dispatched yet, we request you not to process the order further[[ and to refund the advance of {amt}]]. Kindly confirm the cancellation in writing.", "We apologise for any inconvenience and look forward to doing business with you again."]},
{"id": "business-intro", "c": "business", "n": "Business introduction letter", "kw": "introduce company sales letter new client", "r": "org", "k": "none", "to": "The Purchase Manager", "org": "[Company name]", "so": "City Hospital", "f": ["offer2", "since", "clients"], "lb": {"since": "In business since (optional)"}, "sm": {"since": "2012"}, "s": "Introducing {sorg}", "b": ["We would like to introduce {sorg}[[, in business since {since}]]. We offer {offer2}.", "[[Our clients include {clients}. ]]We focus on good quality, fair prices and timely delivery, and we would be glad to understand your needs and send you a proposal made for you.", "May we call on you for a short meeting at your convenience? You can reach {sname} on {sphone}."]},
{"id": "invitation", "c": "personal", "n": "Formal invitation letter", "kw": "invite function wedding housewarming ceremony party", "r": "person", "k": "none", "rs": "n", "cl": "warm", "x": ["sphone"], "to": "{rname}", "org": "", "f": ["event", "d1", "time", "venue"], "lb": {"d1": "Event date"}, "sm": {"event": "the housewarming ceremony of our new home", "time": "11:00 am"}, "s": "Invitation to {event}", "b": ["It gives us great pleasure to invite you and your family to {event} on {d1}[[ at {time}]] at {venue}.", "Your presence will make the occasion truly special for us. We look forward to welcoming you."], "o": [["meal", "Mention a meal after the function", "Please join us for a meal after the function.", 1], ["rsvp", "Ask them to confirm", "Kindly let us know by phone ({sphone}) whether you will be able to come.", 1]], "oa": 1},
{"id": "recommendation", "c": "personal", "n": "Recommendation letter", "kw": "reference letter recommend student employee lor", "r": "person", "p": ["sname", "sg", "sdesig", "sorg", "sphone", "semail"], "x": ["sdesig", "sorg", "sphone", "semail"], "k": "none", "tw": 1, "f": ["person", "pg", "prel", "known", "quality", "purpose"], "sm": {"purpose": "admission to the MBA programme at your university", "sdesig": "Associate Professor of Commerce"}, "s": "Letter of recommendation", "b": ["I am pleased to recommend ((pg:Mr.|Ms.)) {person}, whom I have known for {known} as {prel}.", "((pg:He|She)) is {quality}. I have always found ((pg:him|her)) honest, dependable and willing to take responsibility, and ((pg:he|she)) works well with people.", "I recommend ((pg:him|her)) without reservation for {purpose}. Please feel free to contact me for any further information."]},
{"id": "authority-letter", "c": "personal", "n": "Authority letter (authorise someone)", "kw": "authorization authorisation collect on behalf representative", "r": "person", "k": "none", "to": "The Branch Manager / Officer-in-charge", "org": "[Office / bank name]", "so": "State Bank of India, Model Town branch", "f": ["person", "pg", "prel", "idp", "task", "d1"], "lb": {"prel": "Their relation to you (optional)", "d1": "On date (optional)"}, "sm": {"prel": "brother"}, "s": "Authority letter in favour of {person}", "b": ["I, {sname}, hereby authorise ((pg:Mr.|Ms.)) {person}[[, my {prel},]] to {task} on my behalf[[ on {d1}]], as I am unable to come in person.", "((pg:His|Her)) identity proof is {idp}. Kindly allow ((pg:him|her)) to complete this work. ((pg:His|Her)) signature is given below for verification.", "Signature of the authorised person: ____________________"], "o": [["id", "Mention ID copies attached", "Self-attested copies of my identity proof and ((pg:his|her)) identity proof are attached.", 1]], "oa": 2},
{"id": "char-request", "c": "personal", "n": "Character certificate request (sarpanch / councillor)", "kw": "character certificate sarpanch councillor gazetted officer police verification", "r": "person", "k": "req", "to": "The Sarpanch / Ward Councillor", "org": "[Gram Panchayat / Ward No.]", "so": "Ward No. 12", "f": ["since", "purpose"], "lb": {"since": "Living here since"}, "sm": {"since": "2010", "purpose": "a government job application"}, "s": "Request for a character certificate for {purpose}", "b": ["I am a permanent resident of this area and have been living at the address given below since {since}. You know me and my family well.", "I need a character certificate for {purpose}. I have never been involved in any criminal case or unlawful activity. Kindly issue the certificate at the earliest."], "o": [["docs", "Mention Aadhaar copy and photo", "Copies of my Aadhaar card and a recent photograph are attached.", 1]]},
{"id": "condolence", "c": "personal", "n": "Condolence letter", "kw": "sympathy death passing away sorry for your loss", "r": "person", "k": "none", "rs": "n", "cl": "warm", "x": [], "to": "{rname}", "org": "", "f": ["deceased", "drel"], "s": "Heartfelt condolences", "b": ["I was deeply saddened to learn of the passing of {deceased}[[, your {drel}]]. Please accept my heartfelt condolences.", "I will always remember their kindness and warmth. Words cannot ease the pain of such a loss, but please know that you and your family are in my thoughts and prayers.", "If there is anything at all I can do to help in the days ahead, please do not hesitate to tell me."]},
{"id": "consumer-complaint", "c": "personal", "n": "Consumer complaint to a company", "kw": "defective product refund replacement warranty customer care consumer forum", "r": "person", "k": "comp", "to": "The Customer Care Manager", "org": "[Company name]", "so": "Bright Home Appliances Ltd.", "f": ["product", "d1", "orderno", "amt", "issue", "ask"], "lb": {"d1": "Date of purchase", "amt": "Price paid (optional)", "issue": "What went wrong"}, "sm": {"amt": "32500", "issue": "the machine stopped working within 20 days and the technician has not come despite two complaints"}, "s": "Complaint about {product}", "b": ["I bought {product} from you on {d1}[[ (order / invoice No. {orderno})]][[ for {amt}]]. Unfortunately, {issue}.", "I request you to {ask}. Copies of the invoice and other details are attached."], "o": [["earlier", "Say you already contacted customer care", "I reported this to your customer care earlier, but the problem has not been solved.", 1]], "oa": 1, "t": {"firm": "If the matter is not resolved within 15 days, I will file a complaint on the National Consumer Helpline (1915) and before the District Consumer Commission, entirely at your cost and risk."}},
{"id": "insurance-claim", "c": "personal", "n": "Insurance claim intimation", "kw": "insurance claim accident damage motor health policy", "r": "person", "k": "none", "to": "The Claims Manager", "org": "[Insurance company name]", "so": "Secure General Insurance Co. Ltd.", "f": ["policy", "d1", "issue", "amt"], "lb": {"d1": "Date of the incident", "issue": "What happened", "amt": "Estimated loss (optional)"}, "sm": {"issue": "my car (PB10 CD 5678) was damaged in a road accident near Khanna", "amt": "60000"}, "s": "Intimation of claim under policy No. {policy}", "b": ["I hold policy No. {policy} with your company. I wish to inform you that on {d1}, {issue}.", "[[The estimated loss is about {amt}. ]]Kindly register my claim, appoint a surveyor if needed and let me know the documents and forms required. I will co-operate fully with the claim process."]},
{"id": "thanks-personal", "c": "personal", "n": "Thank-you letter (personal)", "kw": "thanks gratitude help appreciation", "r": "person", "k": "none", "rs": "n", "cl": "warm", "x": [], "to": "{rname}", "org": "", "f": ["help"], "s": "Thank you", "b": ["I am writing to thank you sincerely for {help}. Your help came at a time when I needed it most, and it made a real difference.", "I am truly grateful for your kindness and hope I can return it some day."]}
]}/*END*/;

  /* ---------------- data helpers */
  var F = DATA.F, CATS = DATA.C, TYPES = DATA.T, TONE = DATA.TONE, ROLES = DATA.R, BY = {}, CAT = {};
  TYPES.forEach(function (t) { BY[t.id] = t; });
  CATS.forEach(function (c) { CAT[c[0]] = c; });
  var PERSIST = { sname: 1, sg: 1, saddr: 1, sphone: 1, semail: 1, sorg: 1, sdesig: 1, dept: 1, empid: 1, cls: 1, roll: 1, child: 1, cg: 1, acno: 1, actype: 1, flat: 1, regno: 1 };
  var RKEYS = { rto: 1, rorg: 1, raddr: 1, rname: 1, rsal: 1 };
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var CURS = [['₹', '₹ Rupee'], ['$', '$ Dollar'], ['£', '£ Pound'], ['€', '€ Euro'], ['AED', 'AED Dirham'], ['SAR', 'SAR Riyal']];
  var SAL = [['sm', 'Sir / Madam'], ['s', 'Sir'], ['m', 'Madam'], ['n', 'By name']];
  var LBL = { cls: 'Class: ', roll: 'Roll No.: ', dept: 'Department: ', empid: 'Employee ID: ', sphone: 'Mobile: ', semail: 'Email: ', acno: 'A/c No.: ', flat: 'Flat: ' };
  /* languages: [code, name in its own script, English name]; every one but English is a pack in lw-lang/<code>.js */
  var LANGS = /*LANGS*/[["en", "English", "English"], ["as", "অসমীয়া", "Assamese"], ["bn", "বাংলা", "Bengali"], ["doi", "डोगरी", "Dogri"], ["gu", "ગુજરાતી", "Gujarati"], ["hi", "हिन्दी", "Hindi"], ["kn", "ಕನ್ನಡ", "Kannada"], ["mai", "मैथिली", "Maithili"], ["ml", "മലയാളം", "Malayalam"], ["mr", "मराठी", "Marathi"], ["ne", "नेपाली", "Nepali"], ["or", "ଓଡ଼ିଆ", "Odia"], ["pa", "ਪੰਜਾਬੀ", "Punjabi"], ["sa", "संस्कृतम्", "Sanskrit"], ["sd", "سنڌي", "Sindhi"], ["ta", "தமிழ்", "Tamil"], ["te", "తెలుగు", "Telugu"], ["ur", "اردو", "Urdu"]]/*LANGS*/, PACKV = '202611031920', PACK = {};
  var S = Object.assign({ type: 'sick-leave', lang: 'en', lay: 'in', tone: 'formal', cur: '₹', lday: L.today() }, store('tk-letter-s') || {});
  if (S.lday < L.today()) S.lday = L.today();
  if (!LANGS.some(function (l) { return l[0] === S.lang; })) S.lang = 'en';
  var M = store('tk-letter-m') || {}, D = store('tk-letter-d') || {};
  (function () { var q = new URLSearchParams(location.search), t = q.get('type'), l = q.get('lang'); if (t && BY[t]) S.type = t; if (l && LANGS.some(function (x) { return x[0] === l; })) S.lang = l; })();
  if (!BY[S.type]) S.type = TYPES[0].id;

  /* language packs: loaded once, when the language is first picked */
  window.TKLetterLang = function (p) { PACK[p.code] = p; };
  function loadPack(code) {
    if (code === 'en' || PACK[code]) return Promise.resolve(PACK[code]);
    return new Promise(function (ok, bad) {
      var base = SRC ? SRC.replace(/letter-writer\.js[^\/]*$/, '') : '/assets/tools/';
      var s = h('script', { src: base + 'lw-lang/' + code + '.js?v=' + PACKV, async: true });
      s.onload = function () { PACK[code] ? ok(PACK[code]) : bad(new Error('The language did not load. Please try again.')); };
      s.onerror = function () { s.remove(); bad(new Error('Could not load this language. Check the internet connection and try again.')); };
      doc.head.appendChild(s);
    });
  }
  var FONTS = {};
  function webFont(p) {   // the Noto font for the script, so the preview and the printed letter look right on every device
    if (!p || !p.font || FONTS[p.font]) return; FONTS[p.font] = 1;
    doc.head.appendChild(h('link', { rel: 'stylesheet', href: 'https://fonts.googleapis.com/css2?family=' + p.font.replace(/ /g, '+') + ':wght@400;700&display=swap' }));
  }
  function pack() { return S.lang !== 'en' && PACK[S.lang] || null; }
  function lt(t) { var p = pack(); return p ? p.T[t.id] : t; }   // the letter's text in the chosen language
  function fontStack(p) { return p ? '"' + p.font + '","' + p.docx + '","Nirmala UI","Noto Sans",sans-serif' : ''; }

  function role(t) { var r = ROLES[t.r]; return { p: t.p || r.p, a: t.a || r.a, x: t.x || r.x, cl: t.cl || r.cl, resp: r.resp, head: r.head }; }
  function draft(t) { return D[t.id] || (D[t.id] = {}); }
  function rkey(t, k) { return 'r.' + (CAT[t.c][4] && (k === 'rorg' || k === 'raddr') ? t.c : t.id) + '.' + k; }
  function getv(t, k) {
    if (PERSIST[k]) return M[k];
    if (RKEYS[k]) return M[rkey(t, k)];
    return draft(t)[k];
  }
  function setv(t, k, v) {
    if (PERSIST[k]) M[k] = v; else if (RKEYS[k]) M[rkey(t, k)] = v; else draft(t)[k] = v;
  }
  function persist() { store('tk-letter-s', S); store('tk-letter-m', M); store('tk-letter-d', D); }
  function label(t, k) { return (t.lb && t.lb[k]) || (F[k] ? F[k][0] : k); }
  function sample(t, k) {
    var p = pack(), x = p && p.T[t.id];
    if (x && x.sm && x.sm[k]) return x.sm[k];
    if (p && p.F[k]) return p.F[k];
    if (t.sm && t.sm[k]) return t.sm[k];
    return F[k] ? F[k][2] : '';
  }
  function tpl(s, v) {   // a pack's frame words: {key} from v, [[ ]] dropped when a key inside is empty
    return s.replace(/\[\[([\s\S]*?)\]\]/g, function (_, inner) { return (inner.match(/\{\w+\}/g) || []).every(function (m) { return v[m.slice(1, -1)]; }) ? inner : ''; })
      .replace(/\{(\w+)\}/g, function (_, k) { return v[k] || ''; });
  }

  /* ---------------- formatting */
  function fmtDate(iso, lay) {
    var p = String(iso || '').split('-'); if (p.length !== 3) return iso || '';
    if (lay === 'in') return p[2] + '-' + p[1] + '-' + p[0];
    var m = MONTHS[+p[1] - 1] || ''; return lay === 'us' ? m + ' ' + (+p[2]) + ', ' + p[0] : (+p[2]) + ' ' + m + ' ' + p[0];
  }
  function fmtMoney(v) {
    var n = parseFloat(String(v).replace(/[,\s₹$£€]/g, '')); if (!isFinite(n)) return String(v || '');
    var o = { minimumFractionDigits: n % 1 ? 2 : 0, maximumFractionDigits: 2 };
    if (S.cur === '₹') return '₹ ' + n.toLocaleString('en-IN', o);
    var s = n.toLocaleString('en-US', o); return S.cur.length > 1 ? S.cur + ' ' + s : S.cur + s;
  }
  function daysBetween(a, b) {
    var x = Date.parse(a), y = Date.parse(b); if (!isFinite(x) || !isFinite(y) || y < x) return 0;
    return Math.round((y - x) / 864e5) + 1;
  }

  /* ---------------- the letter: fill templates, then lay out blocks */
  function filler(t, lay) {
    var p = pack();
    function val(k) {
      var raw, def = F[k];
      if (k === 'ndays') { var n = daysBetween(getv(t, 'd1'), getv(t, 'd2')); return n ? { v: p ? p.W.days[n === 1 ? 0 : 1].replace('{n}', n) : n + (n === 1 ? ' day' : ' days') } : { v: '', e: 1 }; }
      if (k === 'amtw') {   // amount in words: English only; other languages keep the amount in digits
        if (p || S.cur !== '₹') return { v: '', e: 1 };
        var mk = t.f.filter(function (x) { return F[x] && F[x][1] === 'm'; })[0], n2 = parseFloat(String(getv(t, mk) || '').replace(/[,\s₹]/g, ''));
        return isFinite(n2) && n2 > 0 ? { v: L.words(n2) } : { v: '', e: 1 };
      }
      raw = getv(t, k);
      if (raw == null || String(raw).trim() === '') {
        if (def && def[1] === 's') raw = '0'; else return { v: '[' + label(t, k).replace(/ \(.*\)$/, '') + ']', e: 1 };
      }
      raw = String(raw).trim();
      if (!def) return { v: raw };
      switch (def[1]) {
        case 'd': return { v: fmtDate(raw, lay) };
        case 'm': return { v: fmtMoney(raw) };
        case 's': var i = def[3][+raw] != null ? +raw : 0; return { v: p ? p.O[k][i] : def[3][i] };
        case 'l': return { v: raw.split(/\n+/).map(function (s) { return s.trim(); }).filter(Boolean).map(function (s, i) { return (i + 1) + '. ' + s; }).join('\n') };
        default: return { v: raw };
      }
    }
    function gender(k) { return (getv(t, k) || F[k][2]) === 'f'; }
    return function fill(s) {
      if (!s) return '';
      s = s.replace(/\[\[([\s\S]*?)\]\]/g, function (_, inner) {
        var ok = true; inner.replace(/\{(\w+)\}/g, function (__, k) { if (val(k).e) ok = false; }); return ok ? inner : '';
      });
      s = s.replace(/\(\((?:(\w+):)?([^|()]*)\|([^()]*)\)\)/g, function (_, k, a, b) { return gender(k || 'sg') ? b : a; });
      s = s.replace(/\{(\w+)\}(?=([\s\S]?))/g, function (_, k, next) {   // a typed word ending in a halant must not fuse with a suffix written after it
        var v = val(k).v; return /[्্੍્୍்్್്]$/.test(v) && /[ऀ-ൿ]/.test(next) ? v + '‌' : v;
      });
      return s;
    };
  }

  function model() {
    var t = BY[S.type], R = role(t), p = pack(), W = p && p.W, x = lt(t), lay = p ? 'in' : S.lay, fill = filler(t, lay);
    var B = [], dr = draft(t);
    function add(text, o) { if (text) B.push(Object.assign({ t: text }, o || {})); }
    function g(k) { var v = getv(t, k); return v == null ? '' : String(v).trim(); }
    function lines(v) { return String(v || '').split(/\n+/).map(function (s) { return s.trim(); }).filter(Boolean); }
    var female = (g('sg') || F.sg[2]) === 'f', named = (g('rsal') || t.rs || 'sm') === 'n';
    var subj = g('subj') || fill(x.s), orgBlank = '[' + (p ? W.org : F.sorg[0]) + ']';
    /* letterhead or sender block */
    if (R.head) {
      add(g('sorg') || orgBlank, { al: 'c', b: 1, big: 1 });
      add([lines(g('saddr')).join(', '), [g('sphone') && (p ? W.phone : 'Phone: ') + g('sphone'), g('semail')].filter(Boolean).join('  ·  ')].filter(Boolean).join('\n'), { al: 'c', sm: 1, rule: 1 });
    } else if (lay !== 'in') {
      var sb = lines(g('saddr')).concat([g('sphone'), g('semail')].filter(Boolean));
      if (sb.length) add(sb.join('\n'), { al: lay === 'uk' ? 'r' : 'l' });
    }
    var dateTxt = fmtDate(S.lday, lay);
    if (lay === 'in') add([R.head && g('ref') ? (p ? W.ref : 'Ref: ') + g('ref') : '', (p ? W.date : 'Date: ') + dateTxt].filter(Boolean).join('\n'));
    else { if (R.head && g('ref')) add('Ref: ' + g('ref')); add(dateTxt, { al: lay === 'uk' && !R.head ? 'r' : 'l' }); }
    /* recipient */
    if (!t.tw) {
      var to = [g('rto') || fill(x.to), g('rorg') || fill(x.org), g('raddr') ? lines(g('raddr')).join('\n') : '[' + (p ? W.city : 'City') + ']'].filter(Boolean);
      if (named && g('rname') && to[0] !== g('rname') && !g('rto')) to.unshift(g('rname'));
      add((lay === 'in' ? (p ? W.to : 'To,') + '\n' : '') + to.join('\n'));
    }
    /* subject and salutation */
    var sal = '', rname = g('rname') || '[' + F.rname[0].replace(/ \(.*/, '') + ']';
    if (!t.tw) {
      if (x.sal) sal = fill(x.sal);
      else if (p) sal = named ? (t.cl === 'warm' ? W.sal.warm : W.sal.n).replace('{name}', rname) : W.sal[g('rsal')] || W.sal.sm;
      else {
        var base = R.resp && lay === 'in' ? 'Respected ' : 'Dear ', rs = g('rsal') || t.rs || 'sm';
        sal = rs === 'n' ? 'Dear ' + rname : rs === 's' ? base + 'Sir' : rs === 'm' ? base + 'Madam' : (lay === 'in' ? base + 'Sir/Madam' : 'Dear Sir or Madam');
      }
      sal += p ? (W.salEnd == null ? ',' : W.salEnd) : lay === 'us' && t.cl !== 'warm' ? ':' : ',';
    }
    if (t.tw) {
      add(p ? subj : subj.toUpperCase(), { al: 'c', b: 1, u: 1 });
      add(p ? W.twimc : 'TO WHOM IT MAY CONCERN', { al: 'c', b: 1 });
    } else if (lay === 'uk') { add(sal); add(subj, { b: 1 }); }
    else { add((p ? W.subj : 'Subject: ') + subj, { b: 1 }); add(sal); }
    /* body, chosen sentences, tone */
    var paras = x.b.map(fill);
    var opts = (t.o || []).filter(function (o) { var v = dr['o_' + o[0]]; return v == null ? o[3] : v === 1; }).map(function (o) { return fill(p ? x.o[o[0]] : o[2]); }).join(' ');
    if (opts) paras.splice(t.oa != null ? t.oa : paras.length, 0, opts);
    var tones = p ? p.TONE : TONE, tone = (x.t && x.t[S.tone]) || (tones[t.k] && tones[t.k][S.tone]) || null;
    if (tone) paras.push(fill(tone));
    paras.forEach(function (s) { s = s.replace(/([^.])\.\.(?!\.)/g, '$1.').replace(/\s+([,.।॥،۔؟?!])/g, '$1').replace(/ {2,}/g, ' ').trim(); if (s) add(p ? s : s.charAt(0).toUpperCase() + s.slice(1)); });
    if (lay === 'in' && !R.head && t.cl !== 'warm' && ['req', 'comp', 'info'].indexOf(t.k) >= 0) add(p ? W.thanks : 'Thanking you,');
    /* closing and signature */
    var cl = R.cl, close;
    if (p) close = (cl === 'obed' ? W.close.obed : cl === 'warm' ? W.close.warm : cl === 'app' && (t.k === 'req' || t.k === 'comp') ? W.close.app : W.close.other)[female ? 1 : 0];
    else close = cl === 'warm' ? (lay === 'us' ? 'Warm regards,' : 'With warm regards,') : lay === 'us' ? 'Sincerely,' : cl === 'obed' && lay === 'in' ? 'Yours obediently,' : named || t.sal ? 'Yours sincerely,' : 'Yours faithfully,';
    var sig = [], lbl = p ? W.lbl : LBL;
    if (!t.tw) sig.push(close);
    if (R.head) sig.push(p ? tpl(W.forOrg, { org: g('sorg') || orgBlank }) : 'For ' + (g('sorg') || orgBlank));
    sig.push('', '', g('sname') || '[' + F.sname[0] + ']');
    R.x.forEach(function (k) {
      if (k === 'saddr') { if (lay === 'in') sig = sig.concat(lines(g('saddr'))); return; }
      if (k === '_parent') { if (g('child')) sig.push(p ? tpl(W.parent, { child: g('child'), cls: g('cls') }) : 'Parent of ' + g('child') + (g('cls') ? ', Class ' + g('cls') : '')); return; }
      if ((k === 'sphone' || k === 'semail') && lay !== 'in' && !R.head) return;
      var v = g(k); if (!v) return;
      if (k === 'actype') return;
      sig.push((lbl[k] || '') + v);
    });
    if (t.f.indexOf('acno') >= 0 && R.x.indexOf('acno') < 0 && g('acno') && t.c === 'bank') sig.push(lbl.acno + g('acno'));
    add(sig.join('\n'), { sig: 1 });
    /* enclosures and copy to */
    var enc = lines(g('enc'));
    if (enc.length) add((p ? W.encl : 'Enclosures:') + '\n' + enc.map(function (s, i) { return (i + 1) + '. ' + s; }).join('\n'), { sm: 1 });
    var cc = lines(g('cc'));
    if (cc.length) add((p ? W.cc : 'Copy to:') + '\n' + cc.map(function (s, i) { return (i + 1) + '. ' + s; }).join('\n'), { sm: 1 });
    /* right-to-left letters: keep dates, phone numbers and amounts left to right (bidi would show 03-11-2026 as 2026-11-03) */
    if (p && p.dir === 'rtl') B.forEach(function (b) { b.t = b.t.replace(/(₹\s?)?\d(?:[\d ,.\-\/:]*\d)?/g, function (m) { return '‪' + m + '‬'; }); });
    return { blocks: B, subj: subj, p: p, lay: lay, name: t.n };
  }

  /* hand edits: the edited text replaces the blocks; paragraphs that are unchanged keep their formatting */
  function current() {
    var m = model(), ed = draft(BY[S.type])._edit;
    if (ed == null) return m;
    var fmt = {}, sp = m.p ? m.p.W.subj.replace(/[\s:：]+$/, '') : '';
    m.blocks.forEach(function (b) { fmt[b.t] = b; });
    m.blocks = String(ed).replace(/\r/g, '').split(/\n[ \t]*\n/).map(function (p) { return p.replace(/^\n+|\n+$/g, ''); }).filter(function (p) { return p.trim(); })
      .map(function (p) { return Object.assign({}, fmt[p] || (/^(Subject|Sub|Re|RE)\s*:/.test(p) || (sp && p.indexOf(sp) === 0) ? { b: 1 } : {}), { t: p }); });
    m.edited = true;
    return m;
  }
  function toText(m) { return m.blocks.map(function (b) { return b.t; }).join('\n\n'); }

  /* ---------------- outputs */
  function safe(s) { return String(s || 'letter').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'letter'; }
  function fname(m) { return safe(m.name) + (m.p ? '-' + m.p.en.replace(/\W+/g, '') : '') + '-' + S.lday; }
  function save(blob, name) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name }); doc.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
  }
  function loadDocx() {
    if (window.docx) return Promise.resolve(window.docx);
    return new Promise(function (ok, bad) {
      var s = h('script', { src: 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js', async: true });
      s.onload = function () { window.docx ? ok(window.docx) : bad(new Error('The Word maker did not load.')); };
      s.onerror = function () { bad(new Error('Could not load the Word maker. Check the internet connection and try again.')); };
      doc.head.appendChild(s);
    });
  }
  function toDocx(m) {
    return loadDocx().then(function (X) {
      /* Indian scripts and Urdu are "complex scripts" in Word: the font goes in every slot (w:cs too) and the language in w:lang w:bidi */
      var p = m.p, font = p ? p.docx : 'Calibri', rtl = !!(p && p.dir === 'rtl'), lang = p ? { value: 'en-IN', bidirectional: p.lang.replace(/-[A-Z][a-z]{3}(?=-)/, '') } : undefined;
      var AL = { l: X.AlignmentType.LEFT, r: X.AlignmentType.RIGHT, c: X.AlignmentType.CENTER };
      var kids = m.blocks.map(function (b) {
        var ls = b.t.split('\n');
        return new X.Paragraph({ alignment: rtl && (b.al || 'l') === 'l' ? undefined : AL[b.al || 'l'], bidirectional: rtl || undefined, spacing: { after: b.sig ? 120 : 200, line: rtl ? 320 : 276 },
          border: b.rule ? { bottom: { color: '2563EB', space: 6, style: X.BorderStyle.SINGLE, size: 8 } } : undefined,
          children: ls.map(function (s, i) { return new X.TextRun({ text: s, break: i ? 1 : 0, bold: !!b.b, underline: b.u ? {} : undefined, size: b.big ? 32 : b.sm ? 20 : 22, font: font, color: b.big ? '1D4ED8' : undefined, rightToLeft: rtl || undefined, language: lang }); }) });
      });
      var us = m.lay === 'us';
      var d = new X.Document({ creator: 'Nivesh Sahayak', lastModifiedBy: 'Nivesh Sahayak', keywords: 'Toolkit by GroupChatOn; Nivesh Sahayak; https://www.niveshsahayak.in/',
        description: CREDIT + ' · https://www.niveshsahayak.in/', title: m.subj,
        styles: { default: { document: { run: { font: font, size: 22, language: lang } } } },
        sections: [{ properties: { page: { size: us ? { width: 12240, height: 15840 } : { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
          footers: { default: new X.Footer({ children: [new X.Paragraph({ alignment: X.AlignmentType.CENTER, children: [new X.TextRun({ text: MADE, size: 14, color: '94A3B8', font: 'Calibri' })] })] }) }, children: kids }] });
      return X.Packer.toBlob(d).then(function (blob) { save(blob, fname(m) + '.docx'); });
    });
  }
  function copy(text) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
    return new Promise(function (ok, bad) {
      var ta = h('textarea', { style: 'position:fixed;top:-1000px;opacity:0' }); ta.value = text; doc.body.appendChild(ta); ta.select();
      try { doc.execCommand('copy') ? ok() : bad(new Error('Copy did not work. Select the text and copy it yourself.')); } catch (e) { bad(e); }
      ta.remove();
    });
  }

  /* ---------------- styles for this tool (the shared toolkit.css styles the form and the paper) */
  var CSS = '.lw-pick{border:1px solid var(--line);border-radius:14px;background:var(--bg);margin:0 0 6px}' +
    '.lw-pick>summary{list-style:none;cursor:pointer;display:flex;gap:10px;align-items:center;padding:10px 12px;min-height:48px;font-weight:700}' +
    '.lw-pick>summary::-webkit-details-marker{display:none}.lw-pick>summary .lw-ic{font-size:1.4rem}.lw-pick>summary small{display:block;font-weight:500;color:var(--muted);font-size:.8rem}' +
    '.lw-pick>summary::after{content:"Change";margin-left:auto;font-size:.82rem;color:var(--brand-ink);border:1px solid var(--line);border-radius:99px;padding:3px 10px;background:var(--card)}' +
    'html[data-theme=dark] .lw-pick>summary::after{color:var(--brand)}.lw-pick[open]>summary::after{content:"Close"}' +
    '.lw-pbody{padding:0 10px 10px}.lw-search{width:100%;min-height:44px;padding:9px 12px;border-radius:10px;border:1px solid var(--line);background:var(--card);color:var(--ink);font:inherit}' +
    '.lw-list{max-height:360px;overflow:auto;margin-top:8px;padding-right:2px}.lw-list h4{margin:10px 2px 4px;font-size:.8rem;letter-spacing:.04em;text-transform:uppercase;color:var(--muted)}' +
    '.lw-t{display:block;width:100%;text-align:left;min-height:40px;padding:7px 10px;margin:3px 0;border-radius:10px;border:1px solid var(--line);background:var(--card);color:var(--ink);font:inherit;font-size:.92rem;cursor:pointer}' +
    '.lw-t small{display:block;color:var(--muted);font-size:.8rem}.lw-t:hover,.lw-t[aria-current]{border-color:var(--brand);background:var(--soft)}.lw-none{color:var(--muted);font-size:.9rem;padding:8px 2px}' +
    '.lw-note{font-size:.82rem;color:var(--muted);margin:2px 0 8px}.tk-f textarea{min-height:70px;resize:vertical}' +
    '.lw-opts{display:flex;flex-direction:column;gap:0;margin-top:6px}.lw-opts .tk-chk{min-height:40px;font-weight:500}' +
    '.lw-banner{background:var(--in);border-radius:12px;padding:10px 12px;font-size:.88rem;margin:0 0 10px;display:flex;gap:8px;flex-wrap:wrap;align-items:center}' +
    '.lw-edit{width:100%;min-height:560px;padding:16px;border-radius:8px;border:2px solid var(--brand);background:#fff;color:#14213D;font:14px/1.55 Calibri,Carlito,"Segoe UI",Arial,"Nirmala UI","Noto Sans Devanagari",sans-serif;resize:vertical}' +
    '.lw-paper{font:14px/1.55 Calibri,Carlito,"Segoe UI",Arial,sans-serif;color:#14213D;overflow-wrap:anywhere}.lw-paper.lw-xl{line-height:1.8}.lw-paper.nq,.lw-edit.nq{line-height:2.4}.lw-paper.nq{padding:0 6px}' +
    '.lw-paper p{margin:0 0 12px;white-space:pre-line}.lw-paper .sig{margin-top:4px}.lw-paper .big{font-size:20px;color:#1D4ED8;margin:0 0 2px}.lw-paper .sm{font-size:12.5px}' +
    '.lw-paper .rule{border-bottom:2px solid #2563EB;padding-bottom:8px;margin-bottom:16px}.lw-paper mark{background:#FEF08A;color:inherit;border-radius:3px;padding:0 2px}' +
    '.lw-made{color:#64748B;font-size:11px;text-align:center;margin:18px 0 0}' +
    '.tk-doc.lw-doc{padding:clamp(18px,5vw,48px) clamp(16px,5vw,56px);font-size:14px}' +
    '@media print{.lw-banner,.lw-edit{display:none!important}.lw-paper{font-size:11.5pt}.lw-paper mark{background:none;padding:0}.tk-doc.lw-doc{padding:0}}';

  /* ---------------- the form */
  function mount(box) {
    var spot = box.querySelector('.tk-mount') || box;
    doc.head.appendChild(h('style', { text: CSS }));
    var form = h('form', { 'class': 'tk-form lw-form', novalidate: true, onsubmit: function (e) { e.preventDefault(); } });
    var out = h('div', { 'class': 'tk-out' });
    var msg = h('p', { 'class': 'tk-msg', role: 'status', 'aria-live': 'polite' });
    var banner = h('div', { 'class': 'lw-banner', hidden: true });
    var paper = h('div', { 'class': 'tk-doc lw-doc', id: 'tk-doc' });
    var editor = h('textarea', { 'class': 'lw-edit', hidden: true, 'aria-label': 'Edit the letter text', spellcheck: 'true' });
    var timer, editing = false;
    function T() { return BY[S.type]; }

    function render() {
      var m = current(), t = T();
      paper.innerHTML = '';
      var lp = m.p, nq = lp && /Nastaliq/.test(lp.font) ? ' nq' : '';
      var wrap = h('div', { 'class': 'lw-paper' + (lp ? ' lw-xl' + nq : ''), lang: lp ? lp.lang : 'en', dir: lp && lp.dir === 'rtl' ? 'rtl' : null, style: lp ? 'font-family:' + fontStack(lp) : null });
      editor.className = 'lw-edit' + nq; editor.lang = lp ? lp.lang : 'en'; editor.dir = lp ? lp.dir : 'ltr'; editor.style.fontFamily = lp ? fontStack(lp) : '';
      m.blocks.forEach(function (b) {
        var p = h('p', { 'class': [b.sig ? 'sig' : '', b.big ? 'big' : '', b.sm ? 'sm' : '', b.rule ? 'rule' : ''].join(' ').trim() || null,
          style: [b.al === 'c' ? 'text-align:center' : b.al === 'r' ? 'text-align:right' : '',b.b ? 'font-weight:700' : '', b.u ? 'text-decoration:underline' : ''].filter(Boolean).join(';') || null });
        b.t.split(/(\[[^\[\]\n]{1,60}\])/).forEach(function (part, i) { p.appendChild(i % 2 ? h('mark', { text: part }) : doc.createTextNode(part)); });
        wrap.appendChild(p);
      });
      paper.appendChild(wrap);
      paper.appendChild(h('p', { 'class': 'no-print lw-made', text: MADE }));
      banner.hidden = !m.edited;
      var miss = (toText(m).match(/\[[^\[\]\n]{1,60}\]/g) || []).length;
      msg.textContent = miss ? miss + ' detail' + (miss === 1 ? '' : 's') + ' still to fill (highlighted in yellow).' : 'Letter ready: ' + t.n + '.';
    }
    function update() { clearTimeout(timer); timer = setTimeout(function () { persist(); render(); }, 120); }

    function field(t, k) {
      var def = F[k] || [k, 't'], type = def[1], p = pack(), v = getv(t, k), inp, lab = label(t, k);
      var smp = sample(t, k);
      function on(val) { setv(t, k, val); if (draft(t)._edit != null) { msg.textContent = 'Your hand edits are kept. Press "Undo my edits" to see this change.'; } update(); }
      if (type === 'g' || type === 'k' || type === 's') {
        var opts = type === 'g' ? [['m', 'Male'], ['f', 'Female']] : type === 'k' ? [['m', 'Son'], ['f', 'Daughter']] : def[3].map(function (o, i) { return [String(i), p ? o + ' · ' + p.O[k][i] : o]; });
        var cur = v == null || v === '' ? def[2] : String(v);
        inp = h('select', { onchange: function () { on(inp.value); } }, opts.map(function (o) { return h('option', { value: o[0], selected: o[0] === cur ? true : null, text: o[1] }); }));
      } else if (type === 'a' || type === 'l') {
        inp = h('textarea', { rows: type === 'l' ? 4 : 3, placeholder: smp ? 'e.g. ' + smp.split('\n')[0] : null, oninput: function () { on(inp.value); } }); inp.value = v || '';
      } else {
        inp = h('input', { type: type === 'd' ? 'date' : type === 'e' ? 'email' : type === 'p' ? 'tel' : 'text', value: v || '', placeholder: type === 'd' ? null : smp ? 'e.g. ' + smp : null,
          inputmode: type === 'm' || type === 'n' ? 'decimal' : null, autocomplete: { sname: 'name', saddr: 'street-address', sphone: 'tel', semail: 'email', sorg: 'organization' }[k] || null,
          oninput: function () { on(inp.value); } });
      }
      var MS = TK.masters, mk = { sname: 'person', saddr: 'addr', sphone: 'phone', semail: 'email', sorg: 'org', sdesig: 'desig', dept: 'dept', empid: 'empcode' }[k], pk = { rorg: 'name', raddr: 'addr', rname: 'contact' }[k];
      if (MS && TK.suggest && inp.tagName === 'INPUT' && (mk || pk)) {   // saved details from every Toolkit tool, offered as suggestions only
        if (mk) TK.suggest(inp, 'mine-' + mk, function () { return MS.values('tk-mine', mk); });
        else TK.suggest(inp, 'party-' + pk, function () { return pk === 'name' ? MS.list('tk-parties').map(function (r) { return r.name; }) : MS.values('tk-parties', pk); },
          pk === 'name' ? function (val) { var r = MS.find('tk-parties', val); if (r && r.addr && !getv(t, 'raddr')) { setv(t, 'raddr', r.addr); build(); update(); } } : null);
      }
      var wide = type === 'a' || type === 'l' || ['reason', 'issue', 'what', 'details', 'ach', 'skills', 'quality', 'task', 'fix', 'offer', 'offer2', 'occasion', 'help', 'ask', 'purpose', 'item', 'problem', 'qual', 'event', 'venue', 'prel'].indexOf(k) >= 0;
      return h('label', { 'class': 'tk-f' + (wide ? ' wide' : ''), style: wide ? 'grid-column:1/-1' : null }, [lab, inp]);
    }

    function pickBox() {
      var t = T(), det = h('details', { 'class': 'lw-pick' }), p = pack();
      function local(x) { return p ? p.T[x.id].h : ''; }   // the letter's name in the chosen language
      det.appendChild(h('summary', {}, [h('span', { 'class': 'lw-ic', 'aria-hidden': 'true', text: CAT[t.c][2] }), h('span', {}, [t.n, h('small', { text: CAT[t.c][1] + (p ? ' · ' + local(t) : '') })])]));
      var q = h('input', { type: 'search', 'class': 'lw-search', placeholder: 'Search ' + TYPES.length + ' letters: leave, fee, bank, RTI, resignation…', 'aria-label': 'Search letter types' });
      var list = h('div', { 'class': 'lw-list' });
      function draw() {
        var words = q.value.toLowerCase().split(/\s+/).filter(Boolean); list.innerHTML = ''; var n = 0;
        CATS.forEach(function (c) {
          var ts = TYPES.filter(function (x) { if (x.c !== c[0]) return false; var hay = (x.n + ' ' + local(x) + ' ' + (x.kw || '') + ' ' + c[1]).toLowerCase(); return words.every(function (w) { return hay.indexOf(w) >= 0; }); });
          if (!ts.length) return; n += ts.length;
          list.appendChild(h('h4', { text: c[2] + ' ' + c[1] }));
          ts.forEach(function (x) {
            list.appendChild(h('button', { type: 'button', 'class': 'lw-t', 'aria-current': x.id === S.type ? 'true' : null, onclick: function () { choose(x.id); } }, [x.n, p ? h('small', { text: local(x) }) : null]));
          });
        });
        if (!n) list.appendChild(h('p', { 'class': 'lw-none', text: 'No letter matches. Try fewer words, or pick the nearest one and edit the text.' }));
      }
      q.addEventListener('input', draw);
      det.addEventListener('toggle', function () { if (det.open) { draw(); setTimeout(function () { q.focus(); }, 30); } });
      det.appendChild(h('div', { 'class': 'lw-pbody' }, [q, list]));
      return det;
    }
    function choose(id) {
      S.type = id; editing = false; editBtn.lastChild.textContent = ' Edit text'; persist(); build(); render();
      setUrl();
      var f = form.querySelector('fieldset:nth-of-type(2)'); if (f && window.innerWidth < 960) f.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function setUrl() { try { history.replaceState(null, '', location.pathname + '?type=' + S.type + (S.lang !== 'en' ? '&lang=' + S.lang : '')); } catch (e) {} }
    function useLang() {   // load the chosen language's pack (and its font), then redraw
      var code = S.lang, l = LANGS.filter(function (x) { return x[0] === code; })[0];
      if (code === 'en' || PACK[code]) { webFont(PACK[code]); return; }
      msg.textContent = 'Loading ' + l[2] + '…';
      loadPack(code).then(function (p) { webFont(p); if (S.lang === code) { build(); render(); } }, function (e) { if (S.lang === code) { S.lang = 'en'; persist(); build(); render(); msg.textContent = e.message; } });
    }
    function sel(lab, key, opts, onc) {
      var s = h('select', { onchange: function () { S[key] = s.value; if (onc) onc(); persist(); build(); render(); } },
        opts.map(function (o) { return h('option', { value: o[0], selected: S[key] === o[0] ? true : null, text: o[1] }); }));
      return [h('label', { 'class': 'tk-f' }, [lab, s]), s];
    }

    function build() {
      var t = T(), R = role(t), p = pack(), x = lt(t), ln = LANGS.filter(function (l) { return l[0] === S.lang; })[0][2];
      form.innerHTML = '';
      form.appendChild(h('fieldset', {}, [h('legend', { text: '1. What is the letter about?' }), pickBox()]));
      var la = sel('Language', 'lang', LANGS.map(function (l) { return [l[0], l[0] === 'en' ? 'English' : l[1] + ' (' + l[2] + ')']; }), function () { setUrl(); useLang(); });
      var lo = sel('Layout', 'lay', [['in', 'Indian'], ['uk', 'UK / international'], ['us', 'US']], function () { S.cur = S.lay === 'uk' ? '£' : S.lay === 'us' ? '$' : '₹'; });
      if (S.lang !== 'en') { lo[1].disabled = true; lo[0].title = ln + ' letters use the Indian layout'; }
      var to = sel('Tone', 'tone', [['formal', 'Formal'], ['polite', 'Polite and warm'], ['firm', 'Firm']]);
      var cu = sel('Currency', 'cur', CURS);
      var dt = h('input', { type: 'date', value: S.lday, oninput: function () { S.lday = dt.value || L.today(); update(); } });
      form.appendChild(h('fieldset', {}, [h('legend', { text: '2. Language, layout and tone' }), h('div', { 'class': 'tk-row' }, [la[0], lo[0], to[0], cu[0], h('label', { 'class': 'tk-f' }, ['Letter date', dt])]),
        S.lang !== 'en' ? h('p', { 'class': 'lw-note', text: 'Type names, reasons and places in ' + ln + ' too, so the whole letter reads in one language (press Fill sample to see how). ' + ln + ' letters use the Indian layout; dates and amounts stay in figures.' }) : null]));
      form.appendChild(h('fieldset', {}, [h('legend', { text: '3. Your details (saved on this device)' }), h('div', { 'class': 'tk-row' }, R.p.map(function (k) { return field(t, k); }))]));
      var rs = t.sal || t.tw ? null : (function () {
        var cur = getv(t, 'rsal') || t.rs || 'sm', s = h('select', { onchange: function () { setv(t, 'rsal', s.value); persist(); build(); render(); } },
          SAL.map(function (o) { return h('option', { value: o[0], selected: o[0] === cur ? true : null, text: o[1] }); }));
        return h('label', { 'class': 'tk-f' }, ['Salutation', s]);
      })();
      if (!t.tw) {
        var named = (getv(t, 'rsal') || t.rs || 'sm') === 'n';
        var rto = field(t, 'rto'), rorg = field(t, 'rorg'), rad = field(t, 'raddr');
        rto.firstChild.textContent = 'To (name or designation)'; rto.lastChild.placeholder = (x.to || '').replace(/\{\w+\}/g, '').trim() || 'e.g. ' + sample(t, 'rname');
        rorg.firstChild.textContent = 'Organisation / office'; rorg.lastChild.placeholder = (x.org || '').replace(/\{\w+\}/g, '') || '(optional)';
        rad.firstChild.textContent = 'Address / city';
        form.appendChild(h('fieldset', {}, [h('legend', { text: '4. Sent to' }), h('div', { 'class': 'tk-row' }, [rs, named || t.sal ? field(t, 'rname') : null, rto, rorg, rad]),
          h('p', { 'class': 'lw-note', text: 'Leave "To" empty to use the usual designation shown in grey.' })]));
      }
      var fk = []; R.a.concat(t.f).forEach(function (k) { if (fk.indexOf(k) < 0 && R.p.indexOf(k) < 0) fk.push(k); });
      var about = [h('legend', { text: (t.tw ? '4' : '5') + '. About this letter' }), h('div', { 'class': 'tk-row' }, fk.map(function (k) { return field(t, k); }))];
      if (t.o && t.o.length) {
        var dr = draft(t);
        about.push(h('div', { 'class': 'lw-opts' }, t.o.map(function (o) {
          var on = dr['o_' + o[0]] == null ? !!o[3] : dr['o_' + o[0]] === 1;
          var cb = h('input', { type: 'checkbox', checked: on ? true : null, onchange: function () { dr['o_' + o[0]] = cb.checked ? 1 : 0; update(); } });
          return h('label', { 'class': 'tk-chk' }, [cb, o[1]]);
        })));
      }
      form.appendChild(h('fieldset', {}, about));
      var subj = field(t, 'subj'); subj.firstChild.textContent = t.tw ? 'Heading (optional)' : 'Subject (optional, leave empty for the automatic one)'; subj.lastChild.placeholder = '';
      var enc = field(t, 'enc'); enc.firstChild.textContent = 'Enclosures (optional, one per line)'; enc.lastChild.placeholder = t.e ? 'e.g. ' + t.e.split('\n')[0] : '';
      var cc = field(t, 'cc'); cc.firstChild.textContent = 'Copy to (optional, one per line)';
      form.appendChild(h('fieldset', {}, [h('legend', { text: (t.tw ? '5' : '6') + '. Extras' }), h('div', { 'class': 'tk-row' }, [subj, enc, cc])]));
      form.appendChild(h('div', { 'class': 'tk-acts' }, [
        h('button', { type: 'button', 'class': 'btn btn-2', text: 'Fill sample', onclick: fillSample }),
        h('button', { type: 'button', 'class': 'btn btn-2', text: 'Clear this letter', onclick: function () { D[t.id] = {}; editing = false; persist(); build(); render(); msg.textContent = 'Cleared. Your saved details are kept.'; } }),
        h('button', { type: 'button', 'class': 'btn btn-2', text: 'Forget my saved details', onclick: function () { if (!confirm('Remove your name, address and every saved detail from this device?')) return; M = {}; D = {}; persist(); build(); render(); msg.textContent = 'All saved details removed from this device.'; } })]));
      editor.hidden = !editing; paper.hidden = editing;
    }
    function fillSample() {
      var t = T(), R = role(t), p = pack(), x = lt(t), c = CAT[t.c], cs = p ? p.C[t.c] : c[3];
      R.p.concat(R.a, t.f).forEach(function (k) {
        if (k === 'ref' && !R.head) return;
        var def = F[k]; if (!def) return;
        if (PERSIST[k] && getv(t, k)) return;
        if (def[1] === 'd') {
          var past = k === 'd1' && /lost|theft|incident|received|purchase|closed|moved|order date|transaction/i.test(label(t, k));
          var d = new Date(); d.setDate(d.getDate() + (past ? -7 : { d1: 0, d2: 2, ldate: 30, by: 7, due: -15, invdate: -45, chqdate: -2, jdate: 0, offerdate: -10, rdate: -30, edd: 60, apptdate: 5 }[k] || 0));
          setv(t, k, d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2)); return; }
        setv(t, k, sample(t, k));
      });
      if (!t.tw) {
        var so = x.so || cs[0];
        if (so && !getv(t, 'rorg') && t.org) setv(t, 'rorg', so);
        if (!getv(t, 'raddr')) setv(t, 'raddr', cs[1]);
        if ((getv(t, 'rsal') || t.rs) === 'n' && !getv(t, 'rname')) setv(t, 'rname', sample(t, 'rname'));
      }
      delete draft(t)._edit; editing = false; persist(); build(); render(); msg.textContent = 'Sample filled. Replace it with your own details.';
    }

    /* output buttons */
    var icon = function (d) { return h('span', { 'aria-hidden': 'true', text: d }); };
    var editBtn = h('button', { type: 'button', 'class': 'btn btn-2', onclick: function () {
      var t = T();
      if (!editing) { editor.value = toText(current()); editing = true; editBtn.lastChild.textContent = ' Done editing'; editor.hidden = false; paper.hidden = true; editor.focus(); }
      else { editing = false; editBtn.lastChild.textContent = ' Edit text'; editor.hidden = true; paper.hidden = false; if (editor.value.trim() && editor.value !== toText(model())) draft(t)._edit = editor.value; persist(); render(); }
    } }, [icon('✏️'), ' Edit text']);
    editor.addEventListener('input', function () { draft(T())._edit = editor.value; persist(); });
    function finishEdit() { if (editing) editBtn.click(); }
    function remember() {   // keep what was typed for the suggestions in every Toolkit tool
      var MS = TK.masters, t = T(); if (!MS) return;
      function g(k) { var v = getv(t, k); return v == null ? '' : String(v).trim(); }
      var mine = { person: g('sname'), org: g('sorg'), addr: g('saddr'), phone: g('sphone'), email: g('semail'), desig: g('sdesig'), dept: g('dept'), empcode: g('empid') };
      mine.name = mine.org || mine.person; if (mine.name) MS.upsert('tk-mine', mine);
      if (g('rorg')) MS.upsert('tk-parties', { name: g('rorg'), addr: g('raddr'), contact: g('rname'), type: 'Addressee' });
    }
    function busy(b, text, job) { remember(); finishEdit(); b.disabled = true; msg.textContent = text; Promise.resolve().then(job).then(function (r) { if (r) msg.textContent = r; }, function (e) { msg.textContent = (e && e.message) || 'Something went wrong. Please try again.'; }).then(function () { b.disabled = false; }); }
    out.appendChild(h('div', { 'class': 'tk-out-acts' }, [
      h('button', { type: 'button', 'class': 'btn', onclick: function (e) { busy(e.currentTarget, 'Copying…', function () { return copy(toText(current())).then(function () { return 'Copied. Paste it into an email, WhatsApp or Word.'; }); }); } }, [icon('📋'), ' Copy']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function () { remember(); finishEdit(); persist(); (doc.fonts && doc.fonts.ready ? doc.fonts.ready : Promise.resolve()).then(function () { window.print(); }); } }, [icon('🖨'), ' Print / PDF']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function (e) { busy(e.currentTarget, 'Making the Word file…', function () { return toDocx(current()).then(function () { return 'Word file ready. Check your downloads folder.'; }); }); } }, [icon('📘'), ' Word']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function (e) { busy(e.currentTarget, '', function () { var m = current(); save(new Blob(['﻿' + toText(m).replace(/\n/g, '\r\n')], { type: 'text/plain;charset=utf-8' }), fname(m) + '.txt'); return 'Text file ready. Check your downloads folder.'; }); } }, [icon('📄'), ' Text']),
      editBtn]));
    banner.appendChild(h('span', { text: 'You changed the text by hand, so changes in the form do not show here.' }));
    banner.appendChild(h('button', { type: 'button', 'class': 'btn btn-2', style: 'min-height:36px;padding:6px 12px', text: 'Undo my edits', onclick: function () { delete draft(T())._edit; editing = false; editor.hidden = true; paper.hidden = false; editBtn.lastChild.textContent = ' Edit text'; persist(); render(); } }));
    out.appendChild(msg); out.appendChild(banner); out.appendChild(editor); out.appendChild(paper);
    build();
    spot.appendChild(h('div', { 'class': 'tk-app' }, [form, out]));
    render();
    useLang();
    doc.body.classList.add('pg-tool');
  }

  /* register with the toolkit engine, but keep its table form away from this box: it mounts .tk-tool[data-tool] on DOMContentLoaded */
  TK.register('letter-writer', { name: 'letter', custom: true, mount: mount, types: TYPES.length });
  var boxes = Array.prototype.slice.call(doc.querySelectorAll('.tk-tool[data-tool="letter-writer"]'));
  boxes.forEach(function (b) { b.removeAttribute('data-tool'); b.setAttribute('data-lw', ''); });
  function go() { boxes.forEach(mount); }
  if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', go); else go();
})();
