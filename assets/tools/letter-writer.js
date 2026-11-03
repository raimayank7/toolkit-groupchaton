/* Letters & Applications: the letter writer (toolkit.groupchaton.com/letters/letters-writer/).
   Pick what the letter is about, fill the few details that letter needs, and get a complete formal letter in English or Hindi,
   Indian, UK or US layout, formal / polite / firm tone. Live preview, edit by hand, Copy, Print / PDF, .docx and .txt.
   Everything runs in the browser; details are remembered only on this device. No AI, no server.
   The letter text lives in DATA below, written as a plain literal (quoted keys, strings, numbers, lists) so that
   scripts/makers/letters.py can read the same templates for the Word files. Template marks:
     {key}          a field (shown as [Label] while empty)
     [[ ... ]]      optional words: dropped when any {key} inside is empty
     ((a|b))        a for a male writer, b for a female writer;  ((key:a|b)) the same by another gender field (cg, eg, pg)
   The shared engine (assets/tools.js) is made for item tables, so this tool registers with TKTools and mounts its own form. */
(function () {
  'use strict';
  var TK = window.TKTools; if (!TK) return;
  var L = TK.lib, h = L.h, store = L.store, doc = document;
  var CREDIT = TK.CREDIT || 'Designed & Developed by: Nivesh Sahayak', MADE = TK.MADE || 'toolkit.groupchaton.com';
  var DATA = /*DATA*/{
"C": [
 ["school", "School & college", "🎓", ["Govt. Senior Secondary School", "राजकीय वरिष्ठ माध्यमिक विद्यालय", "Model Town, Ludhiana", "मॉडल टाउन, लुधियाना"], 1],
 ["office", "Office & job", "💼", ["Gupta & Co.", "गुप्ता एंड कंपनी", "Ferozepur Road, Ludhiana", "फिरोजपुर रोड, लुधियाना"], 1],
 ["hr", "Employer & HR letters", "🏢", ["", "", "House 8, Civil Lines, Ludhiana", "मकान नं. 8, सिविल लाइंस, लुधियाना"], 0],
 ["bank", "Bank", "🏦", ["State Bank of India", "भारतीय स्टेट बैंक", "Model Town branch, Ludhiana", "मॉडल टाउन शाखा, लुधियाना"], 1],
 ["govt", "Government & police", "🏛️", ["", "", "Ludhiana", "लुधियाना"], 0],
 ["housing", "Housing, rent & society", "🏠", ["Sunrise Apartments Welfare Society", "सनराइज़ अपार्टमेंट्स वेलफेयर सोसाइटी", "Ludhiana", "लुधियाना"], 0],
 ["business", "Business", "📦", ["Mehta Office Supplies", "मेहता ऑफिस सप्लाइज़", "Industrial Area-A, Ludhiana", "औद्योगिक क्षेत्र-ए, लुधियाना"], 1],
 ["personal", "Personal & certificates", "✉️", ["", "", "Ludhiana", "लुधियाना"], 0]
],
"R": {
 "student": {"p": ["sname", "sg"], "a": ["cls", "roll"], "x": ["cls", "roll"], "cl": "obed", "resp": 1},
 "parent": {"p": ["sname", "sg", "sphone"], "a": ["child", "cg", "cls", "roll"], "x": ["_parent", "sphone"], "cl": "app", "resp": 1},
 "person": {"p": ["sname", "sg", "saddr", "sphone", "semail"], "a": [], "x": ["saddr", "sphone", "semail"], "cl": "app"},
 "employee": {"p": ["sname", "sg", "sdesig", "dept", "empid", "sphone"], "a": [], "x": ["sdesig", "dept", "empid"], "cl": "app"},
 "org": {"p": ["sorg", "saddr", "sphone", "semail", "sname", "sg", "sdesig", "ref"], "a": [], "x": ["sdesig"], "cl": "org", "head": 1}
},
"F": {
 "sname": ["Your full name", "t", "Riya Sharma", "रिया शर्मा"],
 "sg": ["You are", "g", "f", "f"],
 "saddr": ["Your address", "a", "House 24, Model Town\nLudhiana, Punjab 141002", "मकान नं. 24, मॉडल टाउन\nलुधियाना, पंजाब 141002"],
 "sphone": ["Mobile number", "p", "98765 43210", "98765 43210"],
 "semail": ["Email (optional)", "e", "riya.sharma@example.com", "riya.sharma@example.com"],
 "sorg": ["Business / organisation name", "t", "Sharma Traders Pvt. Ltd.", "शर्मा ट्रेडर्स प्रा. लि."],
 "sdesig": ["Your designation", "t", "Accounts Executive", "लेखा अधिकारी"],
 "dept": ["Department", "t", "Accounts", "लेखा"],
 "empid": ["Employee ID", "t", "E-1042", "E-1042"],
 "ref": ["Reference number (optional)", "t", "STPL/HR/2026/118", "STPL/HR/2026/118"],
 "cls": ["Class / course and section", "t", "VIII-B", "आठवीं-ब"],
 "roll": ["Roll number (optional)", "t", "17", "17"],
 "child": ["Child's name", "t", "Aarav Sharma", "आरव शर्मा"],
 "cg": ["Child is", "k", "m", "m"],
 "acno": ["Account number", "t", "30214567890", "30214567890"],
 "actype": ["Account type", "s", "0", "0", [["savings", "बचत"], ["current", "चालू"]]],
 "flat": ["Flat / house number", "t", "B-304", "बी-304"],
 "rname": ["Recipient's name (with Mr. / Ms.)", "t", "Mr. Rajesh Gupta", "श्री राजेश गुप्ता"],
 "rto": ["To", "t", "", ""],
 "rorg": ["Organisation / office", "t", "", ""],
 "raddr": ["Address / city", "a", "", ""],
 "subj": ["Subject", "t", "", ""],
 "enc": ["Enclosures", "a", "", ""],
 "cc": ["Copy to", "a", "", ""],
 "problem": ["What needs repair", "t", "", ""],
 "d1": ["From date", "d", "", ""],
 "d2": ["To date", "d", "", ""],
 "by": ["Needed by (date)", "d", "", ""],
 "reason": ["Reason", "t", "my grandmother is unwell and I must be with her", "दादी जी की तबीयत खराब होने"],
 "illness": ["Illness", "t", "viral fever", "वायरल बुखार"],
 "event": ["Event / occasion", "t", "my elder sister's wedding", "बड़ी बहन के विवाह"],
 "place": ["Place", "t", "Jaipur", "जयपुर"],
 "earner": ["Earning member (e.g. My father)", "t", "My father", "मेरे पिताजी"],
 "job": ["Their occupation", "t", "a daily-wage carpenter", "दिहाड़ी पर बढ़ई का काम करते"],
 "work": ["Your work", "t", "a senior accountant at Gupta & Co.", "गुप्ता एंड कंपनी में वरिष्ठ लेखाकार"],
 "income": ["Monthly income", "m", "12000", "12000"],
 "conc": ["Concession wanted", "s", "0", "0", [["full", "पूर्ण"], ["half", "अर्ध"]]],
 "marks": ["Marks in last exam (optional)", "t", "86%", "86%"],
 "since": ["Since", "t", "2019", "2019"],
 "session": ["Session", "t", "2026-27", "2026-27"],
 "purpose": ["Needed for", "t", "my scholarship application", "छात्रवृत्ति आवेदन"],
 "amt": ["Amount", "m", "500", "500"],
 "month": ["Month", "t", "October 2026", "अक्टूबर 2026"],
 "now": ["Present section / subject", "t", "Section C", "सेक्शन सी"],
 "want": ["Section / subject wanted", "t", "Section A", "सेक्शन ए"],
 "doc": ["Document lost", "t", "identity card", "पहचान पत्र"],
 "scheme": ["Scholarship / scheme", "t", "Post-Matric Scholarship", "पोस्ट-मैट्रिक छात्रवृत्ति"],
 "aclass": ["Class sought", "t", "VI", "छठी"],
 "prev": ["Present school (optional)", "t", "St. Mary's Convent School, Ambala", "सेंट मैरी कॉन्वेंट स्कूल, अंबाला"],
 "regno": ["University registration number", "t", "PU/2023/48215", "PU/2023/48215"],
 "newuni": ["New university / college", "t", "the University of Delhi", "दिल्ली विश्वविद्यालय"],
 "what": ["What happened", "t", "coming late to school three times this week", "इस सप्ताह तीन बार देर से विद्यालय आने"],
 "post": ["Post / position", "t", "Accounts Assistant", "लेखा सहायक"],
 "ad": ["Where you saw the vacancy (optional)", "t", "The Tribune of 1 November 2026", "1 नवंबर 2026 के दैनिक भास्कर"],
 "qual": ["Your qualification", "t", "a B.Com degree from Panjab University", "पंजाब विश्वविद्यालय से बी.कॉम"],
 "exp": ["Experience (optional)", "t", "two years", "दो वर्ष"],
 "skills": ["Key skills (optional)", "t", "Tally Prime, GST returns, Excel and bank reconciliation", "टैली प्राइम, जीएसटी रिटर्न, एक्सेल और बैंक मिलान"],
 "cover": ["Colleague who will handle your work (optional)", "t", "Mr. Anil Verma", "श्री अनिल वर्मा"],
 "ldate": ["Last working day", "d", "", ""],
 "notice": ["Notice period", "t", "one-month", "एक माह"],
 "rdate": ["Resignation date", "d", "", ""],
 "months": ["Number of months", "n", "6", "6"],
 "lastrev": ["Salary last revised (optional)", "t", "April 2025", "अप्रैल 2025"],
 "ach": ["Your main contributions (optional)", "t", "I took over GST filing for 40 clients and cut month-end closing from 7 days to 3", "40 ग्राहकों का जीएसटी फाइलिंग कार्य सँभाला और महीने के अंत की क्लोजिंग 7 दिन से घटाकर 3 दिन की"],
 "offerdate": ["Appointment letter date", "d", "", ""],
 "jdate": ["Date of joining", "d", "", ""],
 "edd": ["Expected delivery date", "d", "", ""],
 "ename": ["Employee's name", "t", "Rahul Mehra", "राहुल मेहरा"],
 "eg": ["Employee is", "g", "m", "m"],
 "edesig": ["Employee's designation", "t", "Sales Executive", "विक्रय अधिकारी"],
 "ecode": ["Employee ID (optional)", "t", "E-1077", "E-1077"],
 "edept": ["Employee's department (optional)", "t", "Sales", "विक्रय"],
 "salary": ["Monthly gross salary", "m", "28000", "28000"],
 "newsal": ["New monthly gross salary", "m", "32000", "32000"],
 "prob": ["Probation (months)", "n", "6", "6"],
 "hours": ["Working hours (optional)", "t", "10:00 am to 6:00 pm, Monday to Saturday", "सोमवार से शनिवार, सुबह 10 से शाम 6 बजे तक"],
 "issue": ["What happened / the problem", "t", "", ""],
 "duties": ["Main duties (optional)", "t", "handling customer accounts, billing and GST returns", "ग्राहक खाते, बिलिंग और जीएसटी रिटर्न"],
 "project": ["Work / project done (optional)", "t", "a customer survey and a sales dashboard", "ग्राहक सर्वेक्षण और बिक्री डैशबोर्ड"],
 "toacc": ["Move the balance to account no. (optional)", "t", "", ""],
 "newaddr": ["New address", "a", "Flat 12, Green Avenue\nSector 70, Mohali, Punjab 160071", "फ्लैट 12, ग्रीन एवेन्यू\nसेक्टर 70, मोहाली, पंजाब 160071"],
 "leaves": ["Cheque leaves", "s", "0", "0", [["25", "25"], ["50", "50"], ["100", "100"]]],
 "loantype": ["Loan type", "s", "0", "0", [["personal", "व्यक्तिगत"], ["home", "गृह"], ["education", "शिक्षा"], ["vehicle", "वाहन"], ["business", "व्यवसाय"], ["gold", "स्वर्ण"]]],
 "card": ["Last 4 digits of the card", "t", "4821", "4821"],
 "chq": ["Cheque number", "t", "004512", "004512"],
 "chqdate": ["Cheque date", "d", "", ""],
 "payee": ["Cheque in favour of", "t", "Kumar Electronics", "कुमार इलेक्ट्रॉनिक्स"],
 "oldmob": ["Old mobile number (optional)", "p", "98140 11111", "98140 11111"],
 "newmob": ["New mobile number", "p", "98765 43210", "98765 43210"],
 "nominee": ["Nominee's name", "t", "Sunita Sharma", "सुनीता शर्मा"],
 "nrel": ["Relationship", "t", "Mother", "माता"],
 "nage": ["Nominee's age (optional)", "t", "52", "52"],
 "mode": ["Type of transaction", "s", "0", "0", [["an ATM withdrawal", "एटीएम निकासी"], ["a UPI payment", "यूपीआई भुगतान"], ["an online card payment", "ऑनलाइन कार्ड भुगतान"], ["a fund transfer", "फंड ट्रांसफर"]]],
 "txnref": ["Transaction / UTR number (optional)", "t", "630512345678", "630512345678"],
 "loanno": ["Loan account number", "t", "HL-778812", "HL-778812"],
 "area": ["Locality / ward", "t", "Ward 12, Shimlapuri, Ludhiana", "वार्ड 12, शिमलापुरी, लुधियाना"],
 "details": ["More details (optional)", "t", "", ""],
 "rproblem": ["Problem", "s", "0", "0", [["broken road and potholes", "टूटी सड़क और गड्ढों"], ["street lights that do not work", "खराब स्ट्रीट लाइटों"], ["broken road and street lights that do not work", "टूटी सड़क और खराब स्ट्रीट लाइटों"]]],
 "wproblem": ["Problem", "s", "0", "0", [["no water supply", "पानी की आपूर्ति बंद होने"], ["very low water pressure", "पानी का दबाव बहुत कम होने"], ["dirty, smelly water", "गंदा और बदबूदार पानी आने"]]],
 "eproblem": ["Problem", "s", "0", "0", [["frequent power cuts", "बार-बार बिजली कटौती"], ["an unusually high bill", "असामान्य रूप से अधिक बिल"], ["a faulty meter", "खराब मीटर"], ["low voltage", "कम वोल्टेज"]]],
 "conno": ["Consumer / connection number", "t", "3002145678", "3002145678"],
 "info": ["Information wanted (one point per line)", "l", "The number of applications for new water connections received in Ward 12\nThe action taken on each application and the present status\nCopies of the file notings on my application dated 12-08-2026", "वार्ड 12 में नए पानी कनेक्शन के लिए प्राप्त आवेदनों की संख्या\nप्रत्येक आवेदन पर की गई कार्रवाई और वर्तमान स्थिति\nमेरे दिनांक 12-08-2026 के आवेदन पर फाइल टिप्पणियों की प्रतियाँ"],
 "fee": ["How the ₹10 fee was paid", "t", "Indian Postal Order No. 12F 345678", "भारतीय पोस्टल ऑर्डर संख्या 12F 345678"],
 "period": ["Period (optional)", "t", "1 April 2025 to 31 March 2026", "1 अप्रैल 2025 से 31 मार्च 2026"],
 "item": ["What was stolen / lost", "t", "motorcycle (Hero Splendor, PB10 AB 1234)", "मेरी मोटरसाइकिल (हीरो स्प्लेंडर, PB10 AB 1234)"],
 "value": ["Approximate value (optional)", "m", "45000", "45000"],
 "time": ["Time (optional)", "t", "9:30 pm", "रात 9:30"],
 "arn": ["Passport application number (ARN)", "t", "26-1000123456", "26-1000123456"],
 "apptdate": ["Appointment date (optional)", "d", "", ""],
 "docs": ["Documents attached (one per line)", "l", "Aadhaar card\nElectricity bill for October 2026\nBank passbook (first page)", "आधार कार्ड\nअक्टूबर 2026 का बिजली बिल\nबैंक पासबुक (पहला पृष्ठ)"],
 "uid": ["Last 4 digits of Aadhaar", "t", "4321", "4321"],
 "prop": ["Rented property", "t", "House 24, Model Town, Ludhiana", "मकान नं. 24, मॉडल टाउन, लुधियाना"],
 "deposit": ["Security deposit", "m", "30000", "30000"],
 "pay": ["Where to send the refund (optional)", "t", "my UPI ID riya@okaxis", "यूपीआई आईडी riya@okaxis पर"],
 "rent": ["Present monthly rent", "m", "15000", "15000"],
 "newrent": ["New monthly rent", "m", "16500", "16500"],
 "mfor": ["Rent pending for", "t", "September and October 2026", "सितंबर और अक्टूबर 2026"],
 "items": ["Items (one per line)", "l", "A4 copier paper, 75 GSM: 100 reams\nPrinter toner HP 88A: 10 pieces\nBox files: 50 pieces", "A4 कॉपियर पेपर, 75 जीएसएम: 100 रिम\nप्रिंटर टोनर एचपी 88A: 10 नग\nबॉक्स फाइल: 50 नग"],
 "dplace": ["Delivery address", "t", "our warehouse at Focal Point, Ludhiana", "फोकल पॉइंट, लुधियाना स्थित हमारे गोदाम"],
 "pono": ["Order number", "t", "PO-2026-118", "PO-2026-118"],
 "qref": ["Quotation reference (optional)", "t", "Q-341 dated 28-10-2026", "Q-341 दिनांक 28-10-2026"],
 "terms": ["Payment terms", "t", "30 days from the date of invoice", "इनवॉइस की तिथि से 30 दिन"],
 "invno": ["Invoice number", "t", "INV-0458", "INV-0458"],
 "invdate": ["Invoice date", "d", "", ""],
 "due": ["Due date", "d", "", ""],
 "bankd": ["Your bank details (optional)", "t", "HDFC Bank, A/c 50200012345678, IFSC HDFC0001234", "एचडीएफसी बैंक, खाता 50200012345678, IFSC HDFC0001234"],
 "orderno": ["Order / invoice number (optional)", "t", "ORD-9921", "ORD-9921"],
 "fix": ["How you are putting it right", "t", "we have sent the goods today by express courier and waived the delivery charge", "हमने आज एक्सप्रेस कूरियर से सामान भेज दिया है और डिलीवरी शुल्क माफ कर दिया है"],
 "occasion": ["What you are thanking them for", "t", "your continued orders over the past year", "पिछले वर्ष भर लगातार ऑर्डर देने"],
 "offer": ["Thank-you offer (optional)", "t", "we are happy to give you a 5% discount on your next order", "आपके अगले ऑर्डर पर हम 5% छूट दे रहे हैं"],
 "offer2": ["What you offer", "t", "office stationery, printer supplies and housekeeping material, delivered within 24 hours in Ludhiana", "कार्यालय स्टेशनरी, प्रिंटर सामग्री और हाउसकीपिंग सामान, लुधियाना में 24 घंटे के भीतर डिलीवरी सहित"],
 "clients": ["Existing clients (optional)", "t", "Gupta & Co. and City Hospital", "गुप्ता एंड कंपनी और सिटी हॉस्पिटल"],
 "venue": ["Venue", "t", "our new home, 45 Green Park, Ludhiana", "हमारे नए घर, 45 ग्रीन पार्क, लुधियाना"],
 "person": ["Person's name", "t", "Aman Verma", "अमन वर्मा"],
 "pg": ["Person is", "g", "m", "m"],
 "prel": ["How you know them", "t", "my student in the B.Com programme", "बी.कॉम कार्यक्रम में मेरे विद्यार्थी"],
 "known": ["Known for", "t", "three years", "तीन वर्षों"],
 "quality": ["What they are like", "t", "a bright, hardworking student who topped the class in accounting", "एक प्रतिभाशाली और परिश्रमी विद्यार्थी, जिन्होंने लेखांकन में कक्षा में प्रथम स्थान पाया"],
 "idp": ["Their ID proof", "t", "Aadhaar card ending 7788", "आधार कार्ड (अंतिम अंक 7788)"],
 "task": ["What they will do for you", "t", "collect my new cheque book and debit card", "मेरी नई चेक बुक और डेबिट कार्ड प्राप्त करने"],
 "deceased": ["Name of the person who passed away", "t", "Shri Mohan Lal ji", "श्री मोहन लाल जी"],
 "drel": ["Their relation to the reader (optional)", "t", "father", "पिताजी"],
 "product": ["Product / service", "t", "a washing machine (model WW70T)", "वॉशिंग मशीन (मॉडल WW70T)"],
 "ask": ["What you want them to do", "t", "replace the machine or refund the full amount", "मशीन बदल दें या पूरी राशि लौटा दें"],
 "policy": ["Policy number", "t", "MOT/2026/554321", "MOT/2026/554321"],
 "help": ["What they helped with", "t", "helping me prepare for my job interview", "नौकरी के साक्षात्कार की तैयारी में मेरी मदद करने"]
},
"TONE": {
 "req": {"formal": ["I shall be thankful for your kind consideration.", "इसके लिए मैं आपका ((आभारी|आभारी)) ((रहूँगा|रहूँगी))।"],
         "polite": ["I would be very grateful if you could kindly consider my request. Thank you for your time and understanding.", "आशा है आप मेरे निवेदन पर सहानुभूतिपूर्वक विचार करेंगे। आपके समय और सहयोग के लिए हार्दिक धन्यवाद।"],
         "firm": ["I request you to act on this at the earliest and let me know the outcome in writing.", "कृपया इस पर शीघ्र कार्रवाई करें और परिणाम की लिखित सूचना दें।"]},
 "comp": {"formal": ["I request you to look into the matter and take the necessary action at the earliest.", "अतः आपसे निवेदन है कि मामले की जाँच कर शीघ्र आवश्यक कार्रवाई करें।"],
          "polite": ["I would be grateful if you could look into this at your earliest convenience. Thank you for your help.", "आशा है आप इस मामले पर शीघ्र ध्यान देंगे। आपके सहयोग के लिए धन्यवाद।"],
          "firm": ["If the matter is not resolved within 15 days of this letter, I will be compelled to take it up with the higher authorities and the appropriate forum without further notice.", "यदि इस पत्र के 15 दिनों के भीतर समस्या का समाधान नहीं हुआ, तो मुझे बिना किसी और सूचना के उच्च अधिकारियों एवं उचित मंच पर शिकायत करने के लिए विवश होना पड़ेगा।"]},
 "info": {"formal": ["Kindly take this on record.", "कृपया इसे अपने रिकॉर्ड में दर्ज करें।"],
          "polite": ["Kindly take this on record. Thank you for your support.", "कृपया इसे अपने रिकॉर्ड में दर्ज करें। आपके सहयोग के लिए धन्यवाद।"],
          "firm": ["Kindly take this on record and acknowledge receipt in writing.", "कृपया इसे रिकॉर्ड में दर्ज कर लिखित पावती दें।"]},
 "remind": {"formal": ["If you have already made the payment, please ignore this letter and share the payment details for our records.", "यदि आप भुगतान कर चुके हैं, तो कृपया यह पत्र अनदेखा करें और हमारे रिकॉर्ड के लिए भुगतान का विवरण भेजें।"],
            "polite": ["We value our relationship with you. If there is any issue with the invoice, please let us know so that we can resolve it quickly.", "हम आपके साथ अपने संबंधों को महत्व देते हैं। यदि इनवॉइस में कोई समस्या हो, तो कृपया बताएँ ताकि हम उसे शीघ्र सुलझा सकें।"],
            "firm": ["Please note that if the payment is not received by this date, we will be forced to stop further supplies, charge interest on the overdue amount as per our terms and take steps to recover it.", "कृपया ध्यान दें कि इस तिथि तक भुगतान न मिलने पर हमें आगे की आपूर्ति रोकनी पड़ेगी, शर्तों के अनुसार बकाया राशि पर ब्याज लगाना पड़ेगा और वसूली के लिए कदम उठाने पड़ेंगे।"]}
},
"T": [
{"id": "sick-leave", "c": "school", "n": "Sick leave application (student)", "h": "बीमारी के कारण छुट्टी हेतु प्रार्थना पत्र", "kw": "medical fever ill absent", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School name]", "[विद्यालय का नाम]"],
 "f": ["d1", "d2", "illness"],
 "s": ["Application for sick leave from {d1} to {d2}", "{d1} से {d2} तक बीमारी के कारण अवकाश हेतु प्रार्थना पत्र"],
 "b": [["With due respect, I beg to state that I am a student of Class {cls} of your school. Since {d1}, I have been suffering from {illness}, and the doctor has advised me complete rest.", "Therefore, I am unable to attend school from {d1} to {d2}[[ ({ndays})]]. Kindly grant me leave for this period."],
       ["सविनय निवेदन है कि मैं आपके विद्यालय की कक्षा {cls} ((का छात्र|की छात्रा)) हूँ। {d1} से मैं {illness} से पीड़ित हूँ और डॉक्टर ने मुझे पूर्ण आराम करने की सलाह दी है।", "अतः मैं {d1} से {d2} तक[[ ({ndays})]] विद्यालय आने में असमर्थ हूँ। कृपया मुझे इस अवधि का अवकाश प्रदान करने की कृपा करें।"]],
 "o": [["cert", "Mention the medical certificate", "I have attached the medical certificate from my doctor.", "डॉक्टर का चिकित्सा प्रमाण पत्र संलग्न है।", 1],
       ["work", "Promise to cover missed work", "I will complete all the classwork and homework that I miss.", "छूटा हुआ सारा कक्षा-कार्य और गृहकार्य मैं पूरा कर ((लूँगा|लूँगी))।", 1]],
 "e": "Medical certificate"},
{"id": "leave-parent", "c": "school", "n": "Leave application by a parent", "h": "अभिभावक द्वारा बच्चे की छुट्टी हेतु प्रार्थना पत्र", "kw": "child son daughter absent mother father guardian", "r": "parent", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School name]", "[विद्यालय का नाम]"],
 "f": ["d1", "d2", "reason"], "sm": {"reason": ["we have to travel to Jaipur for a family wedding", "परिवार में विवाह के लिए जयपुर जाने"]},
 "s": ["Leave application for {child}, Class {cls}", "{child}, कक्षा {cls} के अवकाश हेतु प्रार्थना पत्र"],
 "b": [["With due respect, I am writing to inform you that my ((cg:son|daughter)) {child}, a student of Class {cls}[[ (Roll No. {roll})]], will not be able to attend school from {d1} to {d2}[[ ({ndays})]] because {reason}.", "Kindly grant ((cg:him|her)) leave for these days."],
       ["सविनय निवेदन है कि ((cg:मेरा पुत्र|मेरी पुत्री)) {child}, जो आपके विद्यालय की कक्षा {cls}[[ (अनुक्रमांक {roll})]] ((cg:का छात्र|की छात्रा)) है, {reason} के कारण {d1} से {d2} तक[[ ({ndays})]] विद्यालय नहीं आ ((cg:सकेगा|सकेगी))।", "कृपया उसे इन दिनों का अवकाश प्रदान करें।"]],
 "o": [["catch", "Promise to cover missed work", "I will make sure that ((cg:he|she)) completes the classwork and homework missed during this period.", "इस अवधि का छूटा हुआ कक्षा-कार्य और गृहकार्य पूरा कराने की जिम्मेदारी मेरी होगी।", 1],
       ["call", "Give your phone number", "You may reach me on {sphone} if anything is needed.", "किसी भी जानकारी के लिए मुझसे {sphone} पर संपर्क किया जा सकता है।", 0]]},
{"id": "leave-family", "c": "school", "n": "Leave for a family function or travel (student)", "h": "पारिवारिक समारोह हेतु छुट्टी का प्रार्थना पत्र", "kw": "wedding marriage function out of station urgent work", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School name]", "[विद्यालय का नाम]"],
 "f": ["event", "place", "d1", "d2"],
 "s": ["Application for leave to attend {event}", "{event} में सम्मिलित होने हेतु अवकाश के लिए प्रार्थना पत्र"],
 "b": [["Respectfully, I beg to state that I am a student of Class {cls} of your school. I have to go to {place} with my family to attend {event}.", "As a result, I will not be able to come to school from {d1} to {d2}[[ ({ndays})]]. Kindly grant me leave for these days."],
       ["सविनय निवेदन है कि मैं आपके विद्यालय की कक्षा {cls} ((का छात्र|की छात्रा)) हूँ। मुझे अपने परिवार के साथ {event} में सम्मिलित होने के लिए {place} जाना है।", "इस कारण मैं {d1} से {d2} तक[[ ({ndays})]] विद्यालय नहीं आ ((सकूँगा|सकूँगी))। कृपया मुझे इन दिनों का अवकाश प्रदान करें।"]],
 "o": [["work", "Promise to cover missed work", "I will complete the work missed during my absence as soon as I return.", "लौटने के बाद मैं छूटा हुआ सारा कार्य शीघ्र पूरा कर ((लूँगा|लूँगी))।", 1]]},
{"id": "fee-concession", "c": "school", "n": "Fee concession (fee maafi) application", "h": "शुल्क माफी हेतु प्रार्थना पत्र", "kw": "fees waiver poor financial help maafi", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School name]", "[विद्यालय का नाम]"],
 "f": ["earner", "job", "income", "conc", "marks"],
 "s": ["Request for {conc} fee concession", "{conc} शुल्क माफी हेतु प्रार्थना पत्र"],
 "b": [["With due respect, I am a student of Class {cls} in your school. {earner} works as {job}, and our family's total monthly income is only {income}. With this income it has become very difficult to pay my school fee along with the other needs of the family.", "[[I have always worked hard at my studies and scored {marks} in my last examination. ]]I wish to continue my studies in this school without a break. I therefore request you to kindly grant me {conc} fee concession for the current session."],
       ["सविनय निवेदन है कि मैं आपके विद्यालय की कक्षा {cls} ((का छात्र|की छात्रा)) हूँ। {earner} {job} हैं और हमारे परिवार की कुल मासिक आय केवल {income} है। इतनी आय में परिवार के अन्य खर्चों के साथ मेरी फीस भरना बहुत कठिन हो गया है।", "[[मैं सदा मन लगाकर पढ़ाई ((करता|करती)) हूँ और पिछली परीक्षा में मैंने {marks} अंक प्राप्त किए। ]]मैं इसी विद्यालय में अपनी पढ़ाई जारी रखना ((चाहता|चाहती)) हूँ। अतः आपसे विनम्र निवेदन है कि मुझे इस सत्र के लिए {conc} शुल्क माफी प्रदान करने की कृपा करें।"]],
 "o": [["proof", "Mention the income certificate", "A copy of our income certificate is attached for your reference.", "आय प्रमाण पत्र की प्रति संलग्न है।", 1],
       ["promise", "Promise to keep up good results", "I assure you that I will keep up my good results and conduct.", "मैं आपको विश्वास ((दिलाता|दिलाती)) हूँ कि अपने अच्छे परिणाम और आचरण को बनाए ((रखूँगा|रखूँगी))।", 1]],
 "e": "Income certificate"},
{"id": "tc-request", "c": "school", "n": "Transfer certificate (TC) application", "h": "स्थानांतरण प्रमाण पत्र (टी.सी.) हेतु प्रार्थना पत्र", "kw": "school leaving certificate slc transfer", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School name]", "[विद्यालय का नाम]"],
 "f": ["since", "reason"], "lb": {"since": "Studying here since (year)"}, "sm": {"reason": ["my father has been transferred to Pune", "पिताजी का पुणे तबादला होने"]},
 "s": ["Application for Transfer Certificate (TC)", "स्थानांतरण प्रमाण पत्र (टी.सी.) हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I beg to state that I have been studying in your school since {since} and am at present a student of Class {cls}. As {reason}, I have to leave this school and take admission in another school.", "Kindly issue my Transfer Certificate at the earliest so that I can complete my new admission in time."],
       ["सविनय निवेदन है कि मैं {since} से आपके विद्यालय में अध्ययनरत हूँ और वर्तमान में कक्षा {cls} ((का छात्र|की छात्रा)) हूँ। {reason} के कारण मुझे यह विद्यालय छोड़कर किसी अन्य विद्यालय में प्रवेश लेना है।", "कृपया मुझे शीघ्र स्थानांतरण प्रमाण पत्र प्रदान करें, ताकि मैं समय पर नए विद्यालय में प्रवेश ले सकूँ।"]],
 "o": [["dues", "Say all dues are cleared", "I have paid all my fees and returned the library books and identity card.", "मैंने सभी शुल्क जमा कर दिए हैं और पुस्तकालय की पुस्तकें तथा पहचान पत्र लौटा दिए हैं।", 1],
       ["cc", "Also ask for a character certificate", "I also request you to issue my character certificate along with the TC.", "टी.सी. के साथ मेरा चरित्र प्रमाण पत्र भी जारी करने की कृपा करें।", 1]]},
{"id": "bonafide", "c": "school", "n": "Bonafide certificate request", "h": "बोनाफाइड (वास्तविक छात्र) प्रमाण पत्र हेतु प्रार्थना पत्र", "kw": "bona fide student certificate scholarship bus pass", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School / college name]", "[विद्यालय / महाविद्यालय का नाम]"],
 "f": ["session", "purpose", "by"],
 "s": ["Request for a bonafide certificate", "बोनाफाइड प्रमाण पत्र हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I am a regular student of Class {cls}[[ (Roll No. {roll})]] in your institution for the session {session}. I need a bonafide certificate for {purpose}.", "Kindly issue the certificate at the earliest[[, as it has to be submitted by {by}]]."],
       ["सविनय निवेदन है कि मैं सत्र {session} में आपके संस्थान की कक्षा {cls}[[ (अनुक्रमांक {roll})]] ((का नियमित छात्र|की नियमित छात्रा)) हूँ। मुझे {purpose} के लिए बोनाफाइड प्रमाण पत्र की आवश्यकता है।", "कृपया यह प्रमाण पत्र शीघ्र जारी करने की कृपा करें[[, क्योंकि इसे {by} तक जमा करना है]]।"]],
 "o": [["fee", "Offer to pay any fee", "I am ready to pay the prescribed fee, if any.", "यदि कोई निर्धारित शुल्क है तो मैं उसे जमा करने के लिए तैयार हूँ।", 0]]},
{"id": "char-school", "c": "school", "n": "Character certificate from school / college", "h": "विद्यालय से चरित्र प्रमाण पत्र हेतु प्रार्थना पत्र", "kw": "conduct certificate", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School / college name]", "[विद्यालय / महाविद्यालय का नाम]"],
 "f": ["session", "purpose"], "sm": {"purpose": ["admission to college", "महाविद्यालय में प्रवेश"]},
 "s": ["Request for a character certificate", "चरित्र प्रमाण पत्र हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I studied in Class {cls} of your institution in the session {session}. I need a character certificate for {purpose}.", "During my time here I took part in school activities and was never involved in any act of indiscipline. Kindly issue the certificate at the earliest."],
       ["सविनय निवेदन है कि मैं सत्र {session} में आपके संस्थान की कक्षा {cls} ((का छात्र|की छात्रा)) ((रहा|रही)) हूँ। मुझे {purpose} के लिए चरित्र प्रमाण पत्र की आवश्यकता है।", "अध्ययन के दौरान मैंने विद्यालय की गतिविधियों में भाग लिया और कभी किसी अनुशासनहीनता में शामिल नहीं ((रहा|रही))। कृपया शीघ्र प्रमाण पत्र जारी करने की कृपा करें।"]]},
{"id": "fine-waiver", "c": "school", "n": "Request to waive a late fee fine", "h": "विलंब शुल्क (जुर्माना) माफ करने हेतु प्रार्थना पत्र", "kw": "penalty fine late fee", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School name]", "[विद्यालय का नाम]"],
 "f": ["month", "amt", "reason"], "lb": {"amt": "Fine charged"}, "sm": {"reason": ["my father was in hospital that week", "उस सप्ताह पिताजी के अस्पताल में भर्ती होने"], "amt": ["200", "200"]},
 "s": ["Request to waive the late fee fine for {month}", "{month} के विलंब शुल्क (जुर्माने) को माफ करने हेतु प्रार्थना पत्र"],
 "b": [["With due respect, I am a student of Class {cls}. My fee for {month} was deposited late because {reason}, and a late fine of {amt} has been charged.", "This was the first time it happened and the delay was not intentional. I request you to kindly waive the fine. I will make sure that the fee is paid on time in future."],
       ["सविनय निवेदन है कि मैं कक्षा {cls} ((का छात्र|की छात्रा)) हूँ। {reason} के कारण {month} का शुल्क देर से जमा हुआ और मुझ पर {amt} का विलंब शुल्क लगाया गया है।", "ऐसा पहली बार हुआ है और यह देरी जानबूझकर नहीं की गई। कृपया यह जुर्माना माफ करने की कृपा करें। भविष्य में शुल्क समय पर जमा होगा, इसका मैं पूरा ध्यान ((रखूँगा|रखूँगी))।"]]},
{"id": "section-change", "c": "school", "n": "Change of section or subject", "h": "सेक्शन या विषय बदलने हेतु प्रार्थना पत्र", "kw": "stream subject change section", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School / college name]", "[विद्यालय / महाविद्यालय का नाम]"],
 "f": ["now", "want", "reason"], "sm": {"reason": ["my younger brother is in that section and we travel to school together", "छोटे भाई के उसी सेक्शन में होने और साथ विद्यालय आने-जाने"]},
 "s": ["Request for change from {now} to {want}", "{now} से {want} में परिवर्तन हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I am a student of Class {cls}[[ (Roll No. {roll})]] and am at present in {now}. I request a change to {want} because {reason}.", "I will complete any work or formalities needed for the change. Kindly allow the change from the earliest possible date."],
       ["सविनय निवेदन है कि मैं कक्षा {cls}[[ (अनुक्रमांक {roll})]] ((का छात्र|की छात्रा)) हूँ और अभी {now} में हूँ। {reason} के कारण मैं {want} में परिवर्तन ((चाहता|चाहती)) हूँ।", "परिवर्तन के लिए आवश्यक सभी कार्य और औपचारिकताएँ मैं पूरी ((करूँगा|करूँगी))। कृपया शीघ्रातिशीघ्र यह परिवर्तन करने की अनुमति दें।"]]},
{"id": "duplicate-doc", "c": "school", "n": "Duplicate ID card or mark sheet", "h": "डुप्लीकेट पहचान पत्र / अंक तालिका हेतु प्रार्थना पत्र", "kw": "lost id card marksheet admit card library card", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School / college name]", "[विद्यालय / महाविद्यालय का नाम]"],
 "f": ["doc", "d1"], "lb": {"d1": "Date lost (optional)"},
 "s": ["Request for a duplicate {doc}", "डुप्लीकेट {doc} हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I am a student of Class {cls}[[ (Roll No. {roll})]]. I regret to inform you that I have lost my {doc}[[ on {d1}]]. I have searched everywhere but could not find it.", "Kindly issue me a duplicate {doc}. I am ready to pay the fee charged for it."],
       ["सविनय निवेदन है कि मैं कक्षा {cls}[[ (अनुक्रमांक {roll})]] ((का छात्र|की छात्रा)) हूँ। मुझे खेद है कि[[ {d1} को]] {doc} कहीं खो जाने के कारण अब वह मेरे पास नहीं है। काफी खोजबीन करने पर भी उसका पता नहीं चल सका।", "कृपया मुझे डुप्लीकेट {doc} जारी करने की कृपा करें। इसके लिए निर्धारित शुल्क देने को मैं तैयार हूँ।"]]},
{"id": "scholarship", "c": "school", "n": "Scholarship application", "h": "छात्रवृत्ति हेतु प्रार्थना पत्र", "kw": "stipend merit financial help", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School / college name]", "[विद्यालय / महाविद्यालय का नाम]"],
 "f": ["scheme", "marks", "income"], "lb": {"income": "Family's yearly income (optional)"}, "sm": {"income": ["180000", "180000"]},
 "s": ["Application for the {scheme}", "{scheme} हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I am a student of Class {cls}[[ (Roll No. {roll})]] and wish to apply for the {scheme}. I scored {marks} in my last examination[[, and my family's yearly income is {income}]].", "I am sincere about my studies, and this support will help me continue my education without financial worry. Kindly consider my application and forward it to the authority concerned."],
       ["सविनय निवेदन है कि मैं कक्षा {cls}[[ (अनुक्रमांक {roll})]] ((का छात्र|की छात्रा)) हूँ और {scheme} के लिए आवेदन करना ((चाहता|चाहती)) हूँ। पिछली परीक्षा में मैंने {marks} अंक प्राप्त किए[[ तथा मेरे परिवार की वार्षिक आय {income} है]]।", "मैं अपनी पढ़ाई के प्रति गंभीर हूँ और यह सहायता मुझे बिना आर्थिक चिंता के पढ़ाई जारी रखने में मदद करेगी। कृपया मेरे आवेदन पर विचार कर इसे संबंधित अधिकारी को अग्रेषित करने की कृपा करें।"]],
 "e": "Mark sheet\nIncome certificate"},
{"id": "admission", "c": "school", "n": "Admission request by a parent", "h": "बच्चे के प्रवेश हेतु प्रार्थना पत्र", "kw": "new admission school seat nursery", "r": "parent", "a": ["child", "cg"], "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School name]", "[विद्यालय का नाम]"],
 "f": ["aclass", "prev", "reason"], "lb": {"reason": "Why a new school (optional)"}, "sm": {"reason": ["we have recently moved to this area", "हाल ही में इस क्षेत्र में आने"]},
 "s": ["Application for admission of {child} in Class {aclass}", "{child} के कक्षा {aclass} में प्रवेश हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I wish to seek admission for my ((cg:son|daughter)) {child} in Class {aclass} of your school for the coming session.[[ ((cg:He|She)) is at present studying in {prev}.]]", "[[We are looking for a new school because {reason}. ]]We have heard very good things about the teaching and discipline of your school. Kindly let me know the admission procedure, the documents required and the date of the admission test or interaction, if any."],
       ["सविनय निवेदन है कि मैं आगामी सत्र के लिए ((cg:अपने पुत्र|अपनी पुत्री)) {child} का आपके विद्यालय की कक्षा {aclass} में प्रवेश कराना ((चाहता|चाहती)) हूँ।[[ वर्तमान में वह {prev} में पढ़ ((cg:रहा|रही)) है।]]", "[[{reason} के कारण हम नया विद्यालय ढूँढ रहे हैं। ]]आपके विद्यालय की पढ़ाई और अनुशासन के बारे में हमने बहुत अच्छा सुना है। कृपया प्रवेश की प्रक्रिया, आवश्यक दस्तावेज और प्रवेश परीक्षा अथवा बातचीत की तिथि (यदि कोई हो) बताने की कृपा करें।"]]},
{"id": "migration", "c": "school", "n": "Migration certificate (college / university)", "h": "प्रवासन (माइग्रेशन) प्रमाण पत्र हेतु प्रार्थना पत्र", "kw": "university migration college registrar", "r": "student", "k": "req",
 "to": ["The Registrar", "कुलसचिव महोदय"], "org": ["[University / college name]", "[विश्वविद्यालय / महाविद्यालय का नाम]"],
 "f": ["regno", "session", "newuni"], "sm": {"cls": ["B.Com (final year)", "बी.कॉम (अंतिम वर्ष)"]},
 "s": ["Request for a migration certificate", "प्रवासन (माइग्रेशन) प्रमाण पत्र हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I was a student of {cls} at your institution[[ in the session {session}]], with University Registration No. {regno}. I have been offered admission at {newuni}, which requires a migration certificate.", "Kindly issue my migration certificate at the earliest. I have attached the required documents and the fee receipt."],
       ["सविनय निवेदन है कि मैं[[ सत्र {session} में]] आपके संस्थान में {cls} ((का छात्र|की छात्रा)) ((रहा|रही)) हूँ। मेरा विश्वविद्यालय पंजीकरण क्रमांक {regno} है। मुझे {newuni} में प्रवेश मिला है, जहाँ प्रवासन प्रमाण पत्र आवश्यक है।", "कृपया मेरा प्रवासन प्रमाण पत्र शीघ्र जारी करने की कृपा करें। आवश्यक दस्तावेज और शुल्क की रसीद संलग्न है।"]],
 "e": "Copy of the final mark sheet\nFee receipt"},
{"id": "event-permission", "c": "school", "n": "Permission to take part in an event or competition", "h": "प्रतियोगिता में भाग लेने की अनुमति हेतु प्रार्थना पत्र", "kw": "sports competition tournament quiz debate trip tour", "r": "student", "k": "req",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School / college name]", "[विद्यालय / महाविद्यालय का नाम]"],
 "f": ["event", "place", "d1"], "lb": {"d1": "Event date"}, "sm": {"event": ["the inter-school debate competition", "अंतर-विद्यालय वाद-विवाद प्रतियोगिता"], "place": ["Chandigarh", "चंडीगढ़"]},
 "s": ["Request for permission to take part in {event}", "{event} में भाग लेने की अनुमति हेतु प्रार्थना पत्र"],
 "b": [["Respectfully, I am a student of Class {cls}. I have been selected to take part in {event}, which will be held at {place} on {d1}.", "Kindly allow me to take part and grant me leave from classes for the event. I will make up for any classwork that I miss."],
       ["सविनय निवेदन है कि मैं कक्षा {cls} ((का छात्र|की छात्रा)) हूँ। मेरा चयन {event} में भाग लेने के लिए हुआ है। यह आयोजन {d1} को {place} में होगा।", "कृपया मुझे इसमें भाग लेने की अनुमति तथा उस दिन कक्षाओं से अवकाश प्रदान करें। छूटा हुआ कक्षा-कार्य मैं पूरा कर ((लूँगा|लूँगी))।"]]},
{"id": "apology-school", "c": "school", "n": "Apology letter to the principal or teacher", "h": "क्षमा-याचना पत्र (विद्यालय)", "kw": "sorry late coming misbehaviour mistake", "r": "student", "k": "apol",
 "to": ["The Principal", "प्रधानाचार्य महोदय/महोदया"], "org": ["[School name]", "[विद्यालय का नाम]"],
 "f": ["what", "d1"], "lb": {"d1": "Date (optional)"},
 "s": ["Apology for {what}", "{what} के लिए क्षमा-याचना"],
 "b": [["Respectfully, I am a student of Class {cls}. I sincerely apologise for {what}[[ on {d1}]]. I realise that it was wrong and that it disturbed the discipline of the class.", "I assure you that this will not happen again. Kindly forgive me this once."],
       ["सविनय निवेदन है कि मैं कक्षा {cls} ((का छात्र|की छात्रा)) हूँ।[[ {d1} को]] {what} के लिए मैं हृदय से क्षमा ((माँगता|माँगती)) हूँ। मुझे अपनी गलती का एहसास है कि इससे कक्षा का अनुशासन भंग हुआ।", "मैं आपको विश्वास ((दिलाता|दिलाती)) हूँ कि भविष्य में ऐसा दोबारा नहीं होगा। कृपया मुझे इस बार क्षमा कर दें।"]]},

{"id": "job-application", "c": "office", "n": "Job application (cover letter)", "h": "नौकरी हेतु आवेदन पत्र", "kw": "cover letter vacancy apply resume cv", "r": "person", "k": "req",
 "to": ["The HR Manager", "मानव संसाधन प्रबंधक"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["post", "ad", "qual", "exp", "skills"],
 "s": ["Application for the post of {post}", "{post} पद हेतु आवेदन पत्र"],
 "b": [["I am writing to apply for the post of {post}[[ advertised in {ad}]]. I hold {qual}[[ and have {exp} of experience]].", "[[My key skills include {skills}. ]]I am hardworking, quick to learn and comfortable working in a team, and I am confident that I can contribute to your organisation from the first day.", "I have attached my résumé for your consideration and would welcome the opportunity to meet you for an interview at your convenience."],
       ["[[{ad} में प्रकाशित विज्ञापन के संदर्भ में ]]मैं {post} पद के लिए आवेदन करना ((चाहता|चाहती)) हूँ। मेरी शैक्षणिक योग्यता {qual} है[[ और मुझे {exp} का अनुभव है]]।", "[[मेरे मुख्य कौशल हैं: {skills}। ]]मैं परिश्रमी हूँ, जल्दी ((सीखता|सीखती)) हूँ और टीम में काम करना ((जानता|जानती)) हूँ। मुझे विश्वास है कि मैं पहले दिन से आपकी संस्था के लिए उपयोगी ((सिद्ध होऊँगा|सिद्ध होऊँगी))।", "विचार हेतु मेरा बायोडाटा संलग्न है। आपकी सुविधानुसार साक्षात्कार का अवसर मिलने पर मुझे प्रसन्नता होगी।"]],
 "e": "Résumé\nCopies of certificates"},
{"id": "casual-leave", "c": "office", "n": "Leave application (office)", "h": "कार्यालय में अवकाश हेतु आवेदन", "kw": "casual leave earned leave personal work days off", "r": "employee", "k": "req",
 "to": ["The Manager", "प्रबंधक महोदय"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["d1", "d2", "reason", "cover"], "sm": {"reason": ["I have to attend my cousin's wedding in Jaipur", "चचेरे भाई के विवाह में जयपुर जाने"]},
 "s": ["Leave application from {d1} to {d2}", "{d1} से {d2} तक अवकाश हेतु आवेदन"],
 "b": [["I would like to request leave from {d1} to {d2}[[ ({ndays})]] because {reason}.", "I will complete my urgent work before going on leave[[, and {cover} will take care of my duties in my absence]]."],
       ["{reason} के कारण मैं {d1} से {d2} तक[[ ({ndays})]] अवकाश लेना ((चाहता|चाहती)) हूँ।", "अवकाश पर जाने से पहले मैं अपने आवश्यक कार्य पूरे कर ((लूँगा|लूँगी))[[ और मेरी अनुपस्थिति में {cover} मेरा कार्य देखेंगे]]।"]],
 "o": [["reach", "Say you can be reached", "I will be available on my mobile, {sphone}, for anything urgent.", "आवश्यक कार्य के लिए मैं मोबाइल {sphone} पर उपलब्ध ((रहूँगा|रहूँगी))।", 1]]},
{"id": "sick-leave-office", "c": "office", "n": "Sick leave application (office)", "h": "कार्यालय में बीमारी अवकाश हेतु आवेदन", "kw": "medical leave unwell fever", "r": "employee", "k": "req",
 "to": ["The Manager", "प्रबंधक महोदय"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["illness", "d1", "d2"],
 "s": ["Sick leave application", "बीमारी के कारण अवकाश हेतु आवेदन"],
 "b": [["I am unwell with {illness}, and the doctor has advised me to rest. I am therefore unable to come to the office from {d1} to {d2}[[ ({ndays})]].", "Kindly grant me sick leave for this period. I will resume work as soon as I recover."],
       ["मैं {illness} से पीड़ित हूँ और डॉक्टर ने मुझे आराम करने की सलाह दी है। इसलिए मैं {d1} से {d2} तक[[ ({ndays})]] कार्यालय आने में असमर्थ हूँ।", "कृपया इस अवधि का चिकित्सा अवकाश स्वीकृत करें। स्वस्थ होते ही मैं कार्य पर लौट ((आऊँगा|आऊँगी))।"]],
 "o": [["cert", "Mention the medical certificate", "I will submit the medical certificate when I rejoin.", "कार्य पर लौटने पर मैं चिकित्सा प्रमाण पत्र जमा कर ((दूँगा|दूँगी))।", 1]]},
{"id": "wfh", "c": "office", "n": "Work from home request", "h": "घर से कार्य करने की अनुमति हेतु आवेदन", "kw": "wfh remote work", "r": "employee", "k": "req",
 "to": ["The Manager", "प्रबंधक महोदय"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["d1", "d2", "reason"], "sm": {"reason": ["my mother has had surgery and needs care at home", "माताजी की सर्जरी के बाद घर पर उनकी देखभाल करने"]},
 "s": ["Request to work from home from {d1} to {d2}", "{d1} से {d2} तक घर से कार्य करने की अनुमति हेतु"],
 "b": [["I request your permission to work from home from {d1} to {d2}[[ ({ndays})]] because {reason}.", "I will be online during normal office hours, join all meetings by phone or video call and complete my work on time. My laptop and internet connection are ready for remote work."],
       ["{reason} के कारण मैं {d1} से {d2} तक[[ ({ndays})]] घर से कार्य करने की अनुमति ((चाहता|चाहती)) हूँ।", "मैं कार्यालय के सामान्य समय में ऑनलाइन ((रहूँगा|रहूँगी)), सभी बैठकों में फोन या वीडियो कॉल से ((जुड़ूँगा|जुड़ूँगी)) और अपना कार्य समय पर पूरा ((करूँगा|करूँगी))। घर से कार्य करने के लिए मेरे पास लैपटॉप और इंटरनेट की व्यवस्था है।"]]},
{"id": "resignation", "c": "office", "n": "Resignation letter", "h": "त्यागपत्र", "kw": "resign quit notice period leaving job", "r": "employee", "k": "none",
 "to": ["The Manager", "प्रबंधक महोदय"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["notice", "ldate", "reason"], "lb": {"reason": "Reason (optional)"}, "sm": {"reason": ["I have accepted a role closer to my family in Chandigarh", "चंडीगढ़ में परिवार के पास नई नौकरी मिलने"]},
 "s": ["Resignation from the post of {sdesig}", "{sdesig} पद से त्यागपत्र"],
 "b": [["Please accept this letter as formal notice of my resignation from the post of {sdesig}[[ in the {dept} department]]. As per my {notice} notice period, my last working day will be {ldate}.", "[[I have decided to move on because {reason}. ]]I am grateful for the opportunities and support I have received here, and I have learnt a great deal from you and the team.", "I will complete a proper handover of my work and documents before my last day. Kindly arrange my full and final settlement and relieving letter."],
       ["मैं {sdesig} पद[[ ({dept} विभाग)]] से त्यागपत्र दे ((रहा|रही)) हूँ। मेरी {notice} की सूचना अवधि के अनुसार मेरा अंतिम कार्य दिवस {ldate} होगा।", "[[{reason} के कारण मैंने यह निर्णय लिया है। ]]यहाँ मुझे जो अवसर और सहयोग मिला, उसके लिए मैं आभारी हूँ। आपसे और पूरी टीम से मैंने बहुत कुछ सीखा है।", "अंतिम दिन से पहले मैं अपना सारा कार्य और दस्तावेज विधिवत सौंप ((दूँगा|दूँगी))। कृपया मेरा पूर्ण एवं अंतिम भुगतान तथा कार्यमुक्ति पत्र तैयार करवाने की कृपा करें।"]],
 "t": {"formal": ["Kindly accept my resignation.", "कृपया मेरा त्यागपत्र स्वीकार करें।"], "polite": ["Thank you once again for everything. I wish the company continued success.", "सभी बातों के लिए एक बार फिर धन्यवाद। संस्था की निरंतर प्रगति के लिए शुभकामनाएँ।"], "firm": ["Kindly acknowledge this resignation in writing and confirm my last working day.", "कृपया इस त्यागपत्र की लिखित पावती दें और मेरे अंतिम कार्य दिवस की पुष्टि करें।"]}},
{"id": "relieving-request", "c": "office", "n": "Request for relieving letter and experience certificate", "h": "कार्यमुक्ति पत्र एवं अनुभव प्रमाण पत्र हेतु आवेदन", "kw": "relieving experience letter full and final", "r": "employee", "k": "req",
 "to": ["The HR Manager", "मानव संसाधन प्रबंधक"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["rdate", "ldate"],
 "s": ["Request for relieving letter and experience certificate", "कार्यमुक्ति पत्र एवं अनुभव प्रमाण पत्र हेतु आवेदन"],
 "b": [["I resigned from the post of {sdesig} on {rdate}, and {ldate} was my last working day. I have handed over all my work, files and company property.", "Kindly issue my relieving letter and experience certificate, and let me know the date of my full and final settlement. These documents are needed for my next job."],
       ["मैंने {rdate} को {sdesig} पद से त्यागपत्र दिया था और {ldate} मेरा अंतिम कार्य दिवस था। मैंने अपना सारा कार्य, फाइलें और संस्था की संपत्ति सौंप दी है।", "कृपया मेरा कार्यमुक्ति पत्र और अनुभव प्रमाण पत्र जारी करें तथा पूर्ण एवं अंतिम भुगतान की तिथि बताएँ। अगली नौकरी के लिए इन दस्तावेजों की आवश्यकता है।"]]},
{"id": "salary-cert-request", "c": "office", "n": "Salary certificate request", "h": "वेतन प्रमाण पत्र हेतु आवेदन", "kw": "salary slip proof income loan visa", "r": "employee", "k": "req",
 "to": ["The HR Manager", "मानव संसाधन प्रबंधक"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["purpose", "months"], "lb": {"months": "Also ask for payslips of last … months (optional)"}, "sm": {"purpose": ["a home loan application", "गृह ऋण आवेदन"], "months": ["3", "3"]},
 "s": ["Request for a salary certificate", "वेतन प्रमाण पत्र हेतु आवेदन"],
 "b": [["I need a salary certificate for {purpose}. Kindly issue a certificate showing my designation, date of joining and present monthly salary[[, along with my salary slips for the last {months} months]].", "The papers have to be submitted soon, so I request you to issue it within this week if possible."],
       ["मुझे {purpose} के लिए वेतन प्रमाण पत्र की आवश्यकता है। कृपया मेरे पद, कार्यग्रहण तिथि और वर्तमान मासिक वेतन का प्रमाण पत्र[[ तथा पिछले {months} माह की वेतन पर्चियाँ]] जारी करें।", "दस्तावेज शीघ्र जमा करने हैं, अतः संभव हो तो इसी सप्ताह जारी करने की कृपा करें।"]]},
{"id": "increment", "c": "office", "n": "Salary increment request", "h": "वेतन वृद्धि हेतु आवेदन", "kw": "raise hike appraisal pay rise", "r": "employee", "k": "req",
 "to": ["The Manager", "प्रबंधक महोदय"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["since", "ach", "lastrev"], "lb": {"since": "Working here since"}, "sm": {"since": ["June 2023", "जून 2023"]},
 "s": ["Request for salary increment", "वेतन वृद्धि हेतु आवेदन"],
 "b": [["I have been working as {sdesig} since {since}. During this time I have taken on more responsibility and have consistently met my targets.", "[[Among other things, {ach}. ]]My salary has not been revised[[ since {lastrev}]], while my work and the cost of living have both grown. I request you to kindly review my performance and consider a suitable increment.", "I would be happy to discuss this with you at a convenient time."],
       ["मैं {since} से {sdesig} के पद पर कार्यरत हूँ। इस दौरान मैंने अधिक जिम्मेदारियाँ सँभाली हैं और अपने लक्ष्य लगातार पूरे किए हैं।", "[[अन्य कार्यों के साथ मैंने {ach}। ]]मेरे वेतन में[[ {lastrev} से]] कोई संशोधन नहीं हुआ है, जबकि मेरा कार्य और महँगाई दोनों बढ़े हैं। कृपया मेरे कार्य का मूल्यांकन कर उचित वेतन वृद्धि पर विचार करें।", "सुविधानुसार मैं इस विषय पर आपसे चर्चा करना ((चाहूँगा|चाहूँगी))।"]]},
{"id": "salary-advance", "c": "office", "n": "Salary advance request", "h": "वेतन अग्रिम हेतु आवेदन", "kw": "advance loan from employer", "r": "employee", "k": "req",
 "to": ["The Manager", "प्रबंधक महोदय"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["amt", "reason", "months"], "lb": {"months": "Repay in how many months"}, "sm": {"amt": ["30000", "30000"], "reason": ["I have to pay for my father's medical treatment", "पिताजी के इलाज का खर्च होने"], "months": ["6", "6"]},
 "s": ["Request for a salary advance of {amt}", "{amt} के वेतन अग्रिम हेतु आवेदन"],
 "b": [["I request a salary advance of {amt}[[ ({amtw})]] because {reason}.", "I request that the amount be recovered from my salary in {months} equal monthly instalments, starting next month."],
       ["{reason} के कारण मुझे {amt} के वेतन अग्रिम की आवश्यकता है।", "कृपया यह राशि अगले माह से मेरे वेतन में से {months} समान मासिक किश्तों में काटने की व्यवस्था करें।"]]},
{"id": "transfer-request", "c": "office", "n": "Transfer request", "h": "स्थानांतरण (तबादला) हेतु आवेदन", "kw": "transfer branch location city", "r": "employee", "k": "req",
 "to": ["The Manager", "प्रबंधक महोदय"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["place", "reason"], "lb": {"place": "Branch / city wanted"}, "sm": {"place": ["the Chandigarh branch", "चंडीगढ़ शाखा"], "reason": ["my parents live there and need my support", "माता-पिता के वहाँ रहने और उन्हें मेरी आवश्यकता होने"]},
 "s": ["Request for transfer to {place}", "{place} में स्थानांतरण हेतु आवेदन"],
 "b": [["I am working as {sdesig}[[ in the {dept} department]] at this branch. I request a transfer to {place} because {reason}.", "I will continue to work with the same commitment at the new location and will hand over my present work properly before moving."],
       ["मैं इस शाखा में {sdesig}[[ ({dept} विभाग)]] के पद पर कार्यरत हूँ। {reason} के कारण मैं {place} में स्थानांतरण ((चाहता|चाहती)) हूँ।", "नए स्थान पर भी मैं उसी लगन से कार्य ((करूँगा|करूँगी)) और जाने से पहले वर्तमान कार्य विधिवत सौंप ((दूँगा|दूँगी))।"]]},
{"id": "noc-employer", "c": "office", "n": "NOC request from employer", "h": "नियोक्ता से अनापत्ति प्रमाण पत्र (एनओसी) हेतु आवेदन", "kw": "no objection passport higher studies visa travel", "r": "employee", "k": "req",
 "to": ["The HR Manager", "मानव संसाधन प्रबंधक"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["purpose", "d1", "d2"], "lb": {"d1": "Travel from (optional)", "d2": "Travel to (optional)"}, "sm": {"purpose": ["my passport application", "पासपोर्ट आवेदन"]},
 "s": ["Request for a No Objection Certificate (NOC)", "अनापत्ति प्रमाण पत्र (एनओसी) हेतु आवेदन"],
 "b": [["I need a No Objection Certificate from the company for {purpose}[[ (travel from {d1} to {d2})]].", "This will not affect my work; I will plan my duties and leave in advance. Kindly issue the NOC on the company letterhead at the earliest."],
       ["मुझे {purpose} के लिए संस्था से अनापत्ति प्रमाण पत्र की आवश्यकता है[[ (यात्रा {d1} से {d2} तक)]]।", "इससे मेरे कार्य पर कोई प्रभाव नहीं पड़ेगा; मैं अपने कार्य और अवकाश की योजना पहले से बना ((लूँगा|लूँगी))। कृपया संस्था के लेटरहेड पर शीघ्र अनापत्ति प्रमाण पत्र जारी करें।"]]},
{"id": "joining-report", "c": "office", "n": "Joining report", "h": "कार्यग्रहण रिपोर्ट", "kw": "joining letter first day report duty", "r": "employee", "k": "none",
 "to": ["The HR Manager", "मानव संसाधन प्रबंधक"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["offerdate", "jdate"],
 "s": ["Joining report", "कार्यग्रहण रिपोर्ट"],
 "b": [["With reference to your appointment letter dated {offerdate}, I am pleased to report that I have joined duty as {sdesig}[[ in the {dept} department]] today, {jdate}, in the forenoon.", "I have brought the documents asked for in the letter. Kindly take my joining on record. I look forward to working with you and the team."],
       ["आपके नियुक्ति पत्र दिनांक {offerdate} के संदर्भ में सूचित करना है कि मैंने आज {jdate} को पूर्वाह्न में {sdesig}[[ ({dept} विभाग)]] के पद पर कार्यभार ग्रहण कर लिया है।", "पत्र में माँगे गए दस्तावेज मैं साथ ((लाया|लाई)) हूँ। कृपया मेरी कार्यग्रहण रिपोर्ट रिकॉर्ड में दर्ज करें। आपके और टीम के साथ कार्य करने के लिए मैं उत्सुक हूँ।"]]},
{"id": "maternity-leave", "c": "office", "n": "Maternity leave application", "h": "मातृत्व अवकाश हेतु आवेदन", "kw": "pregnancy baby delivery maternity benefit", "r": "employee", "k": "req",
 "to": ["The HR Manager", "मानव संसाधन प्रबंधक"], "org": ["[Company name]", "[संस्था का नाम]"],
 "f": ["edd", "d1", "d2"],
 "s": ["Application for maternity leave", "मातृत्व अवकाश हेतु आवेदन"],
 "b": [["I am expecting a baby, and my expected date of delivery is {edd}. On my doctor's advice, I request maternity leave from {d1} to {d2}[[ ({ndays})]], as allowed under the company's leave policy and the Maternity Benefit Act.", "I will hand over my ongoing work before I go on leave and keep my reporting manager informed of any change in dates."],
       ["मैं गर्भवती हूँ और प्रसव की संभावित तिथि {edd} है। डॉक्टर की सलाह पर, संस्था की अवकाश नीति और मातृत्व लाभ अधिनियम के अनुसार, मैं {d1} से {d2} तक[[ ({ndays})]] मातृत्व अवकाश हेतु निवेदन करती हूँ।", "अवकाश पर जाने से पहले मैं अपना चालू कार्य सौंप दूँगी और तिथियों में किसी भी बदलाव की सूचना अपने प्रबंधक को देती रहूँगी।"]],
 "e": "Medical certificate from the doctor"},

{"id": "appointment", "c": "hr", "n": "Appointment letter (employer)", "h": "नियुक्ति पत्र (नियोक्ता द्वारा)", "kw": "offer letter job letter hire new employee", "r": "org", "k": "none",
 "to": ["{ename}", "{ename}"], "org": ["", ""], "sal": ["Dear {ename}", "प्रिय {ename}"],
 "f": ["ename", "eg", "post", "edept", "jdate", "place", "salary", "prob", "hours"], "lb": {"place": "Place of work (optional)", "edept": "Department (optional)"}, "sm": {"place": ["our Ludhiana office", "हमारा लुधियाना कार्यालय"], "edept": ["Sales", "विक्रय"], "post": ["Sales Executive", "विक्रय अधिकारी"]},
 "s": ["Appointment as {post}", "{post} पद पर नियुक्ति"],
 "b": [["We are pleased to appoint you as {post}[[ in the {edept} department]] of {sorg}. Your appointment is effective from {jdate}[[, and your place of work will be {place}]].", "Your gross salary will be {salary} per month[[ ({amtw})]]. Statutory deductions such as PF, ESI, professional tax and income tax will be made as per law. You will be on probation for {prob} months, after which your confirmation will be communicated in writing.", "[[Your working hours will be {hours}. ]]During your employment you will follow the rules and policies of the company and keep all business information confidential. Either side may end this appointment by giving one month's notice in writing or one month's salary in lieu of notice.", "Please sign the duplicate copy of this letter as your acceptance. We welcome you to {sorg} and wish you a rewarding career with us."],
       ["हमें आपको {sorg} में {post}[[ ({edept} विभाग)]] के पद पर नियुक्त करते हुए प्रसन्नता हो रही है। आपकी नियुक्ति {jdate} से प्रभावी होगी[[ और आपका कार्यस्थल {place} होगा]]।", "आपका सकल मासिक वेतन {salary} होगा। पीएफ, ईएसआई, व्यवसाय कर और आयकर जैसी वैधानिक कटौतियाँ नियमानुसार की जाएँगी। आप {prob} माह की परिवीक्षा पर ((eg:रहेंगे|रहेंगी)), जिसके बाद आपकी पुष्टि की सूचना लिखित में दी जाएगी।", "[[आपके कार्य का समय {hours} होगा। ]]सेवा के दौरान आपको संस्था के नियमों एवं नीतियों का पालन करना होगा तथा व्यावसायिक जानकारी गोपनीय रखनी होगी। कोई भी पक्ष एक माह की लिखित सूचना या एक माह के वेतन के बदले यह नियुक्ति समाप्त कर सकता है।", "स्वीकृति के रूप में कृपया इस पत्र की दूसरी प्रति पर हस्ताक्षर करके लौटाएँ। {sorg} में आपका स्वागत है।"]],
 "o": [["docs", "List documents to bring on joining", "Please bring your identity and address proof, educational certificates, the relieving letter from your last employer and two photographs on the day of joining.", "कार्यग्रहण के दिन कृपया पहचान एवं पते का प्रमाण, शैक्षणिक प्रमाण पत्र, पिछले नियोक्ता का कार्यमुक्ति पत्र और दो फोटो साथ लाएँ।", 1]], "oa": 3},
{"id": "warning", "c": "hr", "n": "Warning letter to an employee", "h": "कर्मचारी को चेतावनी पत्र", "kw": "misconduct late coming discipline show cause", "r": "org", "k": "none",
 "to": ["{ename}", "{ename}"], "org": ["{edesig}", "{edesig}"], "sal": ["Dear {ename}", "प्रिय {ename}"],
 "f": ["ename", "eg", "edesig", "issue", "d1"], "lb": {"issue": "What the employee did", "d1": "Date of the incident (optional)"}, "sm": {"issue": ["you have been reaching the office late on most days this month without informing your manager", "आप इस माह अधिकांश दिनों में बिना सूचना दिए देर से कार्यालय पहुँचे हैं"]},
 "s": ["Warning letter", "चेतावनी पत्र"],
 "b": [["It has come to the management's notice that[[ on {d1}]] {issue}. This is against the rules of the company and the standard of conduct expected from every employee.", "You are hereby warned to correct this immediately. Any repetition will lead to strict disciplinary action as per company rules, which may include termination of service.", "Please treat this as a final opportunity to improve, and sign the copy of this letter as acknowledgement."],
       ["प्रबंधन के संज्ञान में आया है कि[[ {d1} को]] {issue}। यह संस्था के नियमों और प्रत्येक कर्मचारी से अपेक्षित आचरण के विरुद्ध है।", "आपको चेतावनी दी जाती है कि इसमें तुरंत सुधार करें। ऐसा दोबारा होने पर संस्था के नियमों के अनुसार कड़ी अनुशासनात्मक कार्रवाई की जाएगी, जिसमें सेवा समाप्ति भी शामिल हो सकती है।", "कृपया इसे सुधार का अंतिम अवसर समझें और पावती के रूप में इस पत्र की प्रति पर हस्ताक्षर करें।"]],
 "o": [["earlier", "Say it was discussed before", "This matter was discussed with you earlier, but there has been no improvement.", "इस विषय पर आपसे पहले भी बात की गई थी, परंतु कोई सुधार नहीं हुआ।", 0],
       ["reply", "Ask for a written explanation", "Please give your written explanation within three days of receiving this letter.", "इस पत्र की प्राप्ति के तीन दिनों के भीतर अपना लिखित स्पष्टीकरण दें।", 1]], "oa": 1},
{"id": "experience-cert", "c": "hr", "n": "Experience certificate (employer)", "h": "अनुभव प्रमाण पत्र (नियोक्ता द्वारा)", "kw": "service certificate experience letter to whom it may concern", "r": "org", "k": "none", "tw": 1,
 "f": ["ename", "eg", "edesig", "edept", "jdate", "ldate", "duties"], "lb": {"edept": "Department (optional)"},
 "s": ["Experience certificate", "अनुभव प्रमाण पत्र"],
 "b": [["This is to certify that ((eg:Mr.|Ms.)) {ename} worked with {sorg} from {jdate} to {ldate}. At the time of leaving, ((eg:he|she)) held the position of {edesig}[[ in the {edept} department]].", "[[((eg:His|Her)) main responsibilities included {duties}. ]]During ((eg:his|her)) service we found ((eg:him|her)) sincere, hardworking and honest, and ((eg:his|her)) conduct was good throughout.", "((eg:He|She)) has been relieved of ((eg:his|her)) duties after completing the handover. We wish ((eg:him|her)) every success in the future."],
       ["प्रमाणित किया जाता है कि ((eg:श्री|सुश्री)) {ename} ने {jdate} से {ldate} तक {sorg} में कार्य किया। संस्था छोड़ते समय वे {edesig}[[ ({edept} विभाग)]] के पद पर ((eg:कार्यरत थे|कार्यरत थीं))।", "[[उनके मुख्य दायित्व थे: {duties}। ]]सेवा के दौरान हमने उन्हें ईमानदार, परिश्रमी और निष्ठावान पाया। उनका आचरण सदैव अच्छा रहा।", "कार्य सौंपने के बाद उन्हें कार्यमुक्त कर दिया गया है। हम उनके उज्ज्वल भविष्य की कामना करते हैं।"]]},
{"id": "relieving", "c": "hr", "n": "Relieving letter (employer)", "h": "कार्यमुक्ति पत्र (नियोक्ता द्वारा)", "kw": "release letter resignation accepted", "r": "org", "k": "none",
 "to": ["{ename}", "{ename}"], "org": ["{edesig}", "{edesig}"], "sal": ["Dear {ename}", "प्रिय {ename}"],
 "f": ["ename", "eg", "ecode", "edesig", "rdate", "ldate"],
 "s": ["Relieving letter", "कार्यमुक्ति पत्र"],
 "b": [["This refers to your resignation letter dated {rdate}. Your resignation has been accepted, and you are relieved from the services of {sorg} at the close of working hours on {ldate}.", "We confirm that you worked with us as {edesig}[[ (Employee ID {ecode})]] and have handed over your charge and company property. Your full and final settlement will be processed as per company policy.", "We thank you for your contribution and wish you success in your future career."],
       ["यह आपके दिनांक {rdate} के त्यागपत्र के संदर्भ में है। आपका त्यागपत्र स्वीकार कर लिया गया है और आपको {ldate} को कार्य समय की समाप्ति पर {sorg} की सेवाओं से कार्यमुक्त किया जाता है।", "हम पुष्टि करते हैं कि आपने हमारे साथ {edesig}[[ (कर्मचारी कोड {ecode})]] के रूप में कार्य किया और अपना कार्यभार तथा संस्था की संपत्ति सौंप दी है। आपका पूर्ण एवं अंतिम भुगतान संस्था की नीति के अनुसार किया जाएगा।", "आपके योगदान के लिए धन्यवाद। आपके भावी करियर के लिए शुभकामनाएँ।"]]},
{"id": "internship-cert", "c": "hr", "n": "Internship certificate", "h": "इंटर्नशिप प्रमाण पत्र", "kw": "intern trainee training completion", "r": "org", "k": "none", "tw": 1,
 "f": ["ename", "eg", "edept", "d1", "d2", "project"], "lb": {"ename": "Intern's name", "eg": "Intern is", "edept": "Department (optional)"},
 "s": ["Internship certificate", "इंटर्नशिप प्रमाण पत्र"],
 "b": [["This is to certify that ((eg:Mr.|Ms.)) {ename} completed an internship with {sorg}[[ in the {edept} department]] from {d1} to {d2}.", "[[During the internship ((eg:he|she)) worked on {project}. ]]((eg:He|She)) was punctual, eager to learn and worked well with the team.", "We wish ((eg:him|her)) all the best for ((eg:his|her)) studies and career."],
       ["प्रमाणित किया जाता है कि ((eg:श्री|सुश्री)) {ename} ने {d1} से {d2} तक {sorg}[[ के {edept} विभाग]] में इंटर्नशिप पूरी की।", "[[इंटर्नशिप के दौरान उन्होंने {project} पर कार्य किया। ]]वे समय के पाबंद रहे, सीखने के लिए उत्सुक रहे और टीम के साथ उनका तालमेल अच्छा रहा।", "हम उनकी पढ़ाई और करियर के लिए शुभकामनाएँ देते हैं।"]]},
{"id": "salary-cert", "c": "hr", "n": "Salary certificate (employer)", "h": "वेतन प्रमाण पत्र (नियोक्ता द्वारा)", "kw": "income certificate employee salary proof", "r": "org", "k": "none", "tw": 1,
 "f": ["ename", "eg", "edesig", "jdate", "salary", "purpose"], "sm": {"purpose": ["a home loan application", "गृह ऋण आवेदन"]},
 "s": ["Salary certificate", "वेतन प्रमाण पत्र"],
 "b": [["This is to certify that ((eg:Mr.|Ms.)) {ename} has been working with {sorg} as {edesig} since {jdate}.", "((eg:His|Her)) present gross salary is {salary} per month[[ ({amtw})]].", "This certificate is issued at ((eg:his|her)) request for {purpose} and carries no financial liability on the part of the company."],
       ["प्रमाणित किया जाता है कि ((eg:श्री|सुश्री)) {ename} {jdate} से {sorg} में {edesig} के पद पर कार्यरत हैं।", "उनका वर्तमान सकल मासिक वेतन {salary} है।", "यह प्रमाण पत्र उनके अनुरोध पर {purpose} हेतु जारी किया जाता है और इससे संस्था का कोई वित्तीय दायित्व नहीं बनता।"]]},
{"id": "increment-letter", "c": "hr", "n": "Salary revision letter (employer)", "h": "वेतन संशोधन पत्र (नियोक्ता द्वारा)", "kw": "increment letter raise hike revised salary", "r": "org", "k": "none",
 "to": ["{ename}", "{ename}"], "org": ["", ""], "sal": ["Dear {ename}", "प्रिय {ename}"],
 "f": ["ename", "eg", "salary", "newsal", "d1"], "lb": {"salary": "Present monthly gross salary", "d1": "Effective from"},
 "s": ["Revision of salary", "वेतन संशोधन"],
 "b": [["We are pleased to inform you that, in recognition of your performance, your gross monthly salary has been revised from {salary} to {newsal} with effect from {d1}.", "All other terms of your employment remain the same. We appreciate your hard work and look forward to your continued contribution."],
       ["हमें यह सूचित करते हुए प्रसन्नता है कि आपके कार्य प्रदर्शन को देखते हुए {d1} से आपका सकल मासिक वेतन {salary} से बढ़ाकर {newsal} कर दिया गया है।", "आपकी सेवा की अन्य सभी शर्तें यथावत रहेंगी। हम आपके परिश्रम की सराहना करते हैं और आगे भी आपके योगदान की आशा करते हैं।"]]},

{"id": "account-closure", "c": "bank", "n": "Bank account closure", "h": "बैंक खाता बंद करने हेतु आवेदन", "kw": "close account bank", "r": "person", "k": "req", "x": ["saddr", "sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "actype", "reason", "toacc"], "lb": {"reason": "Reason (optional)"}, "sm": {"reason": ["I am moving to another city", "मैं दूसरे शहर में जा रहा हूँ"]},
 "s": ["Request to close my {actype} account No. {acno}", "{actype} खाता संख्या {acno} बंद करने हेतु आवेदन"],
 "b": [["I hold a {actype} account, No. {acno}, at your branch. I wish to close this account[[ as {reason}]].", "Kindly close the account and pay me the remaining balance[[ by transfer to my account No. {toacc}]]. I am returning the unused cheque leaves and my debit card with this letter."],
       ["आपकी शाखा में मेरा {actype} खाता, संख्या {acno}, है। मैं यह खाता बंद कराना ((चाहता|चाहती)) हूँ[[, क्योंकि {reason}]]।", "कृपया खाता बंद कर शेष राशि का भुगतान करें[[ (मेरे खाता संख्या {toacc} में स्थानांतरण द्वारा)]]। अप्रयुक्त चेक और डेबिट कार्ड इस पत्र के साथ लौटा ((रहा|रही)) हूँ।"]],
 "o": [["kyc", "Mention the ID proof copy", "A self-attested copy of my identity proof is attached.", "मेरे पहचान प्रमाण की स्व-सत्यापित प्रति संलग्न है।", 1]],
 "e": "Unused cheque leaves\nDebit card\nCopy of ID proof"},
{"id": "bank-address-change", "c": "bank", "n": "Change of address in bank account", "h": "बैंक खाते में पता बदलने हेतु आवेदन", "kw": "address update kyc bank new address", "r": "person", "k": "req", "x": ["sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "newaddr"],
 "s": ["Change of address in account No. {acno}", "खाता संख्या {acno} में पता परिवर्तन हेतु आवेदन"],
 "b": [["I hold account No. {acno} at your branch. I have moved house, and my new address is:\n{newaddr}", "Kindly update this address in your records for all letters, statements and cards. A self-attested copy of my new address proof is attached."],
       ["आपकी शाखा में मेरा खाता संख्या {acno} है। मैं नए मकान में ((आ गया|आ गई)) हूँ। मेरा नया पता है:\n{newaddr}", "कृपया सभी पत्रों, विवरणों और कार्ड के लिए यह पता अपने रिकॉर्ड में अद्यतन करें। नए पते के प्रमाण की स्व-सत्यापित प्रति संलग्न है।"]],
 "e": "Address proof (self-attested)"},
{"id": "cheque-book", "c": "bank", "n": "New cheque book request", "h": "नई चेक बुक हेतु आवेदन", "kw": "cheque book chequebook", "r": "person", "k": "req", "x": ["saddr", "sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "actype", "leaves"],
 "s": ["Request for a new cheque book", "नई चेक बुक हेतु आवेदन"],
 "b": [["I hold a {actype} account, No. {acno}, at your branch. My present cheque book is almost used up. Kindly issue a new cheque book of {leaves} leaves.", "Please send it to my registered address, or let me know when I can collect it from the branch."],
       ["आपकी शाखा में मेरा {actype} खाता, संख्या {acno}, है। मेरी वर्तमान चेक बुक लगभग समाप्त हो गई है। कृपया {leaves} पन्नों की नई चेक बुक जारी करें।", "कृपया इसे मेरे पंजीकृत पते पर भेजें, या बताएँ कि मैं इसे शाखा से कब ले ((सकता|सकती)) हूँ।"]]},
{"id": "bank-statement", "c": "bank", "n": "Bank statement request", "h": "बैंक खाता विवरण (स्टेटमेंट) हेतु आवेदन", "kw": "statement passbook account details", "r": "person", "k": "req", "x": ["saddr", "sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "d1", "d2", "purpose"], "lb": {"purpose": "Needed for (optional)"}, "sm": {"purpose": ["my visa application", "वीज़ा आवेदन"]},
 "s": ["Request for account statement from {d1} to {d2}", "{d1} से {d2} तक खाता विवरण हेतु आवेदन"],
 "b": [["I hold account No. {acno} at your branch. I need a statement of this account for the period {d1} to {d2}[[ for {purpose}]].", "Kindly issue the statement duly signed and stamped by the branch. You may debit any applicable charges from my account."],
       ["आपकी शाखा में मेरा खाता संख्या {acno} है। मुझे[[ {purpose} के लिए]] {d1} से {d2} तक की अवधि का खाता विवरण चाहिए।", "कृपया शाखा द्वारा हस्ताक्षरित एवं मुहरयुक्त विवरण जारी करें। लागू शुल्क मेरे खाते से काटा जा सकता है।"]]},
{"id": "loan-application", "c": "bank", "n": "Loan application", "h": "ऋण (लोन) हेतु आवेदन", "kw": "personal loan home loan education loan vehicle business", "r": "person", "k": "req", "x": ["saddr", "sphone", "semail"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["loantype", "amt", "months", "purpose", "work", "income", "acno"], "lb": {"months": "Repay in how many months", "purpose": "What the loan is for", "acno": "Your account number here (optional)", "income": "Monthly income (optional)"}, "sm": {"amt": ["300000", "300000"], "months": ["36", "36"], "purpose": ["my daughter's college fees", "बेटी की कॉलेज फीस"], "income": ["65000", "65000"]},
 "s": ["Application for a {loantype} loan of {amt}", "{amt} के {loantype} ऋण हेतु आवेदन"],
 "b": [["I wish to apply for a {loantype} loan of {amt}[[ ({amtw})]] for {purpose}, to be repaid in {months} monthly instalments.", "I work as {work}[[ with a monthly income of {income}]][[, and I hold account No. {acno} at your branch]]. My income is regular and I have a good repayment record.", "Kindly let me know the interest rate, processing charges and documents required, and process my application at the earliest."],
       ["मैं {purpose} के लिए {amt} का {loantype} ऋण लेना ((चाहता|चाहती)) हूँ, जिसे {months} मासिक किश्तों में चुकाया जाएगा।", "मैं {work} हूँ[[ और मेरी मासिक आय {income} है]][[। आपकी शाखा में मेरा खाता संख्या {acno} है]]। मेरी आय नियमित है और मेरा भुगतान रिकॉर्ड अच्छा है।", "कृपया ब्याज दर, प्रोसेसिंग शुल्क और आवश्यक दस्तावेजों की जानकारी दें तथा मेरे आवेदन पर शीघ्र कार्रवाई करें।"]],
 "e": "Identity and address proof\nSalary slips / income proof for the last 3 months\nBank statement for the last 6 months"},
{"id": "card-block", "c": "bank", "n": "Block lost ATM / debit card", "h": "खोया हुआ एटीएम / डेबिट कार्ड ब्लॉक करने हेतु आवेदन", "kw": "atm card lost stolen block hotlist", "r": "person", "k": "req", "x": ["saddr", "sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "card", "d1", "place"], "lb": {"d1": "Date lost", "place": "Where (optional)"}, "sm": {"place": ["the railway station", "रेलवे स्टेशन"]},
 "s": ["Request to block my debit card ending {card}", "अंतिम अंक {card} वाले डेबिट कार्ड को ब्लॉक करने हेतु आवेदन"],
 "b": [["I hold account No. {acno} at your branch. My debit card ending in {card} was lost on {d1}[[ at {place}]].", "Kindly block the card immediately so that it cannot be misused, and issue a new card to my registered address. I will not be responsible for any transaction made with this card after it was lost."],
       ["आपकी शाखा में मेरा खाता संख्या {acno} है। {d1} को[[ {place} पर]] मेरा डेबिट कार्ड, जिसके अंतिम अंक {card} हैं, खो गया।", "कृपया दुरुपयोग रोकने के लिए कार्ड को तुरंत ब्लॉक करें और मेरे पंजीकृत पते पर नया कार्ड जारी करें। कार्ड खोने के बाद इससे हुए किसी भी लेन-देन के लिए मैं उत्तरदायी नहीं ((रहूँगा|रहूँगी))।"]],
 "o": [["called", "Say you already called customer care", "I reported the loss to customer care on the same day.", "मैंने उसी दिन ग्राहक सेवा पर इसकी सूचना दे दी थी।", 1]]},
{"id": "stop-payment", "c": "bank", "n": "Stop payment of a cheque", "h": "चेक का भुगतान रोकने हेतु आवेदन", "kw": "cheque stop lost cheque", "r": "person", "k": "req", "x": ["saddr", "sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "chq", "chqdate", "amt", "payee", "reason"], "lb": {"reason": "Reason (optional)"}, "sm": {"amt": ["18500", "18500"], "reason": ["the cheque has been lost in transit", "चेक के रास्ते में खो जाने"]},
 "s": ["Stop payment of cheque No. {chq}", "चेक संख्या {chq} का भुगतान रोकने हेतु आवेदन"],
 "b": [["I hold account No. {acno} at your branch. I request you to stop payment of the following cheque:\nCheque No.: {chq}\nDate: {chqdate}\nAmount: {amt}\nIn favour of: {payee}", "[[This is needed because {reason}. ]]Please do not honour the cheque if it is presented. You may debit the stop-payment charges from my account."],
       ["आपकी शाखा में मेरा खाता संख्या {acno} है। कृपया निम्नलिखित चेक का भुगतान रोकें:\nचेक संख्या: {chq}\nदिनांक: {chqdate}\nराशि: {amt}\nप्राप्तकर्ता: {payee}", "[[{reason} के कारण यह आवश्यक है। ]]यदि यह चेक प्रस्तुत हो तो इसका भुगतान न करें। भुगतान रोकने का शुल्क मेरे खाते से काटा जा सकता है।"]]},
{"id": "mobile-update", "c": "bank", "n": "Change of registered mobile number", "h": "बैंक में मोबाइल नंबर बदलने हेतु आवेदन", "kw": "mobile number update otp sms alerts", "r": "person", "k": "req", "x": ["saddr"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "oldmob", "newmob"],
 "s": ["Change of registered mobile number in account No. {acno}", "खाता संख्या {acno} में पंजीकृत मोबाइल नंबर बदलने हेतु आवेदन"],
 "b": [["I hold account No. {acno} at your branch.[[ The mobile number now registered with it, {oldmob}, is no longer in use.]]", "Kindly register my new mobile number, {newmob}, for SMS alerts, OTPs and internet and mobile banking. A copy of my identity proof is attached."],
       ["आपकी शाखा में मेरा खाता संख्या {acno} है।[[ इसमें पंजीकृत मोबाइल नंबर {oldmob} अब उपयोग में नहीं है।]]", "कृपया एसएमएस अलर्ट, ओटीपी तथा इंटरनेट और मोबाइल बैंकिंग के लिए मेरा नया मोबाइल नंबर {newmob} दर्ज करें। मेरे पहचान प्रमाण की प्रति संलग्न है।"]],
 "e": "Copy of ID proof"},
{"id": "nominee", "c": "bank", "n": "Add or change nominee in bank account", "h": "बैंक खाते में नॉमिनी जोड़ने / बदलने हेतु आवेदन", "kw": "nomination nominee", "r": "person", "k": "req", "x": ["saddr", "sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "nominee", "nrel", "nage"],
 "s": ["Registration of nominee in account No. {acno}", "खाता संख्या {acno} में नॉमिनी दर्ज करने हेतु आवेदन"],
 "b": [["I hold account No. {acno} at your branch. I wish to register or change the nominee for this account as follows:\nName: {nominee}\nRelationship: {nrel}[[\nAge: {nage}]]", "Kindly update the nomination in your records and give me an acknowledgement. I will fill in the bank's nomination form if required."],
       ["आपकी शाखा में मेरा खाता संख्या {acno} है। मैं इस खाते में निम्नानुसार नॉमिनी दर्ज / परिवर्तित कराना ((चाहता|चाहती)) हूँ:\nनाम: {nominee}\nसंबंध: {nrel}[[\nआयु: {nage}]]", "कृपया अपने रिकॉर्ड में नामांकन अद्यतन कर मुझे पावती दें। आवश्यकता होने पर मैं बैंक का नामांकन फॉर्म भर ((दूँगा|दूँगी))।"]]},
{"id": "failed-txn", "c": "bank", "n": "Complaint about money debited in a failed transaction", "h": "असफल लेन-देन में कटी राशि की शिकायत", "kw": "atm upi failed transaction refund amount debited not received", "r": "person", "k": "comp", "x": ["saddr", "sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["acno", "mode", "d1", "amt", "txnref"], "lb": {"d1": "Transaction date"}, "sm": {"amt": ["5000", "5000"]},
 "s": ["Complaint about failed transaction of {amt} on {d1}", "{d1} को {amt} के असफल लेन-देन की शिकायत"],
 "b": [["I hold account No. {acno} at your branch. On {d1}, I attempted {mode} of {amt}[[ (reference No. {txnref})]]. The transaction failed, but the amount was debited from my account and has not been returned so far.", "Kindly look into this and credit the amount back to my account. Under the RBI rules on failed transactions, such amounts are to be reversed within a fixed time, failing which compensation is payable to the customer."],
       ["आपकी शाखा में मेरा खाता संख्या {acno} है। {d1} को मैंने {amt} का {mode} करने का प्रयास किया[[ (संदर्भ संख्या {txnref})]]। लेन-देन असफल रहा, परंतु राशि मेरे खाते से कट गई और अब तक वापस नहीं आई है।", "कृपया इसकी जाँच कर राशि मेरे खाते में वापस जमा करें। असफल लेन-देन पर आरबीआई के नियमों के अनुसार ऐसी राशि निर्धारित समय में लौटानी होती है, अन्यथा ग्राहक को क्षतिपूर्ति देय होती है।"]],
 "t": {"firm": ["If the matter is not resolved within 30 days, I will be compelled to file a complaint with the RBI Ombudsman.", "यदि 30 दिनों में समस्या का समाधान नहीं हुआ, तो मुझे आरबीआई लोकपाल के पास शिकायत दर्ज करने के लिए विवश होना पड़ेगा।"]}},
{"id": "loan-noc", "c": "bank", "n": "NOC and documents after loan closure", "h": "ऋण बंद होने पर एनओसी और दस्तावेज हेतु आवेदन", "kw": "loan closed no dues certificate original documents hypothecation", "r": "person", "k": "req", "x": ["saddr", "sphone"],
 "to": ["The Branch Manager", "शाखा प्रबंधक"], "org": ["[Bank name]", "[बैंक का नाम]"],
 "f": ["loanno", "d1"], "lb": {"d1": "Loan closed on"},
 "s": ["Request for NOC and documents after closure of loan account No. {loanno}", "ऋण खाता संख्या {loanno} बंद होने पर अनापत्ति प्रमाण पत्र एवं दस्तावेज हेतु आवेदन"],
 "b": [["I had taken a loan from your branch under loan account No. {loanno}. I have repaid the entire loan, and the account was closed on {d1}.", "Kindly issue the No Objection Certificate (NOC) and return all the original documents held by the bank. Please also remove any lien or charge on the security and report the closure to the credit bureaus."],
       ["मैंने आपकी शाखा से ऋण खाता संख्या {loanno} के अंतर्गत ऋण लिया था। मैंने पूरा ऋण चुका दिया है और यह खाता {d1} को बंद हो गया है।", "कृपया अनापत्ति प्रमाण पत्र (एनओसी) जारी करें और बैंक के पास रखे सभी मूल दस्तावेज लौटाएँ। कृपया जमानत पर दर्ज ग्रहणाधिकार (लियन) भी हटाएँ और क्रेडिट ब्यूरो को खाता बंद होने की सूचना दें।"]],
 "o": [["hyp", "Vehicle loan: ask for Form 35", "As this was a vehicle loan, please also issue Form 35 for removing the hypothecation from the registration certificate.", "यह वाहन ऋण था, अतः पंजीकरण प्रमाण पत्र से दृष्टिबंधक (हाइपोथिकेशन) हटाने के लिए फॉर्म 35 भी जारी करें।", 0]]},

{"id": "sanitation-complaint", "c": "govt", "n": "Complaint about garbage and blocked drains", "h": "कूड़ा न उठने और नालियाँ जाम होने की शिकायत", "kw": "municipal corporation sanitation garbage cleanliness drain sewer nagar nigam", "r": "person", "k": "comp",
 "to": ["The Municipal Commissioner", "नगर आयुक्त महोदय"], "org": ["[Municipal Corporation / Council]", "[नगर निगम / नगर परिषद]"], "so": ["Municipal Corporation", "नगर निगम"],
 "f": ["area", "since", "details"], "lb": {"since": "For how long (optional)"}, "sm": {"since": ["two weeks", "दो सप्ताह"], "details": ["The garbage point near the park has not been cleared for days.", "पार्क के पास का कूड़ाघर कई दिनों से साफ नहीं हुआ है।"]},
 "s": ["Complaint about garbage and blocked drains in {area}", "{area} में कूड़ा न उठने और नालियाँ जाम होने की शिकायत"],
 "b": [["I am a resident of {area}, and I wish to bring to your notice the poor state of sanitation in our locality. Garbage is not being collected regularly[[ for the last {since}]], and the drains are blocked.", "[[{details} ]]The heaps of waste and dirty water are breeding mosquitoes and flies, and there is a real risk of disease, especially for children and elderly people.", "I request you to arrange regular garbage collection, cleaning of the drains and spraying of insecticide in the area."],
       ["मैं {area} ((का निवासी|की निवासी)) हूँ और आपका ध्यान हमारे क्षेत्र की खराब सफाई व्यवस्था की ओर दिलाना ((चाहता|चाहती)) हूँ।[[ पिछले {since} से]] कूड़ा नियमित रूप से नहीं उठाया जा रहा है और नालियाँ जाम हैं।", "[[{details} ]]कूड़े के ढेर और गंदे पानी के कारण मच्छर-मक्खियाँ पनप रहे हैं और बीमारियों का खतरा बढ़ गया है, विशेषकर बच्चों और बुजुर्गों के लिए।", "आपसे निवेदन है कि क्षेत्र में नियमित कूड़ा उठाने, नालियों की सफाई और कीटनाशक छिड़काव की व्यवस्था करवाएँ।"]]},
{"id": "road-light", "c": "govt", "n": "Road repair or street light complaint", "h": "सड़क मरम्मत / स्ट्रीट लाइट हेतु शिकायत", "kw": "pothole road street light municipal pwd", "r": "person", "k": "comp",
 "to": ["The Executive Engineer", "अधिशासी अभियंता महोदय"], "org": ["[Municipal Corporation / PWD]", "[नगर निगम / लोक निर्माण विभाग]"], "so": ["Municipal Corporation", "नगर निगम"],
 "f": ["area", "rproblem", "since"], "lb": {"since": "For how long (optional)"}, "sm": {"since": ["over a month", "एक महीने से अधिक समय"]},
 "s": ["Repair of {rproblem} in {area}", "{area} में {rproblem} की मरम्मत हेतु"],
 "b": [["I wish to draw your attention to the {rproblem} in {area}[[, which have not been repaired for {since}]].", "The road has become unsafe, especially after dark. Two-wheelers slip, pedestrians and children are at risk of accidents, and dark stretches encourage theft and harassment.", "I request you to get the repair work done at the earliest."],
       ["मैं आपका ध्यान {area} में {rproblem} की ओर दिलाना ((चाहता|चाहती)) हूँ[[, जिनकी मरम्मत {since} से नहीं हुई है]]।", "इससे मार्ग असुरक्षित हो गया है, विशेषकर अँधेरा होने के बाद। दोपहिया वाहन फिसलते हैं, पैदल चलने वालों और बच्चों को दुर्घटना का खतरा रहता है और अँधेरे में चोरी तथा छेड़छाड़ की घटनाओं का डर बना रहता है।", "आपसे निवेदन है कि शीघ्र मरम्मत कार्य करवाने की कृपा करें।"]]},
{"id": "water-complaint", "c": "govt", "n": "Water supply complaint", "h": "जल आपूर्ति की शिकायत", "kw": "water supply jal board no water dirty water pressure", "r": "person", "k": "comp",
 "to": ["The Executive Engineer", "अधिशासी अभियंता महोदय"], "org": ["[Water supply department / Jal Board]", "[जल आपूर्ति विभाग / जल बोर्ड]"], "so": ["Water Supply and Sewerage Board", "जल आपूर्ति एवं सीवरेज बोर्ड"],
 "f": ["area", "conno", "wproblem", "since"], "lb": {"conno": "Connection number (optional)", "since": "For how long (optional)"}, "sm": {"since": ["five days", "पाँच दिन"]},
 "s": ["Complaint about {wproblem} in {area}", "{area} में {wproblem} की शिकायत"],
 "b": [["I am a resident of {area}[[ (connection No. {conno})]]. Our area has been facing {wproblem}[[ for the last {since}]].", "Families are having to buy water or carry it from far away, and unsafe water can cause serious illness.", "Kindly look into the matter and restore a regular, clean water supply at the earliest."],
       ["मैं {area} ((का निवासी|की निवासी)) हूँ[[ (कनेक्शन संख्या {conno})]]।[[ पिछले {since} से]] हमारे क्षेत्र में {wproblem} की समस्या है।", "परिवारों को पानी खरीदना पड़ रहा है या दूर से लाना पड़ रहा है, और असुरक्षित पानी से गंभीर बीमारियाँ हो सकती हैं।", "कृपया मामले की जाँच कर शीघ्र नियमित एवं स्वच्छ जल आपूर्ति बहाल करवाएँ।"]]},
{"id": "electricity-complaint", "c": "govt", "n": "Electricity complaint (power cuts, bill, meter)", "h": "बिजली की शिकायत (कटौती, बिल, मीटर)", "kw": "electricity bill meter power cut voltage discom", "r": "person", "k": "comp",
 "to": ["The Assistant Engineer / SDO", "सहायक अभियंता / उपमंडल अधिकारी"], "org": ["[Electricity distribution company]", "[विद्युत वितरण निगम]"], "so": ["Power Distribution Corporation", "विद्युत वितरण निगम"],
 "f": ["conno", "eproblem", "since", "details"], "lb": {"since": "Since (optional)"}, "sm": {"since": ["last month", "पिछले महीने"], "details": ["The bill for October shows 1,240 units, about four times my usual use.", "अक्टूबर के बिल में 1,240 यूनिट दिखाए गए हैं, जो मेरी सामान्य खपत से लगभग चार गुना है।"]},
 "s": ["Complaint about {eproblem}, consumer No. {conno}", "{eproblem} की शिकायत, उपभोक्ता संख्या {conno}"],
 "b": [["I am a consumer of your department with consumer No. {conno}. I am writing about {eproblem} at my premises[[ since {since}]].", "[[{details} ]]Kindly send your staff to inspect the problem and set it right. If the bill or meter reading is wrong, please revise the bill and adjust any excess amount in the next bill."],
       ["मैं आपके विभाग ((का उपभोक्ता|की उपभोक्ता)) हूँ, मेरी उपभोक्ता संख्या {conno} है।[[ {since} से]] मेरे परिसर में {eproblem} की समस्या है।", "[[{details} ]]कृपया अपने कर्मचारियों को भेजकर समस्या की जाँच कराएँ और उसे ठीक करवाएँ। यदि बिल या मीटर रीडिंग गलत है, तो कृपया बिल संशोधित कर अधिक राशि अगले बिल में समायोजित करें।"]],
 "e": "Copy of the latest bill"},
{"id": "rti", "c": "govt", "n": "RTI application (Right to Information)", "h": "सूचना का अधिकार (आरटीआई) आवेदन", "kw": "rti right to information pio public information officer", "r": "person", "k": "none",
 "to": ["The Public Information Officer", "लोक सूचना अधिकारी"], "org": ["[Name of department / office]", "[विभाग / कार्यालय का नाम]"], "so": ["Office of the Municipal Corporation", "नगर निगम कार्यालय"],
 "f": ["info", "period", "fee"],
 "s": ["Application under Section 6(1) of the Right to Information Act, 2005", "सूचना का अधिकार अधिनियम, 2005 की धारा 6(1) के अंतर्गत आवेदन"],
 "b": [["I request the following information under the Right to Information Act, 2005[[, for the period {period}]]:\n{info}", "I have paid the application fee of ₹10 by {fee}.", "If any part of this information is held by another public authority, kindly transfer this application to it under Section 6(3) of the Act and inform me. Please send the information to my address given below."],
       ["मैं सूचना का अधिकार अधिनियम, 2005 के अंतर्गत[[ {period} की अवधि से संबंधित]] निम्नलिखित सूचना ((चाहता|चाहती)) हूँ:\n{info}", "मैंने ₹10 का आवेदन शुल्क {fee} द्वारा जमा किया है।", "यदि इस सूचना का कोई भाग किसी अन्य लोक प्राधिकरण के पास है, तो कृपया अधिनियम की धारा 6(3) के अंतर्गत यह आवेदन उसे स्थानांतरित कर मुझे सूचित करें। कृपया सूचना नीचे दिए गए पते पर भेजें।"]],
 "o": [["bpl", "I am below the poverty line (no fee)", "I belong to the Below Poverty Line (BPL) category, so no fee is payable; a copy of my BPL card is attached.", "मैं गरीबी रेखा से नीचे (बीपीएल) श्रेणी से हूँ, अतः कोई शुल्क देय नहीं है; बीपीएल कार्ड की प्रति संलग्न है।", 0],
       ["cert", "Ask for certified copies", "Kindly provide certified copies of the documents asked for; I will pay the copying charges.", "कृपया माँगे गए दस्तावेजों की प्रमाणित प्रतियाँ दें; प्रतिलिपि शुल्क मैं जमा कर ((दूँगा|दूँगी))।", 1]]},
{"id": "police-theft", "c": "govt", "n": "Theft complaint to police (FIR request)", "h": "चोरी की शिकायत (एफआईआर हेतु प्रार्थना पत्र)", "kw": "fir police theft stolen bike mobile robbery", "r": "person", "k": "comp",
 "to": ["The Station House Officer (SHO)", "थाना प्रभारी (एसएचओ)"], "org": ["[Police station name]", "[थाने का नाम]"], "so": ["Police Station Division No. 5", "थाना डिवीजन नंबर 5"],
 "f": ["item", "value", "d1", "time", "place", "details"], "lb": {"d1": "Date of theft", "place": "Where it happened"}, "sm": {"place": ["outside Ghumar Mandi market", "घुमार मंडी बाज़ार के बाहर"], "details": ["I had parked it and locked the handle; when I came back after twenty minutes, it was gone.", "मैंने इसे खड़ा करके हैंडल लॉक किया था; बीस मिनट बाद लौटने पर वह वहाँ नहीं थी।"]},
 "s": ["Request to register an FIR for theft of my {item}", "{item} की चोरी की प्राथमिकी (एफआईआर) दर्ज करने हेतु"],
 "b": [["I wish to report the theft of my {item}[[, worth about {value},]] on {d1}[[ at about {time}]] {place}.", "[[{details} ]]I searched the area and asked people nearby, but could not find it.", "I request you to register an FIR, investigate the matter and help me recover my property. Kindly give me a copy of the FIR, which I need for insurance and other formalities."],
       ["मैं सूचित करना ((चाहता|चाहती)) हूँ कि {d1} को[[ लगभग {time} बजे]] {place} से {item} की चोरी हो गई[[, जिसकी कीमत लगभग {value} है]]।", "[[{details} ]]मैंने आसपास खोजबीन की और लोगों से पूछताछ की, परंतु कुछ पता नहीं चला।", "आपसे निवेदन है कि प्राथमिकी दर्ज कर मामले की जाँच करें और मेरी संपत्ति बरामद कराने में सहायता करें। कृपया मुझे एफआईआर की प्रति दें, जो बीमा और अन्य औपचारिकताओं के लिए आवश्यक है।"]]},
{"id": "lost-report", "c": "govt", "n": "Lost documents or mobile report to police", "h": "दस्तावेज / मोबाइल गुम होने की सूचना (पुलिस)", "kw": "lost report gum missing documents mobile phone sim passport", "r": "person", "k": "none",
 "to": ["The Station House Officer (SHO)", "थाना प्रभारी (एसएचओ)"], "org": ["[Police station name]", "[थाने का नाम]"], "so": ["Police Station Division No. 5", "थाना डिवीजन नंबर 5"],
 "f": ["item", "d1", "place", "details"], "lb": {"item": "What was lost", "d1": "Date lost", "place": "Where (place)"}, "sm": {"item": ["wallet with my Aadhaar card, PAN card and driving licence", "आधार कार्ड, पैन कार्ड और ड्राइविंग लाइसेंस वाला बटुआ"], "place": ["the bus stand", "बस स्टैंड"], "details": ["I was travelling by city bus from the bus stand to Model Town.", "यह बस स्टैंड से मॉडल टाउन जाने वाली सिटी बस में यात्रा के दौरान हुआ।"]},
 "s": ["Report of loss of {item}", "{item} गुम होने की सूचना"],
 "b": [["I wish to report that I lost my {item} on {d1} at or near {place}. I have searched for it but could not find it.", "[[{details} ]]Kindly record this report and give me a copy, which I need to apply for duplicates and to prevent misuse. I will inform the police station if the items are found."],
       ["मैं सूचित करना ((चाहता|चाहती)) हूँ कि {d1} को {place} के आसपास {item} कहीं गुम हो जाने के कारण अब मेरे पास नहीं है। काफी खोजने पर भी उसका पता नहीं चला।", "[[{details} ]]कृपया इस गुमशुदगी की सूचना दर्ज कर मुझे उसकी प्रति दें, जो डुप्लीकेट बनवाने और दुरुपयोग रोकने के लिए आवश्यक है। वस्तु मिलने पर मैं थाने को सूचित कर ((दूँगा|दूँगी))।"]]},
{"id": "passport-cover", "c": "govt", "n": "Passport address proof covering letter", "h": "पासपोर्ट हेतु पते के प्रमाण का संलग्न पत्र", "kw": "passport seva kendra arn address proof covering letter", "r": "person", "k": "none",
 "to": ["The Passport Officer", "पासपोर्ट अधिकारी"], "org": ["Regional Passport Office", "क्षेत्रीय पासपोर्ट कार्यालय"],
 "f": ["arn", "apptdate", "docs"],
 "s": ["Submission of address proof for passport application, ARN {arn}", "पासपोर्ट आवेदन (एआरएन {arn}) हेतु पते का प्रमाण प्रस्तुत करना"],
 "b": [["I have applied for a passport online under Application Reference Number (ARN) {arn}[[, and my appointment is on {apptdate}]].", "As proof of my present address, I am submitting self-attested copies of the following documents:\n{docs}", "I confirm that I live at the address given below and that the details in my application are correct. Kindly accept these documents with my application."],
       ["मैंने एप्लीकेशन रेफरेंस नंबर (एआरएन) {arn} के अंतर्गत ऑनलाइन पासपोर्ट हेतु आवेदन किया है[[ और मेरी अपॉइंटमेंट {apptdate} को है]]।", "अपने वर्तमान पते के प्रमाण के रूप में मैं निम्नलिखित दस्तावेजों की स्व-सत्यापित प्रतियाँ प्रस्तुत कर ((रहा|रही)) हूँ:\n{docs}", "मैं पुष्टि ((करता|करती)) हूँ कि मैं नीचे दिए गए पते पर ((रहता|रहती)) हूँ और मेरे आवेदन का विवरण सही है। कृपया इन दस्तावेजों को मेरे आवेदन के साथ स्वीकार करें।"]]},
{"id": "aadhaar-cover", "c": "govt", "n": "Aadhaar address update covering letter", "h": "आधार में पता अद्यतन हेतु आवेदन", "kw": "aadhaar uidai address update enrolment centre", "r": "person", "k": "none",
 "to": ["The Officer-in-charge", "प्रभारी अधिकारी"], "org": ["Aadhaar Seva Kendra", "आधार सेवा केंद्र"],
 "f": ["uid", "newaddr", "docs"],
 "s": ["Request to update address in Aadhaar ending {uid}", "आधार (अंतिम अंक {uid}) में पता अद्यतन हेतु आवेदन"],
 "b": [["I wish to update the address in my Aadhaar, whose last four digits are {uid}. My new address is:\n{newaddr}", "As proof of the new address, I am attaching self-attested copies of:\n{docs}", "Kindly update my address. I will complete biometric verification at the centre, if required."],
       ["मैं अपने आधार, जिसके अंतिम चार अंक {uid} हैं, में पता अद्यतन कराना ((चाहता|चाहती)) हूँ। मेरा नया पता है:\n{newaddr}", "नए पते के प्रमाण के रूप में निम्नलिखित की स्व-सत्यापित प्रतियाँ संलग्न हैं:\n{docs}", "कृपया मेरा पता अद्यतन करें। आवश्यकता होने पर मैं केंद्र पर बायोमेट्रिक सत्यापन करा ((लूँगा|लूँगी))।"]]},

{"id": "landlord-repair", "c": "housing", "n": "Repair request to landlord", "h": "मकान मालिक को मरम्मत हेतु अनुरोध पत्र", "kw": "tenant landlord repair leakage maintenance", "r": "person", "k": "req", "rs": "n", "x": ["sphone"],
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["prop", "problem", "since"], "lb": {"problem": "What needs repair", "since": "Since (optional)"}, "sm": {"problem": ["the bathroom pipe is leaking and the wall has become damp", "बाथरूम का पाइप टपक रहा है और दीवार में सीलन आ गई है"], "since": ["last week", "पिछले सप्ताह"]},
 "s": ["Request for repairs at {prop}", "{prop} में मरम्मत हेतु अनुरोध"],
 "b": [["I am your tenant at {prop}. I would like to bring to your notice that {problem}[[, and this has been the case since {since}]].", "This is making daily life difficult and may cause further damage to the property if it is left as it is. I request you to arrange the repairs at the earliest. Please let me know when the workman will come so that someone is at home."],
       ["मैं {prop} में आपका किरायेदार हूँ। मैं आपके ध्यान में लाना ((चाहता|चाहती)) हूँ कि[[ {since} से]] {problem}।", "इससे दैनिक जीवन में कठिनाई हो रही है और समय पर ठीक न होने पर संपत्ति को और नुकसान हो सकता है। कृपया शीघ्र मरम्मत की व्यवस्था करें और बताएँ कि कारीगर कब आएगा, ताकि उस समय घर पर कोई उपस्थित रहे।"]],
 "o": [["self", "Offer to get it done and adjust in rent", "If it is easier for you, I can get the work done myself and adjust the cost, with bills, in next month's rent.", "यदि आपको सुविधा हो, तो मैं स्वयं काम करवाकर उसका खर्च बिलों सहित अगले माह के किराये में समायोजित कर ((सकता|सकती)) हूँ।", 0]]},
{"id": "vacate-notice", "c": "housing", "n": "Notice to vacate (tenant to landlord)", "h": "मकान खाली करने की सूचना (किरायेदार द्वारा)", "kw": "vacating notice leave house rent agreement", "r": "person", "k": "none", "rs": "n", "x": ["sphone"],
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["prop", "ldate", "deposit"], "lb": {"ldate": "Date you will move out"},
 "s": ["Notice to vacate {prop}", "{prop} खाली करने की सूचना"],
 "b": [["As per our rent agreement, I am giving you notice that I will vacate {prop} on {ldate}.", "I will clear the rent and all utility bills up to that date and hand over the keys with the house in good condition. Kindly arrange a joint inspection on the last day and return my security deposit of {deposit} when I hand over the keys."],
       ["किराया अनुबंध के अनुसार मैं आपको सूचित ((करता|करती)) हूँ कि मैं {ldate} को {prop} खाली कर ((दूँगा|दूँगी))।", "उस तिथि तक का किराया और सभी बिल मैं चुका ((दूँगा|दूँगी)) और मकान अच्छी स्थिति में चाबियाँ सौंप ((दूँगा|दूँगी))। कृपया अंतिम दिन संयुक्त निरीक्षण की व्यवस्था करें और चाबियाँ सौंपते समय मेरी {deposit} की जमानत राशि लौटा दें।"]],
 "t": {"polite": ["Thank you for being so helpful during my stay.", "यहाँ रहने के दौरान आपके सहयोग के लिए धन्यवाद।"], "firm": ["Kindly acknowledge this notice in writing.", "कृपया इस सूचना की लिखित पावती दें।"]}},
{"id": "deposit-refund", "c": "housing", "n": "Security deposit refund request", "h": "जमानत राशि (सिक्योरिटी) वापसी हेतु पत्र", "kw": "deposit refund landlord advance", "r": "person", "k": "req", "rs": "n", "x": ["saddr", "sphone"],
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["prop", "d1", "deposit", "pay"], "lb": {"d1": "Date you moved out"},
 "s": ["Refund of security deposit for {prop}", "{prop} की जमानत राशि वापसी हेतु"],
 "b": [["I vacated {prop} on {d1} after clearing all rent and bills, and handed over the keys to you. My security deposit of {deposit} has not been returned yet.", "Kindly refund the deposit[[ to {pay}]] at the earliest. If you plan to deduct anything for damage, please send me the details with bills."],
       ["मैंने सारा किराया और बिल चुकाकर {d1} को {prop} खाली कर दिया था और चाबियाँ आपको सौंप दी थीं। मेरी {deposit} की जमानत राशि अभी तक वापस नहीं मिली है।", "कृपया शीघ्र जमानत राशि लौटाएँ[[ ({pay})]]। यदि किसी नुकसान के लिए कटौती करनी हो, तो कृपया बिलों सहित उसका विवरण भेजें।"]],
 "t": {"firm": ["If the deposit is not refunded within 15 days, I will have to take legal steps to recover it.", "यदि 15 दिनों में जमानत राशि वापस नहीं मिली, तो मुझे इसकी वसूली के लिए कानूनी कदम उठाने पड़ेंगे।"]}},
{"id": "rent-increase", "c": "housing", "n": "Rent increase notice (landlord to tenant)", "h": "किराया वृद्धि की सूचना (मकान मालिक द्वारा)", "kw": "rent hike revision landlord tenant", "r": "person", "k": "none", "rs": "n", "x": ["sphone"],
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["prop", "rent", "newrent", "d1"], "lb": {"d1": "New rent from"},
 "s": ["Revision of rent for {prop}", "{prop} के किराये में संशोधन"],
 "b": [["As per the terms of our rent agreement, I am writing to inform you that the monthly rent for {prop} will be revised from {rent} to {newrent} with effect from {d1}.", "This revision is in line with the yearly increase agreed in the agreement and the rents now charged in the area. All other terms remain the same. Kindly confirm your acceptance in writing."],
       ["किराया अनुबंध की शर्तों के अनुसार आपको सूचित करना है कि {d1} से {prop} का मासिक किराया {rent} से संशोधित कर {newrent} किया जाएगा।", "यह संशोधन अनुबंध में तय वार्षिक वृद्धि और क्षेत्र में वर्तमान किरायों के अनुरूप है। अन्य सभी शर्तें यथावत रहेंगी। कृपया लिखित में अपनी सहमति दें।"]],
 "t": {"polite": ["Thank you for being a good tenant. Please feel free to talk to me if you have any questions.", "एक अच्छे किरायेदार होने के लिए धन्यवाद। कोई प्रश्न हो तो निःसंकोच मुझसे बात करें।"]}},
{"id": "rent-reminder", "c": "housing", "n": "Rent payment reminder (landlord to tenant)", "h": "बकाया किराये का स्मरण पत्र", "kw": "rent due reminder pending tenant", "r": "person", "k": "none", "rs": "n", "x": ["sphone"],
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["prop", "mfor", "amt", "by"], "lb": {"amt": "Total due", "by": "Pay by"}, "sm": {"amt": ["30000", "30000"]},
 "s": ["Reminder for pending rent of {prop}", "{prop} के बकाया किराये हेतु स्मरण पत्र"],
 "b": [["This is a reminder that the rent for {prop} for {mfor} is still pending. The total amount due is {amt}.", "Kindly pay the amount by {by}."],
       ["यह स्मरण कराना है कि {prop} का {mfor} का किराया अभी बकाया है। कुल देय राशि {amt} है।", "कृपया यह राशि {by} तक जमा करें।"]],
 "t": {"formal": ["If you have already paid, please ignore this letter and share the payment details.", "यदि आप भुगतान कर चुके हैं, तो कृपया यह पत्र अनदेखा करें और भुगतान का विवरण भेज दें।"],
       "polite": ["I understand that delays can happen. If you are facing any difficulty, please talk to me so that we can work out a solution.", "मैं ((समझता|समझती)) हूँ कि कभी-कभी देरी हो जाती है। यदि कोई कठिनाई हो, तो कृपया मुझसे बात करें ताकि हम कोई समाधान निकाल सकें।"],
       "firm": ["Please note that if the payment is not received by this date, I will have to act as per the rent agreement, which may include adjusting the security deposit and asking you to vacate.", "कृपया ध्यान दें कि इस तिथि तक भुगतान न मिलने पर मुझे किराया अनुबंध के अनुसार कार्रवाई करनी पड़ेगी, जिसमें जमानत राशि से समायोजन और मकान खाली करने को कहना शामिल हो सकता है।"]}},
{"id": "society-complaint", "c": "housing", "n": "Complaint to housing society / RWA", "h": "सोसाइटी / आरडब्ल्यूए को शिकायत पत्र", "kw": "rwa society secretary parking noise water leakage neighbour", "r": "person", "k": "comp", "x": ["flat", "sphone"],
 "to": ["The Secretary", "सचिव"], "org": ["[Society / RWA name]", "[सोसाइटी / आरडब्ल्यूए का नाम]"],
 "f": ["flat", "issue", "since"], "lb": {"issue": "The problem", "since": "For how long (optional)"}, "sm": {"issue": ["cars being parked in front of the gate and blocking the exit", "गेट के सामने कारें खड़ी कर रास्ता रोकने"], "since": ["a month", "एक महीने"]},
 "s": ["Complaint regarding {issue}", "{issue} के संबंध में शिकायत"],
 "b": [["I live in Flat {flat} and wish to bring to the committee's notice the problem of {issue}[[, which has continued for {since}]].", "Several residents are facing the same difficulty, and it is affecting our peace and safety. I request the committee to look into it, discuss it at the next meeting if needed, and take suitable action."],
       ["मैं फ्लैट {flat} ((का निवासी|की निवासी)) हूँ और समिति का ध्यान {issue} की समस्या की ओर दिलाना ((चाहता|चाहती)) हूँ[[, जो {since} से बनी हुई है]]।", "कई निवासी इसी परेशानी का सामना कर रहे हैं और इससे हमारी शांति और सुरक्षा प्रभावित हो रही है। कृपया समिति इस मामले पर ध्यान दे, आवश्यकता हो तो अगली बैठक में चर्चा करे और उचित कार्रवाई करे।"]]},
{"id": "society-noc", "c": "housing", "n": "NOC request from housing society", "h": "सोसाइटी से अनापत्ति प्रमाण पत्र हेतु आवेदन", "kw": "society noc sale rent loan flat transfer", "r": "person", "k": "req", "x": ["flat", "sphone"],
 "to": ["The Secretary", "सचिव"], "org": ["[Society name]", "[सोसाइटी का नाम]"],
 "f": ["flat", "purpose"], "sm": {"purpose": ["the sale of the flat", "फ्लैट की बिक्री"]},
 "s": ["Request for a No Objection Certificate for Flat {flat}", "फ्लैट {flat} हेतु अनापत्ति प्रमाण पत्र के लिए आवेदन"],
 "b": [["I am the owner of Flat {flat} in our society. I need a No Objection Certificate from the society for {purpose}.", "I confirm that all maintenance charges and other dues have been paid up to date. Kindly issue the NOC on the society letterhead at the earliest."],
       ["मैं सोसाइटी के फ्लैट {flat} ((का स्वामी|की स्वामिनी)) हूँ। मुझे {purpose} के लिए सोसाइटी से अनापत्ति प्रमाण पत्र की आवश्यकता है।", "मैं पुष्टि ((करता|करती)) हूँ कि आज तक का सारा रखरखाव शुल्क और अन्य बकाया चुका दिया गया है। कृपया सोसाइटी के लेटरहेड पर शीघ्र अनापत्ति प्रमाण पत्र जारी करें।"]]},

{"id": "quotation-request", "c": "business", "n": "Request for quotation", "h": "कोटेशन (मूल्य) हेतु अनुरोध पत्र", "kw": "rfq quotation enquiry price rates supplier", "r": "org", "k": "none",
 "to": ["The Sales Manager", "विक्रय प्रबंधक"], "org": ["[Supplier name]", "[आपूर्तिकर्ता का नाम]"],
 "f": ["items", "dplace", "by"], "lb": {"by": "Quotation needed by", "dplace": "Delivery address (optional)"},
 "s": ["Request for quotation", "कोटेशन हेतु अनुरोध"],
 "b": [["We are interested in buying the following and request you to send your best quotation:\n{items}", "Please mention the unit price, applicable GST, delivery time, payment terms, warranty and the validity of the quotation.[[ Delivery will be required at {dplace}.]]", "We would appreciate receiving your quotation by {by}. For any clarification, please contact {sname} on {sphone}."],
       ["हम निम्नलिखित सामान खरीदना चाहते हैं, कृपया अपनी सर्वोत्तम दरें (कोटेशन) भेजें:\n{items}", "कृपया इकाई मूल्य, लागू जीएसटी, डिलीवरी का समय, भुगतान की शर्तें, वारंटी और कोटेशन की वैधता अवश्य लिखें।[[ डिलीवरी {dplace} पर चाहिए।]]", "कृपया अपना कोटेशन {by} तक भेजें। किसी भी जानकारी के लिए {sname} से {sphone} पर संपर्क करें।"]]},
{"id": "purchase-order", "c": "business", "n": "Purchase order letter", "h": "क्रय आदेश (परचेज़ ऑर्डर) पत्र", "kw": "po order placing order supplier", "r": "org", "k": "none",
 "to": ["The Sales Manager", "विक्रय प्रबंधक"], "org": ["[Supplier name]", "[आपूर्तिकर्ता का नाम]"],
 "f": ["pono", "qref", "items", "amt", "by", "dplace", "terms"], "lb": {"amt": "Total order value (with GST)", "by": "Deliver by"}, "sm": {"amt": ["84960", "84960"]},
 "s": ["Purchase order No. {pono}", "क्रय आदेश संख्या {pono}"],
 "b": [["We are pleased to place an order for the following[[ as per your quotation {qref}]]:\n{items}", "The total order value is {amt}[[ ({amtw})]], including GST. Please deliver the goods to {dplace} by {by}, with the tax invoice and delivery challan quoting this order number.", "Payment terms: {terms}. Goods that do not match the agreed specifications will be returned at your cost. Kindly confirm acceptance of this order."],
       ["हम[[ आपके कोटेशन {qref} के अनुसार]] निम्नलिखित सामान का आदेश देते हैं:\n{items}", "आदेश का कुल मूल्य जीएसटी सहित {amt} है। कृपया इस आदेश संख्या के उल्लेख के साथ टैक्स इनवॉइस और डिलीवरी चालान सहित सामान {by} तक {dplace} पर पहुँचाएँ।", "भुगतान की शर्तें: {terms}। तय विनिर्देशों के अनुरूप न होने वाला सामान आपके खर्च पर लौटाया जाएगा। कृपया इस आदेश की स्वीकृति की पुष्टि करें।"]]},
{"id": "payment-reminder", "c": "business", "n": "Payment reminder to a customer", "h": "ग्राहक को भुगतान स्मरण पत्र", "kw": "overdue invoice dues outstanding collection follow up", "r": "org", "k": "remind",
 "to": ["The Accounts Manager", "लेखा प्रबंधक"], "org": ["[Customer company name]", "[ग्राहक संस्था का नाम]"], "so": ["Gupta & Co.", "गुप्ता एंड कंपनी"],
 "f": ["invno", "invdate", "amt", "due", "by", "bankd"], "lb": {"amt": "Amount due", "by": "Pay by"}, "sm": {"amt": ["118000", "118000"]},
 "s": ["Payment reminder: invoice No. {invno}", "भुगतान स्मरण पत्र: इनवॉइस संख्या {invno}"],
 "b": [["This is to remind you that payment of {amt} against our invoice No. {invno} dated {invdate} was due on {due} and has not been received so far.", "Kindly arrange the payment by {by}[[ to our bank account: {bankd}]]."],
       ["यह स्मरण कराना है कि हमारे इनवॉइस संख्या {invno} दिनांक {invdate} की {amt} की राशि {due} को देय थी, जो अब तक प्राप्त नहीं हुई है।", "कृपया {by} तक भुगतान की व्यवस्था करें[[ (हमारा बैंक विवरण: {bankd})]]।"]],
 "o": [["msme", "Mention the MSME 45-day rule", "As a registered MSME, we are entitled under the MSMED Act, 2006 to payment within the agreed period, which cannot exceed 45 days, failing which interest is payable.", "पंजीकृत एमएसएमई होने के कारण, एमएसएमईडी अधिनियम, 2006 के अंतर्गत हमें तय अवधि (अधिकतम 45 दिन) में भुगतान पाने का अधिकार है, अन्यथा ब्याज देय होता है।", 0]]},
{"id": "apology-customer", "c": "business", "n": "Apology letter to a customer", "h": "ग्राहक को क्षमा-याचना पत्र", "kw": "sorry delay wrong delivery service complaint reply", "r": "org", "k": "none", "rs": "n",
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["orderno", "issue", "fix"], "lb": {"issue": "What went wrong"}, "sm": {"issue": ["the late delivery", "डिलीवरी में देरी"]},
 "s": ["Apology for {issue}", "{issue} के लिए क्षमा-याचना"],
 "b": [["Thank you for bringing this to our attention. We sincerely apologise for {issue}[[ in your order No. {orderno}]]. This is not the standard of service we aim to give, and we understand the inconvenience it caused you.", "To put things right, {fix}. We have also reviewed our process so that this does not happen again.", "We value your trust and hope to serve you better in future. Please contact {sname} on {sphone} if there is anything more we can do."],
       ["इस ओर हमारा ध्यान दिलाने के लिए धन्यवाद।[[ आपके ऑर्डर संख्या {orderno} में]] {issue} के लिए हम हृदय से क्षमा चाहते हैं। यह हमारी सेवा के मानक के अनुरूप नहीं है और हम समझते हैं कि इससे आपको असुविधा हुई।", "भूल सुधारने के लिए {fix}। हमने अपनी प्रक्रिया की भी समीक्षा की है, ताकि ऐसा दोबारा न हो।", "हम आपके विश्वास को महत्व देते हैं और भविष्य में बेहतर सेवा देने की आशा करते हैं। किसी भी सहायता के लिए {sname} से {sphone} पर संपर्क करें।"]]},
{"id": "thanks-customer", "c": "business", "n": "Thank-you letter to a customer", "h": "ग्राहक को धन्यवाद पत्र", "kw": "thank you client appreciation loyal customer", "r": "org", "k": "none", "rs": "n", "cl": "warm",
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["occasion", "offer"],
 "s": ["Thank you for your business", "आपके सहयोग के लिए धन्यवाद"],
 "b": [["On behalf of everyone at {sorg}, thank you for {occasion}. Customers like you are the reason our business keeps growing.", "[[As a small token of our appreciation, {offer}. ]]We look forward to serving you again. If you ever need anything, we are just a phone call away."],
       ["{sorg} परिवार की ओर से {occasion} के लिए हार्दिक धन्यवाद। आप जैसे ग्राहकों के कारण ही हमारा व्यवसाय निरंतर बढ़ रहा है।", "[[आभार के रूप में {offer}। ]]हम आपकी फिर से सेवा करने के लिए उत्सुक हैं। किसी भी आवश्यकता पर बस एक फोन करें।"]]},
{"id": "supplier-complaint", "c": "business", "n": "Complaint to supplier about defective goods", "h": "आपूर्तिकर्ता को दोषपूर्ण सामान की शिकायत", "kw": "defective damaged goods wrong supply replacement return", "r": "org", "k": "comp",
 "to": ["The Sales Manager", "विक्रय प्रबंधक"], "org": ["[Supplier name]", "[आपूर्तिकर्ता का नाम]"],
 "f": ["pono", "d1", "issue", "ask"], "lb": {"d1": "Goods received on", "issue": "What was wrong", "ask": "What you want them to do"}, "sm": {"issue": ["12 of the 50 cartons were damaged and the bottles inside were broken", "50 में से 12 कार्टन क्षतिग्रस्त थे और उनमें रखी बोतलें टूटी हुई थीं"], "ask": ["replace the damaged cartons within 7 days", "7 दिनों के भीतर क्षतिग्रस्त कार्टन बदल दें"]},
 "s": ["Complaint about defective goods against order No. {pono}", "आदेश संख्या {pono} के दोषपूर्ण सामान की शिकायत"],
 "b": [["We received the goods against our purchase order No. {pono} on {d1}. On inspection, we found that {issue}.", "This has affected our work and our commitments to our own customers. We request you to {ask}, at your cost.", "The defective goods have been kept aside for your inspection, and photographs are attached for your reference."],
       ["आदेश संख्या {pono} के अंतर्गत भेजा गया सामान हमें {d1} को प्राप्त हुआ। जाँच करने पर पाया गया कि {issue}।", "इससे हमारा कार्य और हमारे ग्राहकों के प्रति हमारी प्रतिबद्धताएँ प्रभावित हुई हैं। आपसे अनुरोध है कि अपने खर्च पर {ask}।", "दोषपूर्ण सामान आपके निरीक्षण के लिए अलग रखा गया है। संदर्भ हेतु फोटो संलग्न हैं।"]],
 "t": {"firm": ["If this is not resolved within 7 days, we will be forced to return the full consignment, withhold payment and look for another supplier.", "यदि 7 दिनों में इसका समाधान नहीं हुआ, तो हमें पूरा माल लौटाने, भुगतान रोकने और दूसरा आपूर्तिकर्ता ढूँढने के लिए विवश होना पड़ेगा।"]}},
{"id": "order-cancel", "c": "business", "n": "Order cancellation letter", "h": "आदेश (ऑर्डर) निरस्त करने का पत्र", "kw": "cancel order cancellation refund advance", "r": "org", "k": "none",
 "to": ["The Sales Manager", "विक्रय प्रबंधक"], "org": ["[Supplier name]", "[आपूर्तिकर्ता का नाम]"],
 "f": ["pono", "d1", "reason", "amt"], "lb": {"d1": "Order date", "reason": "Reason (optional)", "amt": "Advance paid (optional)"}, "sm": {"reason": ["the project for which the material was ordered has been postponed", "जिस परियोजना के लिए सामान मँगाया गया था वह स्थगित हो गई है"], "amt": ["20000", "20000"]},
 "s": ["Cancellation of order No. {pono}", "आदेश संख्या {pono} का निरस्तीकरण"],
 "b": [["We regret to inform you that we wish to cancel our order No. {pono} dated {d1}[[, as {reason}]].", "As the goods have not been dispatched yet, we request you not to process the order further[[ and to refund the advance of {amt}]]. Kindly confirm the cancellation in writing.", "We apologise for any inconvenience and look forward to doing business with you again."],
       ["खेद के साथ सूचित करना है कि हम अपना आदेश संख्या {pono} दिनांक {d1} निरस्त करना चाहते हैं[[, क्योंकि {reason}]]।", "चूँकि सामान अभी भेजा नहीं गया है, कृपया इस आदेश पर आगे कार्रवाई न करें[[ और {amt} की अग्रिम राशि लौटा दें]]। कृपया लिखित में निरस्तीकरण की पुष्टि करें।", "किसी भी असुविधा के लिए हमें खेद है और भविष्य में आपके साथ व्यापार की आशा है।"]]},
{"id": "business-intro", "c": "business", "n": "Business introduction letter", "h": "व्यवसाय परिचय पत्र", "kw": "introduce company sales letter new client", "r": "org", "k": "none",
 "to": ["The Purchase Manager", "क्रय प्रबंधक"], "org": ["[Company name]", "[संस्था का नाम]"], "so": ["City Hospital", "सिटी हॉस्पिटल"],
 "f": ["offer2", "since", "clients"], "lb": {"since": "In business since (optional)"}, "sm": {"since": ["2012", "2012"]},
 "s": ["Introducing {sorg}", "{sorg} का परिचय"],
 "b": [["We would like to introduce {sorg}[[, in business since {since}]]. We offer {offer2}.", "[[Our clients include {clients}. ]]We focus on good quality, fair prices and timely delivery, and we would be glad to understand your needs and send you a proposal made for you.", "May we call on you for a short meeting at your convenience? You can reach {sname} on {sphone}."],
       ["हम आपको {sorg} का परिचय देना चाहते हैं[[, जो {since} से कार्यरत है]]। हम {offer2} प्रदान करते हैं।", "[[हमारे ग्राहकों में {clients} शामिल हैं। ]]हम अच्छी गुणवत्ता, उचित मूल्य और समय पर डिलीवरी पर ध्यान देते हैं, और आपकी आवश्यकताएँ समझकर आपके अनुरूप प्रस्ताव भेजने में हमें प्रसन्नता होगी।", "क्या हम आपकी सुविधानुसार एक छोटी बैठक के लिए आपसे मिल सकते हैं? आप {sname} से {sphone} पर संपर्क कर सकते हैं।"]]},

{"id": "invitation", "c": "personal", "n": "Formal invitation letter", "h": "निमंत्रण पत्र", "kw": "invite function wedding housewarming ceremony party", "r": "person", "k": "none", "rs": "n", "cl": "warm", "x": ["sphone"],
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["event", "d1", "time", "venue"], "lb": {"d1": "Event date"}, "sm": {"event": ["the housewarming ceremony of our new home", "हमारे नए घर के गृह प्रवेश समारोह"], "time": ["11:00 am", "सुबह 11:00"]},
 "s": ["Invitation to {event}", "{event} हेतु निमंत्रण"],
 "b": [["It gives us great pleasure to invite you and your family to {event} on {d1}[[ at {time}]] at {venue}.", "Your presence will make the occasion truly special for us. We look forward to welcoming you."],
       ["आपको और आपके परिवार को {d1} को[[ {time} बजे]] {venue} में {event} में आमंत्रित करते हुए हमें अत्यंत प्रसन्नता हो रही है।", "आपकी उपस्थिति इस अवसर को हमारे लिए विशेष बना देगी। हम आपके स्वागत की प्रतीक्षा करेंगे।"]],
 "o": [["meal", "Mention a meal after the function", "Please join us for a meal after the function.", "कार्यक्रम के बाद भोजन के लिए अवश्य पधारें।", 1],
       ["rsvp", "Ask them to confirm", "Kindly let us know by phone ({sphone}) whether you will be able to come.", "कृपया फोन ({sphone}) पर अपने आने की सूचना दें।", 1]], "oa": 1},
{"id": "recommendation", "c": "personal", "n": "Recommendation letter", "h": "अनुशंसा (सिफारिश) पत्र", "kw": "reference letter recommend student employee lor", "r": "person", "p": ["sname", "sg", "sdesig", "sorg", "sphone", "semail"], "x": ["sdesig", "sorg", "sphone", "semail"], "k": "none", "tw": 1,
 "f": ["person", "pg", "prel", "known", "quality", "purpose"], "sm": {"purpose": ["admission to the MBA programme at your university", "आपके विश्वविद्यालय के एमबीए कार्यक्रम में प्रवेश"], "sdesig": ["Associate Professor of Commerce", "वाणिज्य के एसोसिएट प्रोफेसर"]},
 "s": ["Letter of recommendation", "अनुशंसा पत्र"],
 "b": [["I am pleased to recommend ((pg:Mr.|Ms.)) {person}, whom I have known for {known} as {prel}.", "((pg:He|She)) is {quality}. I have always found ((pg:him|her)) honest, dependable and willing to take responsibility, and ((pg:he|she)) works well with people.", "I recommend ((pg:him|her)) without reservation for {purpose}. Please feel free to contact me for any further information."],
       ["मैं ((pg:श्री|सुश्री)) {person} की अनुशंसा करते हुए प्रसन्नता अनुभव ((करता|करती)) हूँ, जिन्हें मैं {known} से {prel} के रूप में ((जानता|जानती)) हूँ।", "वे {quality} हैं। मैंने उन्हें सदा ईमानदार, भरोसेमंद और जिम्मेदारी लेने ((pg:वाला|वाली)) पाया है, और लोगों के साथ उनका व्यवहार अच्छा है।", "मैं बिना किसी संकोच के {purpose} के लिए उनकी अनुशंसा ((करता|करती)) हूँ। अधिक जानकारी के लिए मुझसे संपर्क किया जा सकता है।"]]},
{"id": "authority-letter", "c": "personal", "n": "Authority letter (authorise someone)", "h": "प्राधिकार पत्र (किसी को अधिकृत करना)", "kw": "authorization authorisation collect on behalf representative", "r": "person", "k": "none",
 "to": ["The Branch Manager / Officer-in-charge", "शाखा प्रबंधक / प्रभारी अधिकारी"], "org": ["[Office / bank name]", "[कार्यालय / बैंक का नाम]"], "so": ["State Bank of India, Model Town branch", "भारतीय स्टेट बैंक, मॉडल टाउन शाखा"],
 "f": ["person", "pg", "prel", "idp", "task", "d1"], "lb": {"prel": "Their relation to you (optional)", "d1": "On date (optional)"}, "sm": {"prel": ["brother", "भाई"]},
 "s": ["Authority letter in favour of {person}", "{person} के पक्ष में प्राधिकार पत्र"],
 "b": [["I, {sname}, hereby authorise ((pg:Mr.|Ms.)) {person}[[, my {prel},]] to {task} on my behalf[[ on {d1}]], as I am unable to come in person.", "((pg:His|Her)) identity proof is {idp}. Kindly allow ((pg:him|her)) to complete this work. ((pg:His|Her)) signature is given below for verification.", "Signature of the authorised person: ____________________"],
       ["मैं, {sname}, स्वयं उपस्थित न हो पाने के कारण[[ अपने {prel}]] ((pg:श्री|सुश्री)) {person} को[[ {d1} को]] मेरी ओर से {task} के लिए अधिकृत ((करता|करती)) हूँ।", "उनका पहचान प्रमाण {idp} है। कृपया उन्हें यह कार्य पूरा करने दें। सत्यापन हेतु उनके हस्ताक्षर नीचे दिए गए हैं।", "अधिकृत व्यक्ति के हस्ताक्षर: ____________________"]],
 "o": [["id", "Mention ID copies attached", "Self-attested copies of my identity proof and ((pg:his|her)) identity proof are attached.", "मेरे और उनके पहचान प्रमाण की स्व-सत्यापित प्रतियाँ संलग्न हैं।", 1]], "oa": 2},
{"id": "char-request", "c": "personal", "n": "Character certificate request (sarpanch / councillor)", "h": "चरित्र प्रमाण पत्र हेतु आवेदन (सरपंच / पार्षद)", "kw": "character certificate sarpanch councillor gazetted officer police verification", "r": "person", "k": "req",
 "to": ["The Sarpanch / Ward Councillor", "सरपंच / पार्षद महोदय"], "org": ["[Gram Panchayat / Ward No.]", "[ग्राम पंचायत / वार्ड संख्या]"], "so": ["Ward No. 12", "वार्ड संख्या 12"],
 "f": ["since", "purpose"], "lb": {"since": "Living here since"}, "sm": {"since": ["2010", "2010"], "purpose": ["a government job application", "सरकारी नौकरी के आवेदन"]},
 "s": ["Request for a character certificate for {purpose}", "{purpose} हेतु चरित्र प्रमाण पत्र के लिए आवेदन"],
 "b": [["I am a permanent resident of this area and have been living at the address given below since {since}. You know me and my family well.", "I need a character certificate for {purpose}. I have never been involved in any criminal case or unlawful activity. Kindly issue the certificate at the earliest."],
       ["मैं इस क्षेत्र ((का स्थायी निवासी|की स्थायी निवासी)) हूँ और {since} से नीचे दिए गए पते पर रह ((रहा|रही)) हूँ। आप मुझे और मेरे परिवार को भली-भाँति जानते हैं।", "मुझे {purpose} के लिए चरित्र प्रमाण पत्र की आवश्यकता है। मैं कभी किसी आपराधिक मामले या गैरकानूनी गतिविधि में शामिल नहीं ((रहा|रही))। कृपया शीघ्र प्रमाण पत्र जारी करने की कृपा करें।"]],
 "o": [["docs", "Mention Aadhaar copy and photo", "Copies of my Aadhaar card and a recent photograph are attached.", "मेरे आधार कार्ड की प्रति और हाल का फोटो संलग्न है।", 1]]},
{"id": "condolence", "c": "personal", "n": "Condolence letter", "h": "शोक संवेदना पत्र", "kw": "sympathy death passing away sorry for your loss", "r": "person", "k": "none", "rs": "n", "cl": "warm", "x": [],
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["deceased", "drel"],
 "s": ["Heartfelt condolences", "हार्दिक संवेदना"],
 "b": [["I was deeply saddened to learn of the passing of {deceased}[[, your {drel}]]. Please accept my heartfelt condolences.", "I will always remember their kindness and warmth. Words cannot ease the pain of such a loss, but please know that you and your family are in my thoughts and prayers.", "If there is anything at all I can do to help in the days ahead, please do not hesitate to tell me."],
       ["[[आपके {drel} ]]{deceased} के निधन का समाचार सुनकर मुझे गहरा दुःख हुआ। कृपया मेरी हार्दिक संवेदना स्वीकार करें।", "उनका स्नेह और आत्मीयता मुझे सदा याद रहेगी। शब्द इस क्षति की पीड़ा को कम नहीं कर सकते, पर आप और आपका परिवार मेरी प्रार्थनाओं में हैं। ईश्वर दिवंगत आत्मा को शांति प्रदान करे।", "आने वाले दिनों में यदि मैं किसी भी प्रकार से सहायता कर सकूँ, तो निःसंकोच बताएँ।"]]},
{"id": "consumer-complaint", "c": "personal", "n": "Consumer complaint to a company", "h": "कंपनी को उपभोक्ता शिकायत पत्र", "kw": "defective product refund replacement warranty customer care consumer forum", "r": "person", "k": "comp",
 "to": ["The Customer Care Manager", "ग्राहक सेवा प्रबंधक"], "org": ["[Company name]", "[कंपनी का नाम]"], "so": ["Bright Home Appliances Ltd.", "ब्राइट होम अप्लायंसेज़ लि."],
 "f": ["product", "d1", "orderno", "amt", "issue", "ask"], "lb": {"d1": "Date of purchase", "amt": "Price paid (optional)", "issue": "What went wrong"}, "sm": {"amt": ["32500", "32500"], "issue": ["the machine stopped working within 20 days and the technician has not come despite two complaints", "मशीन 20 दिनों में ही बंद हो गई और दो शिकायतों के बाद भी तकनीशियन नहीं आया"]},
 "s": ["Complaint about {product}", "{product} के संबंध में शिकायत"],
 "b": [["I bought {product} from you on {d1}[[ (order / invoice No. {orderno})]][[ for {amt}]]. Unfortunately, {issue}.", "I request you to {ask}. Copies of the invoice and other details are attached."],
       ["मैंने {d1} को आपसे {product} की खरीद की थी[[ (ऑर्डर / इनवॉइस संख्या {orderno})]][[, जिसका मूल्य {amt} था]]। दुर्भाग्य से {issue}।", "आपसे अनुरोध है कि {ask}। इनवॉइस और अन्य विवरण की प्रतियाँ संलग्न हैं।"]],
 "o": [["earlier", "Say you already contacted customer care", "I reported this to your customer care earlier, but the problem has not been solved.", "मैंने पहले भी आपकी ग्राहक सेवा को इसकी सूचना दी थी, परंतु समस्या का समाधान नहीं हुआ।", 1]], "oa": 1,
 "t": {"firm": ["If the matter is not resolved within 15 days, I will file a complaint on the National Consumer Helpline (1915) and before the District Consumer Commission, entirely at your cost and risk.", "यदि 15 दिनों में समस्या का समाधान नहीं हुआ, तो मैं राष्ट्रीय उपभोक्ता हेल्पलाइन (1915) तथा जिला उपभोक्ता आयोग में शिकायत दर्ज ((करूँगा|करूँगी)), जिसकी पूरी जिम्मेदारी आपकी होगी।"]}},
{"id": "insurance-claim", "c": "personal", "n": "Insurance claim intimation", "h": "बीमा दावे की सूचना", "kw": "insurance claim accident damage motor health policy", "r": "person", "k": "none",
 "to": ["The Claims Manager", "दावा प्रबंधक"], "org": ["[Insurance company name]", "[बीमा कंपनी का नाम]"], "so": ["Secure General Insurance Co. Ltd.", "सिक्योर जनरल इंश्योरेंस कं. लि."],
 "f": ["policy", "d1", "issue", "amt"], "lb": {"d1": "Date of the incident", "issue": "What happened", "amt": "Estimated loss (optional)"}, "sm": {"issue": ["my car (PB10 CD 5678) was damaged in a road accident near Khanna", "खन्ना के पास सड़क दुर्घटना में मेरी कार (PB10 CD 5678) क्षतिग्रस्त हो गई"], "amt": ["60000", "60000"]},
 "s": ["Intimation of claim under policy No. {policy}", "पॉलिसी संख्या {policy} के अंतर्गत दावे की सूचना"],
 "b": [["I hold policy No. {policy} with your company. I wish to inform you that on {d1}, {issue}.", "[[The estimated loss is about {amt}. ]]Kindly register my claim, appoint a surveyor if needed and let me know the documents and forms required. I will co-operate fully with the claim process."],
       ["आपकी कंपनी में मेरी पॉलिसी संख्या {policy} है। सूचित करना है कि {d1} को {issue}।", "[[अनुमानित हानि लगभग {amt} है। ]]कृपया मेरा दावा दर्ज करें, आवश्यकता हो तो सर्वेयर नियुक्त करें और आवश्यक दस्तावेज एवं फॉर्म की जानकारी दें। दावा प्रक्रिया में मैं पूरा सहयोग ((करूँगा|करूँगी))।"]]},
{"id": "thanks-personal", "c": "personal", "n": "Thank-you letter (personal)", "h": "धन्यवाद पत्र (व्यक्तिगत)", "kw": "thanks gratitude help appreciation", "r": "person", "k": "none", "rs": "n", "cl": "warm", "x": [],
 "to": ["{rname}", "{rname}"], "org": ["", ""],
 "f": ["help"],
 "s": ["Thank you", "धन्यवाद"],
 "b": [["I am writing to thank you sincerely for {help}. Your help came at a time when I needed it most, and it made a real difference.", "I am truly grateful for your kindness and hope I can return it some day."],
       ["{help} के लिए मैं आपका हृदय से धन्यवाद करना ((चाहता|चाहती)) हूँ। आपकी सहायता ऐसे समय मिली जब मुझे उसकी सबसे अधिक आवश्यकता थी, और इससे सचमुच बड़ा अंतर आया।", "आपकी उदारता के लिए मैं सदा ((आभारी रहूँगा|आभारी रहूँगी)) और आशा ((करता|करती)) हूँ कि कभी मैं भी आपके काम आ सकूँ।"]]}
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
  var S = Object.assign({ type: 'sick-leave', lang: 'en', lay: 'in', tone: 'formal', cur: '₹', lday: L.today() }, store('tk-letter-s') || {});
  if (S.lday < L.today()) S.lday = L.today();
  var M = store('tk-letter-m') || {}, D = store('tk-letter-d') || {};
  (function () { var q = new URLSearchParams(location.search).get('type'); if (q && BY[q]) S.type = q; })();
  if (!BY[S.type]) S.type = TYPES[0].id;

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
  function sample(t, k, hi) { var s = t.sm && t.sm[k]; if (s) return s[hi ? 1 : 0]; return F[k] ? F[k][hi ? 3 : 2] : ''; }

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
  function filler(t, hi, lay) {
    function val(k) {
      var raw, def = F[k];
      if (k === 'ndays') { var n = daysBetween(getv(t, 'd1'), getv(t, 'd2')); return n ? { v: hi ? n + ' दिन' : n + (n === 1 ? ' day' : ' days') } : { v: '', e: 1 }; }
      if (k === 'amtw') {
        if (hi || S.cur !== '₹') return { v: '', e: 1 };
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
        case 's': var o = def[4][+raw] || def[4][0]; return { v: o[hi ? 1 : 0] };
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
      s = s.replace(/\{(\w+)\}/g, function (_, k) { return val(k).v; });
      return s;
    };
  }

  function model() {
    var t = BY[S.type], R = role(t), hi = S.lang === 'hi', lay = hi ? 'in' : S.lay, li = hi ? 1 : 0, fill = filler(t, hi, lay);
    var B = [], dr = draft(t);
    function add(text, o) { if (text) B.push(Object.assign({ t: text }, o || {})); }
    function g(k) { var v = getv(t, k); return v == null ? '' : String(v).trim(); }
    function lines(v) { return String(v || '').split(/\n+/).map(function (s) { return s.trim(); }).filter(Boolean); }
    var female = g('sg') === 'f', named = (g('rsal') || t.rs || 'sm') === 'n';
    var subj = g('subj') || fill(t.s[li]);
    /* letterhead or sender block */
    if (R.head) {
      add(g('sorg') || '[' + F.sorg[0] + ']', { al: 'c', b: 1, big: 1 });
      add([lines(g('saddr')).join(', '), [g('sphone') && (hi ? 'फ़ोन: ' : 'Phone: ') + g('sphone'), g('semail')].filter(Boolean).join('  ·  ')].filter(Boolean).join('\n'), { al: 'c', sm: 1, rule: 1 });
    } else if (lay !== 'in') {
      var sb = lines(g('saddr')).concat([g('sphone'), g('semail')].filter(Boolean));
      if (sb.length) add(sb.join('\n'), { al: lay === 'uk' ? 'r' : 'l' });
    }
    var dateTxt = fmtDate(S.lday, lay);
    if (lay === 'in') add([R.head && g('ref') ? (hi ? 'पत्रांक: ' : 'Ref: ') + g('ref') : '', (hi ? 'दिनांक: ' : 'Date: ') + dateTxt].filter(Boolean).join('\n'));
    else { if (R.head && g('ref')) add('Ref: ' + g('ref')); add(dateTxt, { al: lay === 'uk' && !R.head ? 'r' : 'l' }); }
    /* recipient */
    if (!t.tw) {
      var to = [g('rto') || fill(t.to[li]), g('rorg') || fill(t.org[li]), g('raddr') ? lines(g('raddr')).join('\n') : (hi ? '[शहर]' : '[City]')].filter(Boolean);
      if (named && g('rname') && to[0] !== g('rname') && !g('rto')) to.unshift(g('rname'));
      add((lay === 'in' ? (hi ? 'सेवा में,\n' : 'To,\n') : '') + to.join('\n'));
    }
    /* subject and salutation */
    var sal = '';
    if (!t.tw) {
      if (t.sal) sal = fill(t.sal[li]);
      else if (hi) sal = named ? (t.cl === 'warm' ? 'प्रिय ' : 'आदरणीय ') + (g('rname') || '[' + F.rname[0] + ']') + ' जी' : ({ s: 'महोदय', m: 'महोदया' }[g('rsal')] || 'महोदय/महोदया');
      else {
        var base = R.resp && lay === 'in' ? 'Respected ' : 'Dear ', rs = g('rsal') || t.rs || 'sm';
        sal = rs === 'n' ? 'Dear ' + (g('rname') || '[' + F.rname[0].replace(/ \(.*/, '') + ']') : rs === 's' ? base + 'Sir' : rs === 'm' ? base + 'Madam' : (lay === 'in' ? base + 'Sir/Madam' : 'Dear Sir or Madam');
      }
      sal += lay === 'us' && t.cl !== 'warm' && !hi ? ':' : ',';
    }
    if (t.tw) {
      add(hi ? subj : subj.toUpperCase(), { al: 'c', b: 1, u: 1 });
      add(hi ? 'सर्वसंबंधित के लिए' : 'TO WHOM IT MAY CONCERN', { al: 'c', b: 1 });
    } else if (lay === 'uk') { add(sal); add(subj, { b: 1 }); }
    else { add((hi ? 'विषय: ' : 'Subject: ') + subj, { b: 1 }); add(sal); }
    /* body, chosen sentences, tone */
    var paras = t.b[li].map(fill);
    var opts = (t.o || []).filter(function (o) { var v = dr['o_' + o[0]]; return v == null ? o[4] : v === 1; }).map(function (o) { return fill(o[li + 2]); }).join(' ');
    if (opts) paras.splice(t.oa != null ? t.oa : paras.length, 0, opts);
    var tone = (t.t && t.t[S.tone]) || (TONE[t.k] && TONE[t.k][S.tone]) || null;
    if (tone) paras.push(fill(tone[li]));
    paras.forEach(function (p) { p = p.replace(/([^.])\.\.(?!\.)/g, '$1.').replace(/\s+([,.।])/g, '$1').replace(/ {2,}/g, ' ').trim(); if (p) add(hi ? p : p.charAt(0).toUpperCase() + p.slice(1)); });
    if (lay === 'in' && !R.head && t.cl !== 'warm' && ['req', 'comp', 'info'].indexOf(t.k) >= 0) add(hi ? 'सधन्यवाद,' : 'Thanking you,');
    /* closing and signature */
    var cl = R.cl, close;
    if (hi) close = cl === 'obed' ? (female ? 'आपकी आज्ञाकारिणी शिष्या' : 'आपका आज्ञाकारी शिष्य') : cl === 'warm' ? 'सादर,' : cl === 'app' && (t.k === 'req' || t.k === 'comp') ? (female ? 'प्रार्थिनी' : 'प्रार्थी') : (female ? 'भवदीया' : 'भवदीय');
    else close = cl === 'warm' ? (lay === 'us' ? 'Warm regards,' : 'With warm regards,') : lay === 'us' ? 'Sincerely,' : cl === 'obed' && lay === 'in' ? 'Yours obediently,' : named || t.sal ? 'Yours sincerely,' : 'Yours faithfully,';
    var sig = [];
    if (!t.tw) sig.push(close);
    if (R.head) sig.push(hi ? (g('sorg') || '[संस्था]') + ' की ओर से' : 'For ' + (g('sorg') || '[' + F.sorg[0] + ']'));
    sig.push('', '', g('sname') || '[' + F.sname[0] + ']');
    var LBL = { cls: ['Class: ', 'कक्षा: '], roll: ['Roll No.: ', 'अनुक्रमांक: '], dept: ['Department: ', 'विभाग: '], empid: ['Employee ID: ', 'कर्मचारी कोड: '], sphone: ['Mobile: ', 'मोबाइल: '], semail: ['Email: ', 'ईमेल: '], acno: ['A/c No.: ', 'खाता संख्या: '], flat: ['Flat: ', 'फ्लैट: '] };
    R.x.forEach(function (k) {
      if (k === 'saddr') { if (lay === 'in') sig = sig.concat(lines(g('saddr'))); return; }
      if (k === '_parent') { if (g('child')) sig.push(hi ? g('child') + (g('cls') ? ', कक्षा ' + g('cls') : '') + ' के अभिभावक' : 'Parent of ' + g('child') + (g('cls') ? ', Class ' + g('cls') : '')); return; }
      if ((k === 'sphone' || k === 'semail') && lay !== 'in' && !R.head) return;
      var v = g(k); if (!v) return;
      if (k === 'actype') return;
      sig.push((LBL[k] ? LBL[k][li] : '') + v);
    });
    if (t.f.indexOf('acno') >= 0 && R.x.indexOf('acno') < 0 && g('acno') && t.c === 'bank') sig.push(LBL.acno[li] + g('acno'));
    add(sig.join('\n'), { sig: 1 });
    /* enclosures and copy to */
    var enc = lines(g('enc'));
    if (enc.length) add((hi ? 'संलग्न:\n' : 'Enclosures:\n') + enc.map(function (s, i) { return (i + 1) + '. ' + s; }).join('\n'), { sm: 1 });
    var cc = lines(g('cc'));
    if (cc.length) add((hi ? 'प्रतिलिपि:\n' : 'Copy to:\n') + cc.map(function (s, i) { return (i + 1) + '. ' + s; }).join('\n'), { sm: 1 });
    return { blocks: B, subj: subj, hi: hi, lay: lay, name: t.n };
  }

  /* hand edits: the edited text replaces the blocks; paragraphs that are unchanged keep their formatting */
  function current() {
    var m = model(), ed = draft(BY[S.type])._edit;
    if (ed == null) return m;
    var fmt = {}; m.blocks.forEach(function (b) { fmt[b.t] = b; });
    m.blocks = String(ed).replace(/\r/g, '').split(/\n[ \t]*\n/).map(function (p) { return p.replace(/^\n+|\n+$/g, ''); }).filter(function (p) { return p.trim(); })
      .map(function (p) { return Object.assign({}, fmt[p] || (/^(Subject|Sub|Re|RE|विषय)\s*:/.test(p) ? { b: 1 } : {}), { t: p }); });
    m.edited = true;
    return m;
  }
  function toText(m) { return m.blocks.map(function (b) { return b.t; }).join('\n\n'); }

  /* ---------------- outputs */
  function safe(s) { return String(s || 'letter').replace(/[^\w\- ]+/g, '').trim().replace(/\s+/g, '-').slice(0, 60) || 'letter'; }
  function fname(m) { return safe(m.name) + '-' + S.lday; }
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
      var font = m.hi ? 'Nirmala UI' : 'Calibri', AL = { l: X.AlignmentType.LEFT, r: X.AlignmentType.RIGHT, c: X.AlignmentType.CENTER };
      var kids = m.blocks.map(function (b) {
        var ls = b.t.split('\n');
        return new X.Paragraph({ alignment: AL[b.al || 'l'], spacing: { after: b.sig ? 120 : 200, line: 276 },
          border: b.rule ? { bottom: { color: '2563EB', space: 6, style: X.BorderStyle.SINGLE, size: 8 } } : undefined,
          children: ls.map(function (s, i) { return new X.TextRun({ text: s, break: i ? 1 : 0, bold: !!b.b, underline: b.u ? {} : undefined, size: b.big ? 32 : b.sm ? 20 : 22, font: font, color: b.big ? '1D4ED8' : undefined }); }) });
      });
      var us = m.lay === 'us';
      var d = new X.Document({ creator: 'Nivesh Sahayak', lastModifiedBy: 'Nivesh Sahayak', keywords: 'Toolkit by GroupChatOn; Nivesh Sahayak; https://www.niveshsahayak.in/',
        description: CREDIT + ' · https://www.niveshsahayak.in/', title: m.subj,
        styles: { default: { document: { run: { font: font, size: 22 } } } },
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
    '.lw-paper{font:14px/1.55 Calibri,Carlito,"Segoe UI",Arial,sans-serif;color:#14213D;overflow-wrap:anywhere}.lw-paper.hi{font-family:"Nirmala UI",Mangal,"Noto Sans Devanagari","Segoe UI",sans-serif}' +
    '.lw-paper p{margin:0 0 12px;white-space:pre-line}.lw-paper .sig{margin-top:4px}.lw-paper .big{font-size:20px;color:#1D4ED8;margin:0 0 2px}.lw-paper .sm{font-size:12.5px}' +
    '.lw-paper .rule{border-bottom:2px solid #2563EB;padding-bottom:8px;margin-bottom:16px}.lw-paper mark{background:#FEF08A;color:inherit;border-radius:3px;padding:0 2px}' +
    '.lw-made{color:#94A3B8;font-size:11px;text-align:center;margin:18px 0 0}' +
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
      var wrap = h('div', { 'class': 'lw-paper' + (m.hi ? ' hi' : ''), lang: m.hi ? 'hi' : 'en' });
      m.blocks.forEach(function (b) {
        var p = h('p', { 'class': [b.sig ? 'sig' : '', b.big ? 'big' : '', b.sm ? 'sm' : '', b.rule ? 'rule' : ''].join(' ').trim() || null,
          style: [b.al === 'c' ? 'text-align:center' : b.al === 'r' ? 'text-align:right' : '', b.b ? 'font-weight:700' : '', b.u ? 'text-decoration:underline' : ''].filter(Boolean).join(';') || null });
        b.t.split(/(\[[^\[\]\n]{1,60}\])/).forEach(function (part, i) { p.appendChild(i % 2 ? h('mark', { text: part }) : doc.createTextNode(part)); });
        wrap.appendChild(p);
      });
      paper.appendChild(wrap);
      paper.appendChild(h('p', { 'class': 'no-print lw-made', text: MADE + ' · ' + CREDIT }));   /* on screen only: the credit is never printed or put in the letter */
      banner.hidden = !m.edited;
      var miss = (toText(m).match(/\[[^\[\]\n]{1,60}\]/g) || []).length;
      msg.textContent = miss ? miss + ' detail' + (miss === 1 ? '' : 's') + ' still to fill (highlighted in yellow).' : 'Letter ready: ' + t.n + '.';
    }
    function update() { clearTimeout(timer); timer = setTimeout(function () { persist(); render(); }, 120); }

    function field(t, k) {
      var def = F[k] || [k, 't'], type = def[1], hi = S.lang === 'hi', v = getv(t, k), inp, lab = label(t, k);
      var smp = sample(t, k, hi);
      function on(val) { setv(t, k, val); if (draft(t)._edit != null) { msg.textContent = 'Your hand edits are kept. Press "Undo my edits" to see this change.'; } update(); }
      if (type === 'g' || type === 'k' || type === 's') {
        var opts = type === 'g' ? [['m', 'Male'], ['f', 'Female']] : type === 'k' ? [['m', 'Son'], ['f', 'Daughter']] : def[4].map(function (o, i) { return [String(i), hi ? o[0] + ' · ' + o[1] : o[0]]; });
        var cur = v == null || v === '' ? def[2] : String(v);
        inp = h('select', { onchange: function () { on(inp.value); } }, opts.map(function (o) { return h('option', { value: o[0], selected: o[0] === cur ? true : null, text: o[1] }); }));
      } else if (type === 'a' || type === 'l') {
        inp = h('textarea', { rows: type === 'l' ? 4 : 3, placeholder: smp ? 'e.g. ' + smp.split('\n')[0] : null, oninput: function () { on(inp.value); } }); inp.value = v || '';
      } else {
        inp = h('input', { type: type === 'd' ? 'date' : type === 'e' ? 'email' : type === 'p' ? 'tel' : 'text', value: v || '', placeholder: type === 'd' ? null : smp ? 'e.g. ' + smp : null,
          inputmode: type === 'm' || type === 'n' ? 'decimal' : null, autocomplete: { sname: 'name', saddr: 'street-address', sphone: 'tel', semail: 'email', sorg: 'organization' }[k] || null,
          oninput: function () { on(inp.value); } });
      }
      var wide = type === 'a' || type === 'l' || ['reason', 'issue', 'what', 'details', 'ach', 'skills', 'quality', 'task', 'fix', 'offer', 'offer2', 'occasion', 'help', 'ask', 'purpose', 'item', 'problem', 'qual', 'event', 'venue', 'prel'].indexOf(k) >= 0;
      return h('label', { 'class': 'tk-f' + (wide ? ' wide' : ''), style: wide ? 'grid-column:1/-1' : null }, [lab, inp]);
    }

    function pickBox() {
      var t = T(), det = h('details', { 'class': 'lw-pick' });
      det.appendChild(h('summary', {}, [h('span', { 'class': 'lw-ic', 'aria-hidden': 'true', text: CAT[t.c][2] }), h('span', {}, [t.n, h('small', { text: CAT[t.c][1] + ' · ' + t.h })])]));
      var q = h('input', { type: 'search', 'class': 'lw-search', placeholder: 'Search ' + TYPES.length + ' letters: leave, fee, bank, RTI, resignation…', 'aria-label': 'Search letter types' });
      var list = h('div', { 'class': 'lw-list' });
      function draw() {
        var words = q.value.toLowerCase().split(/\s+/).filter(Boolean); list.innerHTML = ''; var n = 0;
        CATS.forEach(function (c) {
          var ts = TYPES.filter(function (x) { if (x.c !== c[0]) return false; var hay = (x.n + ' ' + x.h + ' ' + (x.kw || '') + ' ' + c[1]).toLowerCase(); return words.every(function (w) { return hay.indexOf(w) >= 0; }); });
          if (!ts.length) return; n += ts.length;
          list.appendChild(h('h4', { text: c[2] + ' ' + c[1] }));
          ts.forEach(function (x) {
            list.appendChild(h('button', { type: 'button', 'class': 'lw-t', 'aria-current': x.id === S.type ? 'true' : null, onclick: function () { choose(x.id); } }, [x.n, h('small', { text: x.h })]));
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
      try { history.replaceState(null, '', location.pathname + '?type=' + id); } catch (e) {}
      var f = form.querySelector('fieldset:nth-of-type(2)'); if (f && window.innerWidth < 960) f.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function sel(lab, key, opts, onc) {
      var s = h('select', { onchange: function () { S[key] = s.value; if (onc) onc(); persist(); build(); render(); } },
        opts.map(function (o) { return h('option', { value: o[0], selected: S[key] === o[0] ? true : null, text: o[1] }); }));
      return [h('label', { 'class': 'tk-f' }, [lab, s]), s];
    }

    function build() {
      var t = T(), R = role(t), hi = S.lang === 'hi';
      form.innerHTML = '';
      form.appendChild(h('fieldset', {}, [h('legend', { text: '1. What is the letter about?' }), pickBox()]));
      var la = sel('Language', 'lang', [['en', 'English'], ['hi', 'हिन्दी (Hindi)']]);
      var lo = sel('Layout', 'lay', [['in', 'Indian'], ['uk', 'UK / international'], ['us', 'US']], function () { S.cur = S.lay === 'uk' ? '£' : S.lay === 'us' ? '$' : '₹'; });
      if (hi) { lo[1].disabled = true; lo[0].title = 'Hindi letters use the Indian layout'; }
      var to = sel('Tone', 'tone', [['formal', 'Formal'], ['polite', 'Polite and warm'], ['firm', 'Firm']]);
      var cu = sel('Currency', 'cur', CURS);
      var dt = h('input', { type: 'date', value: S.lday, oninput: function () { S.lday = dt.value || L.today(); update(); } });
      form.appendChild(h('fieldset', {}, [h('legend', { text: '2. Language, layout and tone' }), h('div', { 'class': 'tk-row' }, [la[0], lo[0], to[0], cu[0], h('label', { 'class': 'tk-f' }, ['Letter date', dt])]),
        hi ? h('p', { 'class': 'lw-note', text: 'Type names, reasons and places in Hindi too, so the whole letter reads in one language. Hindi letters use the Indian layout.' }) : null]));
      form.appendChild(h('fieldset', {}, [h('legend', { text: '3. Your details (saved on this device)' }), h('div', { 'class': 'tk-row' }, R.p.map(function (k) { return field(t, k); }))]));
      var rs = t.sal || t.tw ? null : (function () {
        var cur = getv(t, 'rsal') || t.rs || 'sm', s = h('select', { onchange: function () { setv(t, 'rsal', s.value); persist(); build(); render(); } },
          SAL.map(function (o) { return h('option', { value: o[0], selected: o[0] === cur ? true : null, text: o[1] }); }));
        return h('label', { 'class': 'tk-f' }, ['Salutation', s]);
      })();
      if (!t.tw) {
        var named = (getv(t, 'rsal') || t.rs || 'sm') === 'n';
        var rto = field(t, 'rto'), rorg = field(t, 'rorg'), rad = field(t, 'raddr');
        rto.firstChild.textContent = 'To (name or designation)'; rto.lastChild.placeholder = t.to[hi ? 1 : 0].replace(/\{\w+\}/g, '').trim() || 'e.g. ' + sample(t, 'rname', hi);
        rorg.firstChild.textContent = 'Organisation / office'; rorg.lastChild.placeholder = t.org[hi ? 1 : 0].replace(/\{\w+\}/g, '') || '(optional)';
        rad.firstChild.textContent = 'Address / city';
        form.appendChild(h('fieldset', {}, [h('legend', { text: '4. Sent to' }), h('div', { 'class': 'tk-row' }, [rs, named || t.sal ? field(t, 'rname') : null, rto, rorg, rad]),
          h('p', { 'class': 'lw-note', text: 'Leave "To" empty to use the usual designation shown in grey.' })]));
      }
      var fk = []; R.a.concat(t.f).forEach(function (k) { if (fk.indexOf(k) < 0 && R.p.indexOf(k) < 0) fk.push(k); });
      var about = [h('legend', { text: (t.tw ? '4' : '5') + '. About this letter' }), h('div', { 'class': 'tk-row' }, fk.map(function (k) { return field(t, k); }))];
      if (t.o && t.o.length) {
        var dr = draft(t);
        about.push(h('div', { 'class': 'lw-opts' }, t.o.map(function (o) {
          var on = dr['o_' + o[0]] == null ? !!o[4] : dr['o_' + o[0]] === 1;
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
      var t = T(), R = role(t), hi = S.lang === 'hi', c = CAT[t.c];
      R.p.concat(R.a, t.f).forEach(function (k) {
        if (k === 'ref' && !R.head) return;
        var def = F[k]; if (!def) return;
        if (PERSIST[k] && getv(t, k)) return;
        if (def[1] === 'd') {
          var past = k === 'd1' && /lost|theft|incident|received|purchase|closed|moved|order date|transaction/i.test(label(t, k));
          var d = new Date(); d.setDate(d.getDate() + (past ? -7 : { d1: 0, d2: 2, ldate: 30, by: 7, due: -15, invdate: -45, chqdate: -2, jdate: 0, offerdate: -10, rdate: -30, edd: 60, apptdate: 5 }[k] || 0));
          setv(t, k, d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2)); return; }
        setv(t, k, sample(t, k, hi));
      });
      if (!t.tw) {
        var so = t.so || (c[3][0] ? [c[3][0], c[3][1]] : null);
        if (so && !getv(t, 'rorg') && t.org[0]) setv(t, 'rorg', so[hi ? 1 : 0]);
        if (!getv(t, 'raddr')) setv(t, 'raddr', c[3][hi ? 3 : 2]);
        if ((getv(t, 'rsal') || t.rs) === 'n' && !getv(t, 'rname')) setv(t, 'rname', sample(t, 'rname', hi));
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
    function busy(b, text, job) { finishEdit(); b.disabled = true; msg.textContent = text; Promise.resolve().then(job).then(function (r) { if (r) msg.textContent = r; }, function (e) { msg.textContent = (e && e.message) || 'Something went wrong. Please try again.'; }).then(function () { b.disabled = false; }); }
    out.appendChild(h('div', { 'class': 'tk-out-acts' }, [
      h('button', { type: 'button', 'class': 'btn', onclick: function (e) { busy(e.currentTarget, 'Copying…', function () { return copy(toText(current())).then(function () { return 'Copied. Paste it into an email, WhatsApp or Word.'; }); }); } }, [icon('📋'), ' Copy']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function () { finishEdit(); persist(); window.print(); } }, [icon('🖨'), ' Print / PDF']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function (e) { busy(e.currentTarget, 'Making the Word file…', function () { return toDocx(current()).then(function () { return 'Word file ready. Check your downloads folder.'; }); }); } }, [icon('📘'), ' Word']),
      h('button', { type: 'button', 'class': 'btn btn-2', onclick: function (e) { busy(e.currentTarget, '', function () { var m = current(); save(new Blob(['﻿' + toText(m).replace(/\n/g, '\r\n')], { type: 'text/plain;charset=utf-8' }), fname(m) + '.txt'); return 'Text file ready. Check your downloads folder.'; }); } }, [icon('📄'), ' Text']),
      editBtn]));
    banner.appendChild(h('span', { text: 'You changed the text by hand, so changes in the form do not show here.' }));
    banner.appendChild(h('button', { type: 'button', 'class': 'btn btn-2', style: 'min-height:36px;padding:6px 12px', text: 'Undo my edits', onclick: function () { delete draft(T())._edit; editing = false; editor.hidden = true; paper.hidden = false; editBtn.lastChild.textContent = ' Edit text'; persist(); render(); } }));
    out.appendChild(msg); out.appendChild(banner); out.appendChild(editor); out.appendChild(paper);
    build();
    spot.appendChild(h('div', { 'class': 'tk-app' }, [form, out]));
    render();
    doc.body.classList.add('pg-tool');
  }

  /* register with the toolkit engine, but keep its table form away from this box: it mounts .tk-tool[data-tool] on DOMContentLoaded */
  TK.register('letter-writer', { name: 'letter', custom: true, mount: mount, types: TYPES.length });
  var boxes = Array.prototype.slice.call(doc.querySelectorAll('.tk-tool[data-tool="letter-writer"]'));
  boxes.forEach(function (b) { b.removeAttribute('data-tool'); b.setAttribute('data-lw', ''); });
  function go() { boxes.forEach(mount); }
  if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', go); else go();
})();
