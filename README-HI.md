# Boiler Construction Management System

यह Vite + React + Supabase आधारित शुरुआती वेब ऐप है। इसमें Login, DPR register, Excel import/export, monthly DPR totals, RA monthly summary, Customer Bills और private Drawing uploads शामिल हैं।

## 1. Supabase credentials
Supabase Dashboard > Project > Connect (या Settings > API) से:
- Project URL: पूरा URL, जैसे `https://xxxxxxxx.supabase.co`
- Publishable Key: `sb_publishable_...` (पुराने प्रोजेक्ट में `anon` public key भी हो सकती है)

`.env.example` की कॉपी `.env` नाम से बनाएं और दोनों सही values भरें। Project reference अकेला URL नहीं है और Region key नहीं है। `service_role` / Secret key कभी `.env` frontend में न डालें।

## 2. Database बनाएं
Supabase > SQL Editor > New query में `supabase/schema.sql` की पूरी सामग्री paste करके Run करें।

## 3. Admin login बनाएं
Supabase > Authentication > Users > Add user में अपना admin email और एक मजबूत password बनाएं। App उसी email/password से login करेगा। `Boiler03` जैसा छोटा password production में इस्तेमाल न करें; बेहतर है 12+ अक्षरों का अलग password रखें और किसी से साझा न करें।

## 4. Local run
Node.js 20+ install करें, इस folder में terminal खोलें:
```bash
npm install
npm run dev
```
दिखाया गया localhost URL खोलें।

## 5. Online deployment
इस folder को GitHub repository में डालें और Vercel/Netlify पर import करें। Build command `npm run build`, output folder `dist` रखें। Environment variables में `VITE_SUPABASE_URL` और `VITE_SUPABASE_ANON_KEY` जोड़ें। Deploy के बाद मिले URL को मोबाइल/कंप्यूटर में share कर सकते हैं।

## 6. Sharing / team access
इस starter में user-owned rows हैं: प्रत्येक user केवल अपने records देखता/बदलता है। Team sharing के लिए अगला चरण project/workspace membership और role-based RLS (Admin/Editor/Viewer) जोड़ना है। अभी किसी दूसरे login को link भेजने से उसे आपका data अपने आप नहीं दिखेगा।

## Important limitations
- DPR Excel import पहली worksheet के rows पढ़ता है; column headers database field names के समान होने चाहिए (work_date, department, pg, ma, du, description, quantity, weight, pre_assembly, placement, alignment, bolting, remark).
- Day name display होता है और date weekday browser से निकाला जाता है।
- Remark auto-suggestion केवल form में completed work quantities के आधार पर दिखती है; save करते समय उसे Remark field में manually डालना है।
- RA monthly summary फिलहाल DPR से quantity/weight का snapshot save करता है। Approved item-wise contract rates, previous RA deductions, GST/TDS/retention और customer bill formats मिलने पर final bill calculations जोड़ें।
- Excel workbook templates with 12 monthly tabs and seven department tabs are a follow-up enhancement; current app supports date/month filtering and department field.
