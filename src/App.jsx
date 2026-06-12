import { useState, useEffect, useRef } from "react";
import { AreaChart, Area, BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, Tooltip, ResponsiveContainer } from "recharts";
import * as XLSX from "xlsx";
import { loadData, saveData } from "./firebase.js";

/* ═══ TOKENS ════════════════════════════════════════════════════ */
const T = {
  bg:"#F0F4F8", surface:"#FFFFFF", hi:"#F8FAFC", hi2:"#F1F5F9",
  border:"#E2E8F0", borderHi:"#CBD5E1",
  saffron:"#F59E0B", saffronDk:"#D97706", saffronLt:"#FEF3C7", saffronXLt:"#FFFBEB",
  green:"#059669", greenLt:"#D1FAE5", greenMd:"#34D399",
  red:"#DC2626", redLt:"#FEE2E2",
  blue:"#2563EB", blueLt:"#DBEAFE",
  purple:"#7C3AED", purpleLt:"#EDE9FE",
  teal:"#0D9488", tealLt:"#CCFBF1",
  text:"#0F172A", textMd:"#475569", textSm:"#94A3B8",
  shadow:"0 1px 3px rgba(0,0,0,0.08),0 1px 2px rgba(0,0,0,0.05)",
  shadowMd:"0 4px 12px rgba(0,0,0,0.08),0 2px 4px rgba(0,0,0,0.05)",
  shadowLg:"0 10px 25px rgba(0,0,0,0.1),0 4px 6px rgba(0,0,0,0.04)",
};
const PIE_COLORS=[T.saffron,T.blue,T.green,T.purple,T.teal,"#F472B6",T.red];

/* ═══ TRANSLATIONS ══════════════════════════════════════════════ */
const TR = {
  en:{appName:"Kesri Nandan Spares",appSub:"Advanced AI Manager",home:"Home",analytics:"Analytics",pos:"Sales",stock:"Stock",finance:"Finance",tools:"Tools",revenue:"Revenue",profit:"Profit",pending:"Pending",margin:"Margin",cashInHand:"Cash in Hand",bankBalance:"Bank Balance",totalParts:"Total Parts",supplierDues:"Supplier Dues",recentTx:"Recent Transactions",lowStockAlert:"Low Stock Alert",quickSale:"Quick Sale",salesLog:"Sales Log",gstInvoice:"GST Invoice",purchaseOrder:"Purchase Order",dailyClosing:"Daily Closing",addPart:"Add Part",searchParts:"Search parts, SKU or bike...",ledger:"Ledger",plReport:"P&L Report",upi:"UPI",whatsapp:"WhatsApp AI",reminders:"Reminders",customers:"Customers",loyalty:"Loyalty",staff:"Staff",exportData:"Export Data",save:"Save",markPaid:"Mark Paid",checkout:"Checkout",addEntry:"Add Entry",addReminder:"Add Reminder",addCustomer:"Add Customer",addStaff:"Add Staff",cash:"Cash",bank:"Bank",moneyIn:"Money In",moneyOut:"Money Out",today:"Today",print:"Print Invoice",copyWA:"Copy for WhatsApp",exportExcel:"Export to Excel",netProfit:"Net Profit",grossRevenue:"Gross Revenue",cogs:"Cost of Goods",grossProfit:"Gross Profit",totalExp:"Total Expenses",rent:"Shop Rent",electricity:"Electricity",salary:"Staff Salary",misc:"Miscellaneous",stockValue:"Stock Value",bestSellers:"Top 5 Best Sellers",dailySales:"Daily Sales",loyaltyPoints:"Loyalty Points",totalSpend:"Total Spend",rewardReady:"Reward Ready!",staffManager:"Staff Manager",present:"Present",absent:"Absent",dailyReport:"Daily Closing Report",openingBal:"Opening Balance",closingBal:"Closing Balance",totalSales:"Total Sales",partName:"Part Name",category:"Category",bikeModel:"Bike Model",buyPrice:"Buy Price",sellPrice:"Sell Price",qty:"Qty",minAlert:"Min Alert",subtitle:"Subtitle",sku:"SKU Code",customerName:"Customer Name",phone:"Phone",area:"Area",gstin:"GSTIN",amount:"Amount",note:"Note",time:"Time",repeat:"Repeat",gridView:"Grid",listView:"List",lowStock:"Low Stock",newSale:"New Sale",saleComplete:"Sale Complete!",payMode:"Payment Mode",searchPlaceholder:"Search parts, customers, suppliers...",editStock:"Edit Stock",stockIn:"Stock In",stockOut:"Stock Out",updatePrice:"Update Price",editPart:"Edit Part"},
  hi:{appName:"केसरी नंदन स्पेयर्स",appSub:"AI बिज़नेस मैनेजर",home:"होम",analytics:"विश्लेषण",pos:"बिक्री",stock:"स्टॉक",finance:"वित्त",tools:"टूल्स",revenue:"राजस्व",profit:"मुनाफा",pending:"बकाया",margin:"मार्जिन",cashInHand:"नकद",bankBalance:"बैंक",totalParts:"कुल पार्ट्स",supplierDues:"सप्लायर बकाया",recentTx:"हाल के लेनदेन",lowStockAlert:"कम स्टॉक",quickSale:"त्वरित बिक्री",salesLog:"बिक्री",gstInvoice:"GST बिल",purchaseOrder:"खरीद आर्डर",dailyClosing:"दैनिक बंदी",addPart:"पार्ट जोड़ें",searchParts:"पार्ट खोजें...",ledger:"खाता बही",plReport:"लाभ-हानि",upi:"UPI",whatsapp:"WhatsApp AI",reminders:"रिमाइंडर",customers:"ग्राहक",loyalty:"पॉइंट्स",staff:"स्टाफ",exportData:"निर्यात",save:"सहेजें",markPaid:"भुगतान मिला",checkout:"चेकआउट",addEntry:"एंट्री जोड़ें",addReminder:"रिमाइंडर जोड़ें",addCustomer:"ग्राहक जोड़ें",addStaff:"स्टाफ जोड़ें",cash:"नकद",bank:"बैंक",moneyIn:"पैसा आया",moneyOut:"पैसा गया",today:"आज",print:"प्रिंट करें",copyWA:"WhatsApp कॉपी",exportExcel:"Excel निर्यात",netProfit:"शुद्ध लाभ",grossRevenue:"सकल राजस्व",cogs:"माल लागत",grossProfit:"सकल लाभ",totalExp:"कुल खर्च",rent:"किराया",electricity:"बिजली",salary:"वेतन",misc:"विविध",stockValue:"स्टॉक मूल्य",bestSellers:"टॉप 5",dailySales:"दैनिक बिक्री",loyaltyPoints:"पॉइंट्स",totalSpend:"कुल खरीद",rewardReady:"इनाम तैयार!",staffManager:"स्टाफ",present:"उपस्थित",absent:"अनुपस्थित",dailyReport:"दैनिक रिपोर्ट",openingBal:"शुरुआती",closingBal:"अंतिम",totalSales:"कुल बिक्री",partName:"पार्ट नाम",category:"श्रेणी",bikeModel:"बाइक",buyPrice:"खरीद मूल्य",sellPrice:"बिक्री मूल्य",qty:"मात्रा",minAlert:"न्यूनतम",subtitle:"विवरण",sku:"SKU",customerName:"ग्राहक",phone:"फोन",area:"इलाका",gstin:"GSTIN",amount:"राशि",note:"नोट",time:"समय",repeat:"दोहराएं",gridView:"ग्रिड",listView:"सूची",lowStock:"कम स्टॉक",newSale:"नई बिक्री",saleComplete:"बिक्री पूरी!",payMode:"भुगतान",searchPlaceholder:"पार्ट, ग्राहक खोजें...",editStock:"स्टॉक अपडेट",stockIn:"स्टॉक आया",stockOut:"स्टॉक गया",updatePrice:"कीमत बदलें",editPart:"पार्ट संपादित करें"}
};

/* ═══ INITIAL DATA ══════════════════════════════════════════════ */
const INIT_STOCK=[
  {id:1,name:"Brake Pad",sub:"Honda Shine 125",cat:"Brakes",qty:45,minQty:10,buy:85,sell:140,bike:"Honda Shine",photo:null,sold:38,sku:"BP-HS-001"},
  {id:2,name:"Chain Sprocket Kit",sub:"Hero Splendor",cat:"Drive",qty:8,minQty:5,buy:320,sell:520,bike:"Splendor",photo:null,sold:21,sku:"CS-SP-002"},
  {id:3,name:"Air Filter",sub:"Bajaj Pulsar 150",cat:"Engine",qty:3,minQty:5,buy:110,sell:190,bike:"Pulsar 150",photo:null,sold:17,sku:"AF-PU-003"},
  {id:4,name:"Engine Oil 15W40",sub:"1L Universal",cat:"Lubricants",qty:60,minQty:20,buy:180,sell:260,bike:"Universal",photo:null,sold:54,sku:"EO-UV-004"},
  {id:5,name:"Clutch Cable",sub:"Hero Splendor Plus",cat:"Cables",qty:22,minQty:8,buy:45,sell:85,bike:"Splendor",photo:null,sold:29,sku:"CC-SP-005"},
  {id:6,name:"Headlight Bulb 35W",sub:"Universal Fit",cat:"Electricals",qty:2,minQty:10,buy:60,sell:110,bike:"Universal",photo:null,sold:12,sku:"HB-UV-006"},
  {id:7,name:"Disc Brake Pad",sub:"Bajaj Pulsar NS200",cat:"Brakes",qty:14,minQty:6,buy:180,sell:320,bike:"Pulsar NS200",photo:null,sold:9,sku:"DB-PU-007"},
  {id:8,name:"Spark Plug NGK",sub:"Universal",cat:"Engine",qty:35,minQty:15,buy:55,sell:90,bike:"Universal",photo:null,sold:42,sku:"SP-UV-008"},
];
const INIT_SALES=[
  {id:1,date:"2026-06-07",item:"Air Filter",qty:4,total:760,profit:320,customer:"Walk-in",paid:true,upi:true},
  {id:2,date:"2026-06-05",item:"Clutch Cable",qty:8,total:680,profit:320,customer:"Suresh Yadav",paid:false,upi:false},
  {id:3,date:"2026-06-03",item:"Chain Sprocket Kit",qty:3,total:1560,profit:600,customer:"Walk-in",paid:true,upi:false},
  {id:4,date:"2026-06-02",item:"Engine Oil 15W40",qty:10,total:2600,profit:800,customer:"Ajay Singh",paid:true,upi:true},
  {id:5,date:"2026-06-01",item:"Brake Pad",qty:5,total:700,profit:275,customer:"Ramesh Kumar",paid:true,upi:false},
];
const INIT_CUSTOMERS=[
  {id:1,name:"Ramesh Kumar",phone:"9876543210",bike:"Honda Shine 125",area:"Sector 12",totalBuy:4200,visits:6,points:42,gstin:""},
  {id:2,name:"Suresh Yadav",phone:"9812345678",bike:"Hero Splendor Plus",area:"Gandhi Nagar",totalBuy:1850,visits:3,points:18,gstin:""},
  {id:3,name:"Ajay Singh",phone:"9723456789",bike:"Bajaj Pulsar 150",area:"Model Town",totalBuy:7600,visits:11,points:76,gstin:"07AABCS1234A1Z5"},
];
const INIT_SUPPLIERS=[
  {id:1,name:"Sharma Traders",phone:"9811223344",city:"Delhi",parts:"Honda, Hero Parts",due:0,gstin:"07AAACS1234A1Z1"},
  {id:2,name:"Madan Auto Agencies",phone:"9955667788",city:"Ludhiana",parts:"Bajaj, TVS Parts",due:3200,gstin:"03AABCM1234A1Z2"},
  {id:3,name:"Gupta Brothers",phone:"9966778899",city:"Faridabad",parts:"Lubricants, Filters",due:0,gstin:"06AABCG1234A1Z3"},
];
const INIT_LEDGER=[
  {id:1,date:"2026-06-01",type:"cash",dir:"in",amt:700,note:"Brake Pad sale"},
  {id:2,date:"2026-06-02",type:"bank",dir:"in",amt:2600,note:"Engine Oil UPI"},
  {id:3,date:"2026-06-03",type:"cash",dir:"in",amt:1560,note:"Chain Sprocket sale"},
  {id:4,date:"2026-06-04",type:"bank",dir:"out",amt:5000,note:"Sharma Traders payment"},
  {id:5,date:"2026-06-07",type:"bank",dir:"in",amt:760,note:"Air Filter UPI"},
];
const MONTHLY=[
  {month:"Jan",revenue:18400,profit:6200},{month:"Feb",revenue:21000,profit:7400},
  {month:"Mar",revenue:26800,profit:9200},{month:"Apr",revenue:23500,profit:7900},
  {month:"May",revenue:31200,profit:11100},{month:"Jun",revenue:6300,profit:2315},
];
const DAILY=[
  {day:"Mon",sales:820},{day:"Tue",sales:1240},{day:"Wed",sales:640},
  {day:"Thu",sales:1800},{day:"Fri",sales:2100},{day:"Sat",sales:3200},{day:"Sun",sales:1100},
];

/* ═══ UI COMPONENTS ══════════════════════════════════════════════ */
function Card({children,style={},accent}){
  return <div style={{background:T.surface,borderRadius:16,padding:16,marginBottom:12,boxShadow:T.shadowMd,border:accent?`1.5px solid ${accent}35`:`1px solid ${T.border}`,...style}}>{children}</div>;
}
function StatCard({label,value,icon,color,lt,sub,onClick}){
  return(
    <div onClick={onClick} style={{background:lt||T.surface,borderRadius:14,padding:"13px 12px",boxShadow:T.shadow,border:`1px solid ${color}20`,cursor:onClick?"pointer":"default",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",top:-8,right:-8,width:52,height:52,borderRadius:"50%",background:`${color}12`}}/>
      <div style={{fontSize:"1rem",marginBottom:5}}>{icon}</div>
      <div style={{fontSize:"0.59rem",color:T.textSm,textTransform:"uppercase",letterSpacing:"1.2px",fontWeight:600,marginBottom:4}}>{label}</div>
      <div style={{fontSize:"1.2rem",fontWeight:800,color:T.text,lineHeight:1}}>{value}</div>
      {sub&&<div style={{fontSize:"0.62rem",color:T.textSm,marginTop:4}}>{sub}</div>}
      <div style={{position:"absolute",bottom:0,left:0,right:0,height:3,background:`linear-gradient(90deg,${color},${color}55)`,borderRadius:"0 0 14px 14px"}}/>
    </div>
  );
}
function Btn({children,color=T.saffron,textColor="#fff",onClick,full,sm,outline,ghost,disabled,style={}}){
  return(
    <button onClick={onClick} disabled={disabled} style={{padding:sm?"7px 13px":"10px 18px",borderRadius:10,border:outline?`1.5px solid ${color}`:ghost?"none":"none",background:disabled?`${color}40`:ghost?"transparent":outline?"transparent":color,color:disabled?"#aaa":outline||ghost?color:textColor,fontWeight:700,fontSize:sm?"0.71rem":"0.81rem",cursor:disabled?"not-allowed":"pointer",width:full?"100%":"auto",letterSpacing:"0.3px",transition:"all 0.14s",boxShadow:(!outline&&!ghost&&!disabled)?`0 2px 8px ${color}30`:"none",...style}}>
      {children}
    </button>
  );
}
function Inp({label,hint,...p}){
  const[foc,setFoc]=useState(false);
  return(
    <div style={{marginBottom:10}}>
      {label&&<div style={{fontSize:"0.62rem",color:T.textMd,textTransform:"uppercase",letterSpacing:"1px",fontWeight:700,marginBottom:5}}>{label}</div>}
      <input {...p} onFocus={e=>{setFoc(true);p.onFocus&&p.onFocus(e);}} onBlur={e=>{setFoc(false);p.onBlur&&p.onBlur(e);}}
        style={{width:"100%",padding:"10px 13px",borderRadius:10,border:`1.5px solid ${foc?T.saffron:T.border}`,background:foc?T.saffronXLt:T.surface,color:T.text,fontSize:"0.83rem",outline:"none",boxSizing:"border-box",transition:"all 0.15s",...p.style}}/>
      {hint&&<div style={{fontSize:"0.61rem",color:T.textSm,marginTop:3}}>{hint}</div>}
    </div>
  );
}
function Badge({children,color,lt}){
  return <span style={{display:"inline-flex",alignItems:"center",padding:"3px 9px",borderRadius:20,fontSize:"0.62rem",fontWeight:700,background:lt||`${color}15`,color,border:`1px solid ${color}25`}}>{children}</span>;
}
function Pill({children,active,color,onClick}){
  return <button onClick={onClick} style={{padding:"7px 14px",borderRadius:20,border:`1.5px solid ${active?color:T.border}`,background:active?color:"transparent",color:active?"#fff":T.textMd,fontSize:"0.72rem",fontWeight:active?700:500,cursor:"pointer",whiteSpace:"nowrap",transition:"all 0.14s",boxShadow:active?`0 2px 8px ${color}30`:"none"}}>{children}</button>;
}
function SecHead({title,action}){
  return(
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
      <div style={{display:"flex",alignItems:"center",gap:8}}>
        <div style={{width:3,height:13,background:`linear-gradient(180deg,${T.saffron},${T.saffronDk})`,borderRadius:2}}/>
        <span style={{fontSize:"0.68rem",fontWeight:700,textTransform:"uppercase",letterSpacing:"1.4px",color:T.textMd}}>{title}</span>
      </div>
      {action}
    </div>
  );
}
function Divider(){return <div style={{height:1,background:T.border,margin:"10px 0"}}/>;}
function Sheet({title,open,onClose,children}){
  if(!open)return null;
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.5)",zIndex:999,display:"flex",flexDirection:"column",justifyContent:"flex-end",backdropFilter:"blur(4px)"}}>
      <div style={{background:T.surface,borderRadius:"20px 20px 0 0",maxHeight:"93vh",overflowY:"auto",boxShadow:"0 -8px 40px rgba(0,0,0,0.15)"}}>
        <div style={{position:"sticky",top:0,background:T.surface,padding:"15px 20px 12px",borderBottom:`1px solid ${T.border}`,zIndex:1}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontWeight:800,fontSize:"0.98rem",color:T.text}}>{title}</span>
            <button onClick={onClose} style={{background:T.hi2,border:`1px solid ${T.border}`,color:T.textMd,width:30,height:30,borderRadius:8,cursor:"pointer",fontSize:"1rem",display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
          </div>
        </div>
        <div style={{padding:"16px 20px 28px"}}>{children}</div>
      </div>
    </div>
  );
}
function QtyCtrl({value,onChange,min=0}){
  return(
    <div style={{display:"flex",alignItems:"center",gap:6}}>
      <button onClick={()=>onChange(Math.max(min,value-1))} style={{width:28,height:28,borderRadius:8,border:`1.5px solid ${T.border}`,background:T.hi2,cursor:"pointer",fontSize:"1rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.textMd}}>−</button>
      <span style={{fontSize:"0.9rem",fontWeight:800,minWidth:28,textAlign:"center",color:T.text}}>{value}</span>
      <button onClick={()=>onChange(value+1)} style={{width:28,height:28,borderRadius:8,border:"none",background:T.saffron,cursor:"pointer",fontSize:"1rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",boxShadow:`0 2px 6px ${T.saffron}40`}}>+</button>
    </div>
  );
}

/* ═══ MAIN APP ══════════════════════════════════════════════════ */
export default function App(){
  const[tab,setTab]=useState(0);
  const[lang,setLang]=useState("en");
  const[stock,setStock]=useState(INIT_STOCK);
  const[sales,setSales]=useState(INIT_SALES);
  const[customers,setCustomers]=useState(INIT_CUSTOMERS);
  const[suppliers,setSuppliers]=useState(INIT_SUPPLIERS);
  const[ledger,setLedger]=useState(INIT_LEDGER);
  const[reminders,setReminders]=useState([
    {id:1,title:"Collect ₹680 from Suresh Yadav",cat:"Payment",time:"10:00",done:false,repeat:"once"},
    {id:2,title:"Reorder Headlight Bulbs",cat:"Reorder",time:"11:30",done:false,repeat:"once"},
    {id:3,title:"Pay Madan Auto ₹3200",cat:"Supplier",time:"14:00",done:false,repeat:"daily"},
  ]);
  const[staff,setStaff]=useState([
    {id:1,name:"Raju Verma",role:"Sales Staff",phone:"9876001234",salary:12000,attendance:{}},
    {id:2,name:"Deepak Kumar",role:"Store Helper",phone:"9876005678",salary:9000,attendance:{}},
  ]);
  const[showSearch,setShowSearch]=useState(false);
  const[searchQ,setSearchQ]=useState("");

  const t=TR[lang];

  useEffect(()=>{
    (async()=>{
      try{
        const d = await loadData();
        if(d){
          if(d.stock)setStock(d.stock);if(d.sales)setSales(d.sales);
          if(d.customers)setCustomers(d.customers);if(d.suppliers)setSuppliers(d.suppliers);
          if(d.ledger)setLedger(d.ledger);if(d.reminders)setReminders(d.reminders);
          if(d.staff)setStaff(d.staff);if(d.lang)setLang(d.lang);
        }
      }catch(e){console.error("Load error:",e);}
    })();
  },[]);

  const save=async(patch={})=>{
    try{await saveData({stock,sales,customers,suppliers,ledger,reminders,staff,lang,...patch});}
    catch(e){console.error("Save error:",e);}
  };

  const revenue=sales.reduce((a,s)=>a+s.total,0);
  const profit=sales.reduce((a,s)=>a+s.profit,0);
  const pending=sales.filter(s=>!s.paid).reduce((a,s)=>a+s.total,0);
  const lowStockItems=stock.filter(s=>s.qty<=s.minQty);
  const cashBal=ledger.filter(l=>l.type==="cash").reduce((a,l)=>a+(l.dir==="in"?l.amt:-l.amt),0);
  const bankBal=ledger.filter(l=>l.type==="bank").reduce((a,l)=>a+(l.dir==="in"?l.amt:-l.amt),0);

  const exportToExcel=(data,filename,sheetName)=>{
    try{const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,sheetName);XLSX.writeFile(wb,`KNS_${filename}.xlsx`);}catch(e){alert("Export failed: "+e.message);}
  };

  /* ── SEARCH OVERLAY ───────────────────────────────────────── */
  const SearchOverlay=()=>{
    const results=searchQ.length>1?[
      ...stock.filter(s=>s.name.toLowerCase().includes(searchQ.toLowerCase())||s.sku.toLowerCase().includes(searchQ.toLowerCase())).map(s=>({type:"Part",label:s.name,sub:`${s.sku} · ${s.sub}`,color:T.blue,val:`₹${s.sell} · ${s.qty} left`,icon:"⚙️"})),
      ...customers.filter(c=>c.name.toLowerCase().includes(searchQ.toLowerCase())||c.phone.includes(searchQ)).map(c=>({type:"Customer",label:c.name,sub:`📱 ${c.phone}`,color:T.green,val:`${c.points} pts`,icon:"👤"})),
      ...suppliers.filter(s=>s.name.toLowerCase().includes(searchQ.toLowerCase())).map(s=>({type:"Supplier",label:s.name,sub:`📍 ${s.city}`,color:T.saffron,val:s.due>0?`Due ₹${s.due}`:"Clear",icon:"🏭"})),
    ]:[];
    return(
      <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.55)",zIndex:1000,backdropFilter:"blur(6px)"}}>
        <div style={{background:T.surface}}>
          <div style={{display:"flex",gap:10,padding:"14px 16px",alignItems:"center",borderBottom:`1px solid ${T.border}`}}>
            <span>🔍</span>
            <input autoFocus value={searchQ} onChange={e=>setSearchQ(e.target.value)} placeholder={t.searchPlaceholder}
              style={{flex:1,fontSize:"0.95rem",border:"none",outline:"none",background:"transparent",color:T.text}}/>
            <button onClick={()=>{setShowSearch(false);setSearchQ("");}} style={{background:"none",border:"none",color:T.textSm,fontSize:"1.3rem",cursor:"pointer"}}>✕</button>
          </div>
          <div style={{maxHeight:"72vh",overflowY:"auto"}}>
            {results.length>0?results.map((r,i)=>(
              <div key={i} style={{display:"flex",gap:12,alignItems:"center",padding:"11px 16px",borderBottom:`1px solid ${T.hi2}`}}>
                <div style={{width:38,height:38,borderRadius:10,background:`${r.color}15`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1rem",flexShrink:0}}>{r.icon}</div>
                <div style={{flex:1}}><div style={{fontWeight:600,fontSize:"0.84rem",color:T.text}}>{r.label}</div><div style={{fontSize:"0.65rem",color:T.textSm,marginTop:1}}>{r.sub}</div></div>
                <div style={{textAlign:"right"}}><Badge color={r.color} lt={`${r.color}15`}>{r.type}</Badge><div style={{fontSize:"0.73rem",fontWeight:700,color:T.text,marginTop:3}}>{r.val}</div></div>
              </div>
            )):searchQ.length>1?<div style={{padding:"40px",textAlign:"center",color:T.textSm}}>No results found</div>:<div style={{padding:"40px",textAlign:"center",color:T.textSm}}>Type to search...</div>}
          </div>
        </div>
      </div>
    );
  };

  /* ── DASHBOARD ────────────────────────────────────────────── */
  const Dashboard=()=>(
    <div>
      <div style={{background:`linear-gradient(135deg,${T.saffron},${T.saffronDk})`,borderRadius:20,padding:"22px 20px",marginBottom:14,position:"relative",overflow:"hidden",boxShadow:`0 8px 24px ${T.saffron}40`}}>
        <div style={{position:"absolute",top:-20,right:-20,width:110,height:110,borderRadius:"50%",background:"rgba(255,255,255,0.1)"}}/>
        <div style={{fontSize:"0.59rem",color:"rgba(255,255,255,0.8)",fontWeight:700,textTransform:"uppercase",letterSpacing:"2px",marginBottom:6}}>{new Date().toLocaleDateString(lang==="hi"?"hi-IN":"en-IN",{weekday:"long",day:"numeric",month:"long"})}</div>
        <div style={{fontSize:"2rem",fontWeight:900,color:"#fff",letterSpacing:"-1px"}}>₹{revenue.toLocaleString("en-IN")}</div>
        <div style={{fontSize:"0.7rem",color:"rgba(255,255,255,0.8)",marginTop:4,marginBottom:14}}>{t.revenue} — {lang==="hi"?"इस महीने":"This Month"}</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
          {[[t.profit,`₹${profit.toLocaleString("en-IN")}`],[t.margin,`${revenue?Math.round((profit/revenue)*100):0}%`],[t.pending,`₹${pending.toLocaleString("en-IN")}`]].map(([l,v])=>(
            <div key={l} style={{background:"rgba(255,255,255,0.18)",borderRadius:10,padding:"9px 10px",backdropFilter:"blur(10px)"}}>
              <div style={{fontSize:"0.57rem",color:"rgba(255,255,255,0.75)",textTransform:"uppercase",letterSpacing:1}}>{l}</div>
              <div style={{fontSize:"0.9rem",fontWeight:800,color:"#fff",marginTop:2}}>{v}</div>
            </div>
          ))}
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
        <StatCard label={t.cashInHand} value={`₹${cashBal.toLocaleString("en-IN")}`} icon="💵" color={T.green} lt={T.greenLt} onClick={()=>setTab(4)}/>
        <StatCard label={t.bankBalance} value={`₹${bankBal.toLocaleString("en-IN")}`} icon="🏦" color={T.blue} lt={T.blueLt} onClick={()=>setTab(4)}/>
        <StatCard label={t.totalParts} value={stock.length} icon="📦" color={T.purple} lt={T.purpleLt} sub={`${lowStockItems.length} ${t.lowStock}`} onClick={()=>setTab(3)}/>
        <StatCard label={t.supplierDues} value={`₹${suppliers.reduce((a,s)=>a+s.due,0).toLocaleString("en-IN")}`} icon="🚛" color={T.red} lt={T.redLt}/>
      </div>
      <Card>
        <SecHead title={lang==="hi"?"इस सप्ताह की बिक्री":"Sales This Week"}/>
        <ResponsiveContainer width="100%" height={100}>
          <AreaChart data={DAILY}>
            <defs><linearGradient id="sg" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={T.saffron} stopOpacity={0.25}/><stop offset="95%" stopColor={T.saffron} stopOpacity={0}/></linearGradient></defs>
            <XAxis dataKey="day" tick={{fill:T.textSm,fontSize:10}} axisLine={false} tickLine={false}/>
            <Tooltip contentStyle={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:10,fontSize:"0.8rem"}}/>
            <Area type="monotone" dataKey="sales" stroke={T.saffron} strokeWidth={2.5} fill="url(#sg)" dot={{fill:T.saffron,r:3,strokeWidth:0}}/>
          </AreaChart>
        </ResponsiveContainer>
      </Card>
      {lowStockItems.length>0&&(
        <Card accent={T.red}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <div style={{width:32,height:32,borderRadius:8,background:T.redLt,display:"flex",alignItems:"center",justifyContent:"center"}}>⚠️</div>
            <div><div style={{fontWeight:700,color:T.red,fontSize:"0.84rem"}}>{t.lowStockAlert}</div><div style={{fontSize:"0.63rem",color:T.textSm}}>{lowStockItems.length} items need reorder</div></div>
          </div>
          {lowStockItems.map(i=>(
            <div key={i.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:`1px solid ${T.border}`}}>
              <div><div style={{fontSize:"0.8rem",fontWeight:600,color:T.text}}>{i.name}</div><div style={{fontSize:"0.63rem",color:T.textSm}}>{i.sub}</div></div>
              <Badge color={T.red} lt={T.redLt}>{i.qty} left</Badge>
            </div>
          ))}
        </Card>
      )}
      <Card>
        <SecHead title={t.recentTx} action={<Btn sm ghost color={T.saffron} onClick={()=>setTab(1)}>See All →</Btn>}/>
        {[...sales].slice(0,5).map(s=>(
          <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"9px 0",borderBottom:`1px solid ${T.hi2}`}}>
            <div style={{display:"flex",gap:10,alignItems:"center"}}>
              <div style={{width:34,height:34,borderRadius:10,background:T.saffronLt,display:"flex",alignItems:"center",justifyContent:"center"}}>⚙️</div>
              <div><div style={{fontSize:"0.8rem",fontWeight:600,color:T.text}}>{s.item}</div><div style={{fontSize:"0.63rem",color:T.textSm}}>{s.customer} · {s.date.slice(5)}</div></div>
            </div>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:"0.88rem",fontWeight:800,color:T.text}}>₹{s.total}</div>
              <Badge color={s.paid?(s.upi?T.blue:T.green):T.red} lt={s.paid?(s.upi?T.blueLt:T.greenLt):T.redLt}>{s.paid?(s.upi?"UPI":"Cash"):"Pending"}</Badge>
            </div>
          </div>
        ))}
      </Card>
    </div>
  );

  /* ── POS + SALES ──── FULLY FIXED ─────────────────────────── */
  const SalesTab=()=>{
    const[view,setView]=useState("pos");
    // POS state
    const[cartItems,setCartItems]=useState([]);
    const[custName,setCustName]=useState("");
    const[payMode,setPayMode]=useState("cash");
    const[saleResult,setSaleResult]=useState(null);
    // Invoice state
    const[invCust,setInvCust]=useState({name:"",phone:"",bike:"",gstin:""});
    const[invItems,setInvItems]=useState([{id:1,name:"",qty:1,price:0},{id:2,name:"",qty:1,price:0}]);
    const[invGenerated,setInvGenerated]=useState(false);
    const[invNo]=useState("KNS-"+String(Date.now()).slice(-4));
    // Daily closing state
    const[opBal,setOpBal]=useState("");

    // ── Cart helpers ──
    const cartTotal=cartItems.reduce((a,c)=>a+c.qty*c.sell,0);
    const cartProfit=cartItems.reduce((a,c)=>a+c.qty*(c.sell-c.buy),0);

    const addToCart=(part)=>{
      setCartItems(prev=>{
        const ex=prev.find(c=>c.id===part.id);
        if(ex)return prev.map(c=>c.id===part.id?{...c,qty:c.qty+1}:c);
        return[...prev,{...part,qty:1}];
      });
    };
    const updateCartQty=(id,qty)=>{
      if(qty<1){setCartItems(prev=>prev.filter(c=>c.id!==id));return;}
      setCartItems(prev=>prev.map(c=>c.id===id?{...c,qty}:c));
    };
    const removeFromCart=(id)=>setCartItems(prev=>prev.filter(c=>c.id!==id));

    const doCheckout=()=>{
      if(cartItems.length===0)return;
      const today=new Date().toISOString().slice(0,10);
      const newSales=cartItems.map(c=>({
        id:Date.now()+Math.random(),date:today,item:c.name,
        qty:c.qty,total:c.qty*c.sell,profit:c.qty*(c.sell-c.buy),
        customer:custName.trim()||"Walk-in",paid:payMode!=="credit",upi:payMode==="upi"
      }));
      const updatedSales=[...newSales,...sales];
      const updatedStock=stock.map(s=>{
        const ci=cartItems.find(c=>c.id===s.id);
        return ci?{...s,qty:Math.max(0,s.qty-ci.qty),sold:(s.sold||0)+ci.qty}:s;
      });
      // Loyalty points
      const ci=customers.findIndex(c=>c.name.toLowerCase()===(custName.trim()||"").toLowerCase());
      let updatedCustomers=[...customers];
      if(ci>=0){const pts=Math.floor(cartTotal/100);updatedCustomers[ci]={...updatedCustomers[ci],totalBuy:updatedCustomers[ci].totalBuy+cartTotal,visits:updatedCustomers[ci].visits+1,points:updatedCustomers[ci].points+pts};}
      setSales(updatedSales);setStock(updatedStock);setCustomers(updatedCustomers);
      save({sales:updatedSales,stock:updatedStock,customers:updatedCustomers});
      setSaleResult({total:cartTotal,profit:Math.round(cartProfit),items:cartItems.length,customer:custName.trim()||"Walk-in",mode:payMode});
      setCartItems([]);setCustName("");setPayMode("cash");
    };

    // ── Invoice helpers ──
    const iSubtotal=invItems.reduce((a,i)=>a+i.qty*i.price,0);
    const iCgst=Math.round(iSubtotal*0.09);
    const iSgst=Math.round(iSubtotal*0.09);
    const iGrand=iSubtotal+iCgst+iSgst;
    const updateInvItem=(id,field,val)=>setInvItems(prev=>prev.map(i=>i.id===id?{...i,[field]:field==="name"?val:+val}:i));
    const addInvItem=()=>setInvItems(prev=>[...prev,{id:Date.now(),name:"",qty:1,price:0}]);
    const removeInvItem=(id)=>{if(invItems.length>1)setInvItems(prev=>prev.filter(i=>i.id!==id));};

    const printInvoice=()=>{
      const invHTML=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Invoice ${invNo}</title>
      <style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;padding:28px;color:#111;max-width:580px;margin:0 auto}.hd{text-align:center;padding-bottom:14px;border-bottom:2px solid #F59E0B;margin-bottom:16px}.logo{font-size:1.3rem;font-weight:900;color:#D97706}.meta{font-size:0.75rem;color:#666;margin-top:4px}.info{font-size:0.82rem;margin-bottom:14px;line-height:1.6}table{width:100%;border-collapse:collapse;font-size:0.82rem;margin-bottom:12px}th{background:#FEF3C7;padding:8px 10px;text-align:left;border-bottom:2px solid #F59E0B}td{padding:8px 10px;border-bottom:1px solid #eee}.tot{font-size:0.9rem}.grand{font-size:1.1rem;font-weight:900;color:#D97706;background:#FFFBEB}.foot{text-align:center;font-size:0.75rem;color:#999;margin-top:18px;padding-top:14px;border-top:1px dashed #ddd}@media print{.no-print{display:none!important}}</style></head><body>
      <div class="hd"><div class="logo">🏍️ Kesri Nandan Spares</div><div class="meta">Wholesale & Retail Bike Spare Parts</div><div class="meta">Invoice No: <b>${invNo}</b> &nbsp;|&nbsp; Date: ${new Date().toLocaleDateString("en-IN")}</div></div>
      <div class="info"><b>Bill To:</b><br>${invCust.name||"Walk-in Customer"}<br>${invCust.phone?"📱 "+invCust.phone+"<br>":""}${invCust.bike?"🏍️ "+invCust.bike+"<br>":""}${invCust.gstin?"GSTIN: "+invCust.gstin+"<br>":""}</div>
      <table><thead><tr><th>#</th><th>Item</th><th>Qty</th><th>Rate</th><th>Amount</th></tr></thead><tbody>
      ${invItems.filter(i=>i.name).map((i,n)=>`<tr><td>${n+1}</td><td>${i.name}</td><td>${i.qty}</td><td>₹${i.price.toLocaleString("en-IN")}</td><td>₹${(i.qty*i.price).toLocaleString("en-IN")}</td></tr>`).join("")}
      </tbody></table>
      <table><tbody><tr class="tot"><td>Subtotal</td><td style="text-align:right">₹${iSubtotal.toLocaleString("en-IN")}</td></tr><tr class="tot"><td>CGST @ 9%</td><td style="text-align:right">₹${iCgst.toLocaleString("en-IN")}</td></tr><tr class="tot"><td>SGST @ 9%</td><td style="text-align:right">₹${iSgst.toLocaleString("en-IN")}</td></tr><tr class="grand"><td><b>GRAND TOTAL</b></td><td style="text-align:right"><b>₹${iGrand.toLocaleString("en-IN")}</b></td></tr></tbody></table>
      <div class="foot">Thank you for your business! 🙏<br>Kesri Nandan Spares — Your trusted bike parts partner</div>
      <br><button class="no-print" onclick="window.print()" style="display:block;margin:0 auto;padding:10px 28px;background:#F59E0B;border:none;border-radius:8px;color:#fff;font-size:1rem;font-weight:700;cursor:pointer">🖨️ Print Now</button>
      </body></html>`;
      const w=window.open("","_blank");
      if(w){w.document.write(invHTML);w.document.close();}
      else{
        // Fallback: copy and show
        navigator.clipboard.writeText(`Invoice ${invNo}\n${invCust.name}\n${invCust.phone}\n\n${invItems.filter(i=>i.name).map(i=>`${i.name} × ${i.qty} = ₹${i.qty*i.price}`).join("\n")}\n\nSubtotal: ₹${iSubtotal}\nCGST 9%: ₹${iCgst}\nSGST 9%: ₹${iSgst}\nTotal: ₹${iGrand}`).catch(()=>{});
        alert("Pop-up blocked. Invoice copied to clipboard instead!");
      }
    };

    const copyInvoiceWA=()=>{
      const txt=`🏍️ *KESRI NANDAN SPARES*\nInvoice: ${invNo} | ${new Date().toLocaleDateString("en-IN")}\n\nBill To: ${invCust.name||"Walk-in"}\n📱 ${invCust.phone||"—"} | 🏍️ ${invCust.bike||"—"}\n${invCust.gstin?"GSTIN: "+invCust.gstin+"\n":""}\n${"─".repeat(28)}\n${invItems.filter(i=>i.name).map((i,n)=>`${n+1}. ${i.name}\n   ${i.qty} × ₹${i.price} = ₹${i.qty*i.price}`).join("\n")}\n${"─".repeat(28)}\nSubtotal : ₹${iSubtotal.toLocaleString("en-IN")}\nCGST 9%  : ₹${iCgst.toLocaleString("en-IN")}\nSGST 9%  : ₹${iSgst.toLocaleString("en-IN")}\n${"─".repeat(28)}\n*TOTAL   : ₹${iGrand.toLocaleString("en-IN")}*\n\nDhanyavaad! 🙏 — Kesri Nandan Spares`;
      navigator.clipboard.writeText(txt).then(()=>alert("Invoice copied! Paste in WhatsApp.")).catch(()=>alert("Copy failed. Please try again."));
    };

    const todaySales=sales.filter(s=>s.date===new Date().toISOString().slice(0,10));
    const todayRev=todaySales.reduce((a,s)=>a+s.total,0);
    const todayCash=todaySales.filter(s=>s.paid&&!s.upi).reduce((a,s)=>a+s.total,0);
    const todayUPI=todaySales.filter(s=>s.paid&&s.upi).reduce((a,s)=>a+s.total,0);

    return(
      <div>
        <div style={{display:"flex",gap:6,marginBottom:14,overflowX:"auto",scrollbarWidth:"none",paddingBottom:2}}>
          {[[t.quickSale,"pos"],[t.salesLog,"list"],[t.gstInvoice,"invoice"],[t.dailyClosing,"closing"]].map(([l,v])=>(
            <Pill key={v} active={view===v} color={T.saffron} onClick={()=>{setView(v);if(v==="pos"){setSaleResult(null);}}}>{l}</Pill>
          ))}
        </div>

        {/* ── QUICK SALE ── */}
        {view==="pos"&&(
          saleResult?(
            <div>
              <Card accent={T.green} style={{textAlign:"center",padding:"30px 20px",marginBottom:14}}>
                <div style={{width:60,height:60,borderRadius:"50%",background:T.greenLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1.8rem",margin:"0 auto 14px"}}>✅</div>
                <div style={{fontWeight:800,fontSize:"1.1rem",color:T.green,marginBottom:6}}>{t.saleComplete}</div>
                <div style={{fontSize:"0.82rem",color:T.textMd,marginBottom:16}}>{saleResult.customer} · {saleResult.mode==="upi"?"UPI":saleResult.mode==="credit"?"Credit":"Cash"} · {saleResult.items} item(s)</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:18}}>
                  <div style={{background:T.greenLt,borderRadius:12,padding:"12px"}}>
                    <div style={{fontSize:"0.6rem",color:T.textSm,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>{t.totalSales}</div>
                    <div style={{fontWeight:900,fontSize:"1.2rem",color:T.green}}>₹{saleResult.total.toLocaleString("en-IN")}</div>
                  </div>
                  <div style={{background:T.saffronLt,borderRadius:12,padding:"12px"}}>
                    <div style={{fontSize:"0.6rem",color:T.textSm,textTransform:"uppercase",letterSpacing:1,marginBottom:4}}>{t.profit}</div>
                    <div style={{fontWeight:900,fontSize:"1.2rem",color:T.saffronDk}}>₹{saleResult.profit.toLocaleString("en-IN")}</div>
                  </div>
                </div>
                <Btn full color={T.saffron} onClick={()=>setSaleResult(null)}>⚡ {t.newSale}</Btn>
              </Card>
            </div>
          ):(
            <div>
              {/* Customer & payment */}
              <Card style={{padding:"12px 14px",marginBottom:10}}>
                <div style={{display:"flex",gap:8,marginBottom:10,alignItems:"center"}}>
                  <div style={{fontSize:"1rem"}}>👤</div>
                  <input value={custName} onChange={e=>setCustName(e.target.value)} placeholder={`${t.customerName} (optional)`}
                    style={{flex:1,padding:"9px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,background:T.surface,color:T.text,fontSize:"0.83rem",outline:"none"}}/>
                </div>
                <div style={{display:"flex",gap:7}}>
                  {[[t.cash,"cash","💵"],[t.upi,"upi","📲"],["Credit","credit","📝"]].map(([l,m,ic])=>(
                    <button key={m} onClick={()=>setPayMode(m)} style={{flex:1,padding:"8px 4px",borderRadius:9,border:`1.5px solid ${payMode===m?T.saffron:T.border}`,background:payMode===m?T.saffronLt:T.surface,color:payMode===m?T.saffronDk:T.textMd,fontSize:"0.72rem",fontWeight:payMode===m?700:500,cursor:"pointer",transition:"all 0.14s"}}>{ic} {l}</button>
                  ))}
                </div>
              </Card>

              {/* Parts grid */}
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
                {stock.map(p=>{
                  const inCart=cartItems.find(c=>c.id===p.id);
                  return(
                    <div key={p.id} onClick={()=>addToCart(p)} style={{background:T.surface,border:`1.5px solid ${inCart?T.saffron:T.border}`,borderRadius:13,padding:"11px 10px",cursor:"pointer",boxShadow:inCart?`0 0 0 3px ${T.saffron}20`:T.shadow,transition:"all 0.14s",position:"relative"}}>
                      {inCart&&<div style={{position:"absolute",top:7,right:9,width:22,height:22,borderRadius:"50%",background:T.saffron,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.7rem",fontWeight:800,color:"#fff",boxShadow:`0 2px 6px ${T.saffron}50`}}>{inCart.qty}</div>}
                      <div style={{fontSize:"0.8rem",fontWeight:700,color:T.text,marginBottom:2,paddingRight:inCart?22:0}}>{p.name}</div>
                      <div style={{fontSize:"0.61rem",color:T.textSm,marginBottom:6}}>{p.sub}</div>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                        <span style={{fontSize:"0.9rem",fontWeight:800,color:T.saffronDk}}>₹{p.sell}</span>
                        <Badge color={p.qty<=p.minQty?T.red:T.green} lt={p.qty<=p.minQty?T.redLt:T.greenLt}>{p.qty}</Badge>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Cart */}
              {cartItems.length>0&&(
                <div style={{position:"sticky",bottom:70,zIndex:10}}>
                  <div style={{background:T.surface,borderRadius:16,boxShadow:T.shadowLg,border:`1px solid ${T.border}`,padding:"14px 16px"}}>
                    <div style={{fontSize:"0.67rem",color:T.textSm,marginBottom:10,fontWeight:700,textTransform:"uppercase",letterSpacing:"1px"}}>🛒 Cart — {cartItems.length} item(s)</div>
                    {cartItems.map(c=>(
                      <div key={c.id} style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,padding:"6px 0",borderBottom:`1px solid ${T.hi2}`}}>
                        <div style={{flex:1,fontSize:"0.78rem",color:T.text,fontWeight:600}}>{c.name}</div>
                        <div style={{display:"flex",alignItems:"center",gap:6}}>
                          <button onClick={()=>updateCartQty(c.id,c.qty-1)} style={{width:24,height:24,borderRadius:7,border:`1px solid ${T.border}`,background:T.hi2,cursor:"pointer",fontSize:"0.9rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.textMd}}>−</button>
                          <span style={{fontSize:"0.82rem",fontWeight:800,minWidth:20,textAlign:"center"}}>{c.qty}</span>
                          <button onClick={()=>updateCartQty(c.id,c.qty+1)} style={{width:24,height:24,borderRadius:7,border:"none",background:T.saffron,cursor:"pointer",fontSize:"0.9rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff"}}>+</button>
                        </div>
                        <div style={{fontWeight:700,fontSize:"0.85rem",color:T.text,minWidth:55,textAlign:"right"}}>₹{(c.qty*c.sell).toLocaleString("en-IN")}</div>
                        <button onClick={()=>removeFromCart(c.id)} style={{background:T.redLt,border:"none",color:T.red,width:24,height:24,borderRadius:7,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.85rem",flexShrink:0}}>×</button>
                      </div>
                    ))}
                    <div style={{display:"flex",justifyContent:"space-between",fontWeight:800,fontSize:"1.05rem",marginBottom:12,paddingTop:6}}>
                      <span style={{color:T.text}}>Total</span>
                      <span style={{color:T.saffronDk}}>₹{cartTotal.toLocaleString("en-IN")}</span>
                    </div>
                    <Btn full color={T.saffron} onClick={doCheckout}>{t.checkout} — ₹{cartTotal.toLocaleString("en-IN")}</Btn>
                  </div>
                </div>
              )}
            </div>
          )
        )}

        {/* ── SALES LOG ── */}
        {view==="list"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
              <StatCard label={t.revenue} value={`₹${revenue.toLocaleString("en-IN")}`} icon="💰" color={T.saffron} lt={T.saffronLt}/>
              <StatCard label={t.profit} value={`₹${profit.toLocaleString("en-IN")}`} icon="📈" color={T.green} lt={T.greenLt}/>
            </div>
            {[...sales].map(s=>(
              <Card key={s.id} style={{padding:"12px 14px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div>
                    <div style={{fontWeight:700,fontSize:"0.85rem",color:T.text}}>{s.item}</div>
                    <div style={{fontSize:"0.65rem",color:T.textSm,marginTop:2}}>{s.customer} · {s.date.slice(5)} · Qty {s.qty}</div>
                    <div style={{marginTop:5,display:"flex",gap:5,flexWrap:"wrap"}}>
                      <Badge color={s.paid?(s.upi?T.blue:T.green):T.red} lt={s.paid?(s.upi?T.blueLt:T.greenLt):T.redLt}>{s.paid?(s.upi?"UPI":"Cash"):"Pending"}</Badge>
                      <Badge color={T.textSm} lt={T.hi2}>{t.profit} ₹{s.profit}</Badge>
                    </div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontWeight:800,fontSize:"0.98rem",color:T.text}}>₹{s.total}</div>
                    {!s.paid&&<Btn sm color={T.green} style={{marginTop:6}} onClick={()=>{const u=sales.map(x=>x.id===s.id?{...x,paid:true}:x);setSales(u);save({sales:u});}}>{t.markPaid}</Btn>}
                  </div>
                </div>
              </Card>
            ))}
          </div>
        )}

        {/* ── GST INVOICE ── FULLY FIXED ── */}
        {view==="invoice"&&(
          <div>
            <Card>
              <SecHead title={t.gstInvoice+" — "+invNo}/>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:4}}>
                <Inp label={t.customerName} placeholder="Ramesh Kumar" value={invCust.name} onChange={e=>setInvCust({...invCust,name:e.target.value})}/>
                <Inp label={t.phone} placeholder="9876543210" value={invCust.phone} onChange={e=>setInvCust({...invCust,phone:e.target.value})}/>
                <Inp label={t.bikeModel} placeholder="Honda Shine 125" value={invCust.bike} onChange={e=>setInvCust({...invCust,bike:e.target.value})}/>
                <Inp label={`${t.gstin} (optional)`} placeholder="07AABCS1234A1Z5" value={invCust.gstin} onChange={e=>setInvCust({...invCust,gstin:e.target.value})}/>
              </div>
            </Card>

            <Card>
              <SecHead title="Items" action={<Btn sm ghost color={T.saffron} onClick={addInvItem}>+ Add Item</Btn>}/>
              {invItems.map((item)=>(
                <div key={item.id} style={{marginBottom:10,background:T.hi2,borderRadius:12,padding:"10px 12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                    <span style={{fontSize:"0.67rem",color:T.textSm,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.8px"}}>Item</span>
                    <button onClick={()=>removeInvItem(item.id)} style={{background:T.redLt,border:"none",color:T.red,padding:"3px 8px",borderRadius:6,cursor:"pointer",fontSize:"0.7rem",fontWeight:600}}>Remove</button>
                  </div>
                  <Inp placeholder="Part / item name" value={item.name} onChange={e=>updateInvItem(item.id,"name",e.target.value)} style={{marginBottom:0}}/>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginTop:8}}>
                    <div>
                      <div style={{fontSize:"0.61rem",color:T.textSm,textTransform:"uppercase",letterSpacing:"0.8px",fontWeight:600,marginBottom:4}}>Qty</div>
                      <div style={{display:"flex",alignItems:"center",gap:6}}>
                        <button onClick={()=>updateInvItem(item.id,"qty",Math.max(1,item.qty-1))} style={{width:26,height:26,borderRadius:7,border:`1px solid ${T.border}`,background:T.surface,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.textMd}}>−</button>
                        <span style={{fontWeight:800,fontSize:"0.9rem",minWidth:20,textAlign:"center"}}>{item.qty}</span>
                        <button onClick={()=>updateInvItem(item.id,"qty",item.qty+1)} style={{width:26,height:26,borderRadius:7,border:"none",background:T.saffron,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff"}}>+</button>
                      </div>
                    </div>
                    <div>
                      <div style={{fontSize:"0.61rem",color:T.textSm,textTransform:"uppercase",letterSpacing:"0.8px",fontWeight:600,marginBottom:4}}>Rate ₹</div>
                      <input type="number" value={item.price||""} onChange={e=>updateInvItem(item.id,"price",e.target.value)} placeholder="0"
                        style={{width:"100%",padding:"7px 9px",borderRadius:9,border:`1px solid ${T.border}`,background:T.surface,color:T.text,fontSize:"0.84rem",outline:"none"}}/>
                    </div>
                    <div>
                      <div style={{fontSize:"0.61rem",color:T.textSm,textTransform:"uppercase",letterSpacing:"0.8px",fontWeight:600,marginBottom:4}}>Amount</div>
                      <div style={{padding:"7px 9px",borderRadius:9,background:`${T.saffron}15`,border:`1px solid ${T.saffron}25`,fontSize:"0.84rem",fontWeight:700,color:T.saffronDk}}>₹{(item.qty*item.price).toLocaleString("en-IN")}</div>
                    </div>
                  </div>
                </div>
              ))}
            </Card>

            {/* Tax Summary */}
            <Card accent={T.saffron}>
              <div style={{display:"flex",justifyContent:"space-between",padding:"6px 0",fontSize:"0.83rem"}}><span style={{color:T.textMd}}>Subtotal</span><span style={{fontWeight:600}}>₹{iSubtotal.toLocaleString("en-IN")}</span></div>
              <div style={{display:"flex",justifyContent:"space-between",padding:"6px 0",fontSize:"0.83rem",borderBottom:`1px solid ${T.border}`}}><span style={{color:T.textMd}}>CGST @ 9%</span><span style={{fontWeight:600}}>₹{iCgst.toLocaleString("en-IN")}</span></div>
              <div style={{display:"flex",justifyContent:"space-between",padding:"6px 0",fontSize:"0.83rem",borderBottom:`1px solid ${T.border}`}}><span style={{color:T.textMd}}>SGST @ 9%</span><span style={{fontWeight:600}}>₹{iSgst.toLocaleString("en-IN")}</span></div>
              <div style={{display:"flex",justifyContent:"space-between",padding:"10px 0 4px",fontWeight:800,fontSize:"1.1rem",color:T.saffronDk}}><span>Grand Total</span><span>₹{iGrand.toLocaleString("en-IN")}</span></div>
            </Card>

            {/* Action buttons */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              <Btn color={T.blue} full onClick={printInvoice}>🖨️ {t.print}</Btn>
              <Btn color={T.green} full onClick={copyInvoiceWA}>📋 {t.copyWA}</Btn>
            </div>
            <Btn color={T.saffron} full style={{marginTop:10}} onClick={()=>{
              exportToExcel([{InvoiceNo:invNo,Date:new Date().toLocaleDateString("en-IN"),Customer:invCust.name,Phone:invCust.phone,Bike:invCust.bike,GSTIN:invCust.gstin,...Object.fromEntries(invItems.filter(i=>i.name).map((i,n)=>[`Item${n+1}`,`${i.name} x${i.qty} @ ₹${i.price} = ₹${i.qty*i.price}`])),Subtotal:iSubtotal,"CGST 9%":iCgst,"SGST 9%":iSgst,"Grand Total":iGrand}],"Invoice_"+invNo,"Invoice");
            }}>📊 Export Invoice to Excel</Btn>
          </div>
        )}

        {/* ── DAILY CLOSING ── */}
        {view==="closing"&&(
          <div>
            <div style={{background:`linear-gradient(135deg,${T.blue},#1d4ed8)`,borderRadius:20,padding:"22px 20px",marginBottom:14,color:"#fff",boxShadow:`0 8px 24px ${T.blue}40`}}>
              <div style={{fontSize:"0.59rem",color:"rgba(255,255,255,0.75)",textTransform:"uppercase",letterSpacing:"2px",marginBottom:6}}>{t.dailyReport} — {new Date().toLocaleDateString("en-IN")}</div>
              <div style={{fontSize:"1.8rem",fontWeight:900,letterSpacing:"-0.5px"}}>₹{todayRev.toLocaleString("en-IN")}</div>
              <div style={{fontSize:"0.7rem",color:"rgba(255,255,255,0.8)",marginTop:4,marginBottom:14}}>{t.today} {t.totalSales}</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
                {[[t.cash,`₹${todayCash.toLocaleString("en-IN")}`],[t.upi,`₹${todayUPI.toLocaleString("en-IN")}`],["Bills",todaySales.length]].map(([l,v])=>(
                  <div key={l} style={{background:"rgba(255,255,255,0.18)",borderRadius:10,padding:"9px 10px"}}>
                    <div style={{fontSize:"0.57rem",color:"rgba(255,255,255,0.75)",textTransform:"uppercase",letterSpacing:1}}>{l}</div>
                    <div style={{fontSize:"0.9rem",fontWeight:800,marginTop:2}}>{v}</div>
                  </div>
                ))}
              </div>
            </div>
            <Card>
              <Inp label={`${t.openingBal} ₹`} type="number" placeholder="5000" value={opBal} onChange={e=>setOpBal(e.target.value)} hint="Enter today's opening cash balance"/>
              {opBal&&(
                <div style={{background:T.hi2,borderRadius:12,padding:"12px",marginTop:4}}>
                  {[[t.openingBal,`₹${Number(opBal).toLocaleString("en-IN")}`,T.textMd],[`${t.cash} Sales`,`₹${todayCash.toLocaleString("en-IN")}`,T.green],["UPI Sales",`₹${todayUPI.toLocaleString("en-IN")}`,T.blue],[t.closingBal,`₹${(Number(opBal)+todayCash).toLocaleString("en-IN")}`,T.saffronDk]].map(([l,v,c])=>(
                    <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${T.border}`}}>
                      <span style={{fontSize:"0.81rem",color:T.textMd}}>{l}</span>
                      <span style={{fontSize:"0.83rem",fontWeight:700,color:c}}>{v}</span>
                    </div>
                  ))}
                </div>
              )}
              <Btn full color={T.blue} style={{marginTop:12}} onClick={()=>{const txt=`📅 DAILY CLOSING — ${new Date().toLocaleDateString("en-IN")}\nKesri Nandan Spares\n${"─".repeat(28)}\nOpening Balance : ₹${Number(opBal||0).toLocaleString("en-IN")}\nCash Sales      : ₹${todayCash.toLocaleString("en-IN")}\nUPI Sales       : ₹${todayUPI.toLocaleString("en-IN")}\nTotal Sales     : ₹${todayRev.toLocaleString("en-IN")}\nPending Bills   : ${todaySales.filter(s=>!s.paid).length}\nClosing Balance : ₹${(Number(opBal||0)+todayCash).toLocaleString("en-IN")}\n${"─".repeat(28)}\nKesri Nandan Spares`;navigator.clipboard.writeText(txt).then(()=>alert("Daily report copied!")).catch(()=>{});}}>📋 Copy Daily Report</Btn>
            </Card>
            {todaySales.length>0&&(
              <Card><SecHead title={t.today+"'s Sales"}/>{todaySales.map(s=>(
                <div key={s.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.hi2}`}}>
                  <div><div style={{fontSize:"0.8rem",fontWeight:600,color:T.text}}>{s.item}</div><div style={{fontSize:"0.63rem",color:T.textSm}}>{s.customer}</div></div>
                  <div style={{display:"flex",gap:6,alignItems:"center"}}>
                    <Badge color={s.paid?(s.upi?T.blue:T.green):T.red} lt={s.paid?(s.upi?T.blueLt:T.greenLt):T.redLt}>{s.paid?(s.upi?"UPI":"Cash"):"Pending"}</Badge>
                    <span style={{fontWeight:700,fontSize:"0.84rem"}}>₹{s.total}</span>
                  </div>
                </div>
              ))}</Card>
            )}
          </div>
        )}
      </div>
    );
  };

  /* ── STOCK ─── FULLY UPGRADED WITH EDIT/QTY/PRICE ─────────── */
  const StockTab=()=>{
    const[view,setView]=useState("grid");
    const[search,setSearch]=useState("");
    const[addSheet,setAddSheet]=useState(false);
    const[editSheet,setEditSheet]=useState(null);
    const[qtySheet,setQtySheet]=useState(null);
    const[np,setNp]=useState({name:"",sub:"",cat:"",qty:"",minQty:"",buy:"",sell:"",bike:"",photo:null,sku:""});
    const fileRef=useRef();const photoEditRef=useRef();

    const filtered=stock.filter(s=>!search||(
      s.name.toLowerCase().includes(search.toLowerCase())||
      (s.sku||"").toLowerCase().includes(search.toLowerCase())||
      s.bike.toLowerCase().includes(search.toLowerCase())||
      s.cat.toLowerCase().includes(search.toLowerCase())
    ));
    const display=view==="low"?filtered.filter(s=>s.qty<=s.minQty):filtered;

    const updateStockQty=(id,newQty)=>{
      const u=stock.map(s=>s.id===id?{...s,qty:Math.max(0,newQty)}:s);
      setStock(u);save({stock:u});
    };
    const addPart=()=>{
      if(!np.name)return;
      const u=[...stock,{...np,id:Date.now(),qty:+np.qty||0,minQty:+np.minQty||5,buy:+np.buy||0,sell:+np.sell||0,sold:0}];
      setStock(u);save({stock:u});setAddSheet(false);
      setNp({name:"",sub:"",cat:"",qty:"",minQty:"",buy:"",sell:"",bike:"",photo:null,sku:""});
    };
    const savePart=(updated)=>{
      const u=stock.map(s=>s.id===updated.id?updated:s);
      setStock(u);save({stock:u});setEditSheet(null);
    };
    const deletePart=(id)=>{
      if(!window.confirm("Delete this part?"))return;
      const u=stock.filter(s=>s.id!==id);
      setStock(u);save({stock:u});setEditSheet(null);
    };
    const handlePhoto=(e,forNew)=>{
      const f=e.target.files[0];if(!f)return;
      const r=new FileReader();
      r.onload=ev=>{
        if(forNew)setNp({...np,photo:ev.target.result});
        else setEditSheet({...editSheet,photo:ev.target.result});
      };
      r.readAsDataURL(f);
    };

    const stockVal=stock.reduce((a,s)=>a+s.qty*s.buy,0);

    return(
      <div>
        {/* Search + Add */}
        <div style={{display:"flex",gap:8,marginBottom:12,alignItems:"center"}}>
          <div style={{flex:1,position:"relative"}}>
            <span style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:T.textSm}}>🔍</span>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder={t.searchParts}
              style={{width:"100%",padding:"10px 11px 10px 34px",borderRadius:10,border:`1.5px solid ${T.border}`,background:T.surface,color:T.text,fontSize:"0.82rem",outline:"none",boxSizing:"border-box"}}/>
          </div>
          <Btn sm color={T.saffron} onClick={()=>setAddSheet(true)}>+ {t.addPart}</Btn>
        </div>

        {/* View pills */}
        <div style={{display:"flex",gap:7,marginBottom:14,overflowX:"auto",scrollbarWidth:"none"}}>
          <Pill active={view==="grid"} color={T.saffron} onClick={()=>setView("grid")}>⊞ {t.gridView}</Pill>
          <Pill active={view==="list"} color={T.saffron} onClick={()=>setView("list")}>≡ {t.listView}</Pill>
          <Pill active={view==="table"} color={T.blue} onClick={()=>setView("table")}>📋 Table</Pill>
          <Pill active={view==="low"} color={T.red} onClick={()=>setView("low")}>⚠️ {t.lowStock} ({lowStockItems.length})</Pill>
        </div>

        {/* Stats */}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
          <StatCard label={t.totalParts} value={stock.length} icon="📦" color={T.blue} lt={T.blueLt} sub={`${lowStockItems.length} low stock`}/>
          <StatCard label={t.stockValue} value={`₹${(stockVal/1000).toFixed(1)}k`} icon="💰" color={T.saffron} lt={T.saffronLt} sub="At purchase price"/>
        </div>

        {/* GRID VIEW */}
        {view==="grid"&&(
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
            {display.map(p=>(
              <div key={p.id} style={{background:T.surface,border:`1.5px solid ${p.qty<=p.minQty?T.red+"45":T.border}`,borderRadius:14,overflow:"hidden",boxShadow:T.shadow}}>
                <div style={{height:88,background:p.photo?`url(${p.photo}) center/cover`:`linear-gradient(135deg,${T.saffronLt},${T.blueLt})`,display:"flex",alignItems:"center",justifyContent:"center",backgroundSize:"cover",position:"relative"}}>
                  {!p.photo&&<span style={{fontSize:"2rem",opacity:0.45}}>⚙️</span>}
                  <button onClick={()=>setEditSheet({...p})} style={{position:"absolute",top:6,right:6,width:26,height:26,borderRadius:7,background:"rgba(255,255,255,0.9)",border:"none",cursor:"pointer",fontSize:"0.75rem",display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.15)"}}>✏️</button>
                </div>
                <div style={{padding:"10px"}}>
                  <div style={{fontSize:"0.77rem",fontWeight:700,color:T.text,marginBottom:1}}>{p.name}</div>
                  <div style={{fontSize:"0.61rem",color:T.textSm,marginBottom:8}}>{p.sub}</div>
                  {/* Quick Qty Controls */}
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontSize:"0.87rem",fontWeight:800,color:T.saffronDk}}>₹{p.sell}</span>
                    <div style={{display:"flex",alignItems:"center",gap:5}}>
                      <button onClick={()=>updateStockQty(p.id,p.qty-1)} style={{width:22,height:22,borderRadius:6,border:`1px solid ${T.border}`,background:T.hi2,cursor:"pointer",fontSize:"0.8rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.textMd}}>−</button>
                      <span style={{fontSize:"0.8rem",fontWeight:800,minWidth:22,textAlign:"center",color:p.qty<=p.minQty?T.red:T.green}}>{p.qty}</span>
                      <button onClick={()=>updateStockQty(p.id,p.qty+1)} style={{width:22,height:22,borderRadius:6,border:"none",background:T.saffron,cursor:"pointer",fontSize:"0.8rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff"}}>+</button>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {/* LIST VIEW */}
        {view==="list"&&display.map(p=>(
          <div key={p.id} style={{display:"flex",gap:10,alignItems:"center",background:T.surface,border:`1px solid ${p.qty<=p.minQty?T.red+"35":T.border}`,borderRadius:13,padding:"11px 13px",marginBottom:8,boxShadow:T.shadow}}>
            <div style={{width:44,height:44,borderRadius:10,flexShrink:0,background:p.photo?`url(${p.photo}) center/cover`:`${T.saffronLt}`,backgroundSize:"cover",display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden"}}>
              {!p.photo&&<span style={{fontSize:"1.3rem"}}>⚙️</span>}
            </div>
            <div style={{flex:1,minWidth:0}}>
              <div style={{fontWeight:700,fontSize:"0.84rem",color:T.text}}>{p.name}</div>
              <div style={{fontSize:"0.62rem",color:T.textSm}}>{p.sku||"—"} · {p.cat} · ₹{p.sell}</div>
            </div>
            {/* Qty Controls */}
            <div style={{display:"flex",alignItems:"center",gap:5}}>
              <button onClick={()=>updateStockQty(p.id,p.qty-1)} style={{width:26,height:26,borderRadius:7,border:`1px solid ${T.border}`,background:T.hi2,cursor:"pointer",fontSize:"0.9rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.textMd}}>−</button>
              <span style={{fontSize:"0.88rem",fontWeight:800,minWidth:26,textAlign:"center",color:p.qty<=p.minQty?T.red:T.green}}>{p.qty}</span>
              <button onClick={()=>updateStockQty(p.id,p.qty+1)} style={{width:26,height:26,borderRadius:7,border:"none",background:T.saffron,cursor:"pointer",fontSize:"0.9rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff"}}>+</button>
            </div>
            <button onClick={()=>setEditSheet({...p})} style={{background:T.purpleLt,border:"none",color:T.purple,width:32,height:32,borderRadius:9,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.85rem",flexShrink:0}}>✏️</button>
          </div>
        ))}

        {/* TABLE VIEW — bulk edit mode */}
        {view==="table"&&(
          <Card style={{padding:"12px",overflowX:"auto"}}>
            <SecHead title="Stock Table — Tap to edit values"/>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:"0.76rem"}}>
              <thead>
                <tr style={{background:T.hi2}}>
                  {["Part","Qty","Buy ₹","Sell ₹",""].map(h=>(
                    <th key={h} style={{padding:"8px 8px",textAlign:"left",color:T.saffronDk,borderBottom:`2px solid ${T.saffron}30`,fontSize:"0.63rem",textTransform:"uppercase",letterSpacing:"0.8px",fontWeight:700}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {display.map(p=>(
                  <tr key={p.id} style={{borderBottom:`1px solid ${T.hi2}`}}>
                    <td style={{padding:"8px"}}>
                      <div style={{fontWeight:600,color:T.text,fontSize:"0.78rem"}}>{p.name}</div>
                      <div style={{fontSize:"0.6rem",color:T.textSm}}>{p.sku||"—"}</div>
                    </td>
                    <td style={{padding:"8px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:4}}>
                        <button onClick={()=>updateStockQty(p.id,p.qty-1)} style={{width:20,height:20,borderRadius:5,border:`1px solid ${T.border}`,background:T.hi2,cursor:"pointer",fontSize:"0.75rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.textMd,flexShrink:0}}>−</button>
                        <input type="number" value={p.qty} onChange={e=>updateStockQty(p.id,+e.target.value)}
                          style={{width:38,padding:"4px 5px",borderRadius:6,border:`1px solid ${T.border}`,background:T.surface,color:p.qty<=p.minQty?T.red:T.text,fontSize:"0.8rem",textAlign:"center",outline:"none",fontWeight:700}}/>
                        <button onClick={()=>updateStockQty(p.id,p.qty+1)} style={{width:20,height:20,borderRadius:5,border:"none",background:T.saffron,cursor:"pointer",fontSize:"0.75rem",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",flexShrink:0}}>+</button>
                      </div>
                    </td>
                    <td style={{padding:"8px"}}>
                      <input type="number" value={p.buy} onChange={e=>{const u=stock.map(s=>s.id===p.id?{...s,buy:+e.target.value}:s);setStock(u);save({stock:u});}}
                        style={{width:60,padding:"5px 7px",borderRadius:7,border:`1px solid ${T.border}`,background:T.surface,color:T.text,fontSize:"0.79rem",outline:"none",textAlign:"right"}}/>
                    </td>
                    <td style={{padding:"8px"}}>
                      <input type="number" value={p.sell} onChange={e=>{const u=stock.map(s=>s.id===p.id?{...s,sell:+e.target.value}:s);setStock(u);save({stock:u});}}
                        style={{width:60,padding:"5px 7px",borderRadius:7,border:`1.5px solid ${T.saffron}50`,background:T.saffronXLt,color:T.saffronDk,fontSize:"0.79rem",outline:"none",textAlign:"right",fontWeight:700}}/>
                    </td>
                    <td style={{padding:"8px"}}>
                      <button onClick={()=>setEditSheet({...p})} style={{background:T.purpleLt,border:"none",color:T.purple,padding:"5px 8px",borderRadius:7,cursor:"pointer",fontSize:"0.7rem",fontWeight:600}}>Edit</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div style={{marginTop:10,fontSize:"0.67rem",color:T.textSm,textAlign:"center"}}>Changes save automatically as you type</div>
          </Card>
        )}

        {view==="low"&&display.length===0&&<div style={{textAlign:"center",padding:"40px 0",color:T.textSm}}><div style={{fontSize:"2.5rem",marginBottom:8}}>✅</div>All stock levels are healthy!</div>}

        {/* EDIT PART SHEET */}
        {editSheet&&(
          <Sheet title={t.editPart} open={!!editSheet} onClose={()=>setEditSheet(null)}>
            {/* Photo */}
            <input ref={photoEditRef} type="file" accept="image/*" capture="environment" onChange={e=>handlePhoto(e,false)} style={{display:"none"}}/>
            <div onClick={()=>photoEditRef.current.click()} style={{height:110,borderRadius:12,background:editSheet.photo?`url(${editSheet.photo}) center/cover`:`linear-gradient(135deg,${T.saffronLt},${T.blueLt})`,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",marginBottom:14,backgroundSize:"cover",position:"relative"}}>
              {!editSheet.photo&&<><span style={{fontSize:"2rem",opacity:0.4}}>⚙️</span><div style={{position:"absolute",bottom:8,right:10,background:"rgba(0,0,0,0.5)",color:"#fff",padding:"4px 10px",borderRadius:7,fontSize:"0.7rem"}}>📷 Add Photo</div></>}
              {editSheet.photo&&<div style={{position:"absolute",bottom:8,right:10,background:"rgba(0,0,0,0.5)",color:"#fff",padding:"4px 10px",borderRadius:7,fontSize:"0.7rem"}}>🔄 Change</div>}
            </div>

            {/* Qty controls */}
            <div style={{background:T.hi2,borderRadius:12,padding:"14px",marginBottom:14}}>
              <div style={{fontSize:"0.62rem",color:T.textSm,textTransform:"uppercase",letterSpacing:1,fontWeight:700,marginBottom:10}}>Stock Quantity</div>
              <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:12}}>
                <button onClick={()=>setEditSheet({...editSheet,qty:Math.max(0,editSheet.qty-10)})} style={{padding:"8px 14px",borderRadius:9,border:`1px solid ${T.border}`,background:T.surface,cursor:"pointer",fontWeight:700,fontSize:"0.8rem",color:T.red}}>−10</button>
                <button onClick={()=>setEditSheet({...editSheet,qty:Math.max(0,editSheet.qty-1)})} style={{padding:"8px 14px",borderRadius:9,border:`1px solid ${T.border}`,background:T.surface,cursor:"pointer",fontWeight:700,fontSize:"0.8rem",color:T.textMd}}>−1</button>
                <div style={{textAlign:"center",minWidth:60}}>
                  <input type="number" value={editSheet.qty} onChange={e=>setEditSheet({...editSheet,qty:Math.max(0,+e.target.value)})}
                    style={{width:64,padding:"10px",borderRadius:10,border:`2px solid ${T.saffron}`,background:T.saffronXLt,color:T.saffronDk,fontSize:"1.2rem",fontWeight:900,textAlign:"center",outline:"none"}}/>
                  <div style={{fontSize:"0.6rem",color:T.textSm,marginTop:4}}>in stock</div>
                </div>
                <button onClick={()=>setEditSheet({...editSheet,qty:editSheet.qty+1})} style={{padding:"8px 14px",borderRadius:9,border:`1px solid ${T.border}`,background:T.surface,cursor:"pointer",fontWeight:700,fontSize:"0.8rem",color:T.textMd}}>+1</button>
                <button onClick={()=>setEditSheet({...editSheet,qty:editSheet.qty+10})} style={{padding:"8px 14px",borderRadius:9,border:`1px solid ${T.border}`,background:T.surface,cursor:"pointer",fontWeight:700,fontSize:"0.8rem",color:T.green}}>+10</button>
              </div>
            </div>

            {/* Price controls */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
              <div style={{background:`${T.green}10`,borderRadius:12,padding:"12px"}}>
                <div style={{fontSize:"0.62rem",color:T.textSm,textTransform:"uppercase",letterSpacing:1,fontWeight:700,marginBottom:8}}>Buy Price ₹</div>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <button onClick={()=>setEditSheet({...editSheet,buy:Math.max(0,editSheet.buy-5)})} style={{width:26,height:26,borderRadius:7,border:`1px solid ${T.border}`,background:T.surface,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.textMd,fontSize:"0.9rem"}}>−</button>
                  <input type="number" value={editSheet.buy} onChange={e=>setEditSheet({...editSheet,buy:Math.max(0,+e.target.value)})}
                    style={{flex:1,padding:"8px",borderRadius:9,border:`1px solid ${T.border}`,background:T.surface,color:T.text,fontSize:"0.9rem",fontWeight:700,textAlign:"center",outline:"none"}}/>
                  <button onClick={()=>setEditSheet({...editSheet,buy:editSheet.buy+5})} style={{width:26,height:26,borderRadius:7,border:"none",background:T.green,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:"0.9rem"}}>+</button>
                </div>
              </div>
              <div style={{background:`${T.saffron}10`,borderRadius:12,padding:"12px"}}>
                <div style={{fontSize:"0.62rem",color:T.textSm,textTransform:"uppercase",letterSpacing:1,fontWeight:700,marginBottom:8}}>Sell Price ₹</div>
                <div style={{display:"flex",alignItems:"center",gap:6}}>
                  <button onClick={()=>setEditSheet({...editSheet,sell:Math.max(0,editSheet.sell-5)})} style={{width:26,height:26,borderRadius:7,border:`1px solid ${T.border}`,background:T.surface,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:T.textMd,fontSize:"0.9rem"}}>−</button>
                  <input type="number" value={editSheet.sell} onChange={e=>setEditSheet({...editSheet,sell:Math.max(0,+e.target.value)})}
                    style={{flex:1,padding:"8px",borderRadius:9,border:`1.5px solid ${T.saffron}60`,background:T.saffronXLt,color:T.saffronDk,fontSize:"0.9rem",fontWeight:700,textAlign:"center",outline:"none"}}/>
                  <button onClick={()=>setEditSheet({...editSheet,sell:editSheet.sell+5})} style={{width:26,height:26,borderRadius:7,border:"none",background:T.saffron,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:"0.9rem"}}>+</button>
                </div>
              </div>
            </div>

            {/* Margin indicator */}
            <div style={{background:T.hi2,borderRadius:10,padding:"10px 14px",marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{fontSize:"0.78rem",color:T.textMd}}>Profit Margin</span>
              <span style={{fontWeight:800,fontSize:"1rem",color:T.green}}>{editSheet.buy>0?Math.round(((editSheet.sell-editSheet.buy)/editSheet.buy)*100):0}%</span>
            </div>

            {/* Other fields */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <Inp label={t.partName+" *"} value={editSheet.name} onChange={e=>setEditSheet({...editSheet,name:e.target.value})}/>
              <Inp label={t.subtitle} value={editSheet.sub||""} onChange={e=>setEditSheet({...editSheet,sub:e.target.value})}/>
              <Inp label={t.category} value={editSheet.cat||""} onChange={e=>setEditSheet({...editSheet,cat:e.target.value})}/>
              <Inp label={t.bikeModel} value={editSheet.bike||""} onChange={e=>setEditSheet({...editSheet,bike:e.target.value})}/>
              <Inp label={t.sku} value={editSheet.sku||""} onChange={e=>setEditSheet({...editSheet,sku:e.target.value})}/>
              <Inp label={t.minAlert} type="number" value={editSheet.minQty||""} onChange={e=>setEditSheet({...editSheet,minQty:+e.target.value})}/>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginTop:4}}>
              <Btn full color={T.saffron} onClick={()=>savePart(editSheet)}>✅ {t.save}</Btn>
              <Btn full color={T.red} outline onClick={()=>deletePart(editSheet.id)}>🗑️ Delete</Btn>
            </div>
          </Sheet>
        )}

        {/* ADD PART SHEET */}
        <Sheet title={`+ ${t.addPart}`} open={addSheet} onClose={()=>setAddSheet(false)}>
          <input ref={fileRef} type="file" accept="image/*" capture="environment" onChange={e=>handlePhoto(e,true)} style={{display:"none"}}/>
          <div onClick={()=>fileRef.current.click()} style={{height:90,borderRadius:12,background:np.photo?`url(${np.photo}) center/cover`:`${T.saffronXLt}`,border:`2px dashed ${T.saffron}55`,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",cursor:"pointer",marginBottom:14,backgroundSize:"cover"}}>
            {!np.photo&&<><span style={{fontSize:"1.5rem"}}>📷</span><span style={{fontSize:"0.7rem",color:T.textMd,marginTop:4}}>Add part photo</span></>}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <div style={{gridColumn:"1/-1"}}><Inp label={`${t.partName} *`} placeholder="Brake Pad Honda CB" value={np.name} onChange={e=>setNp({...np,name:e.target.value})}/></div>
            <Inp label={t.subtitle} placeholder="Front disc" value={np.sub} onChange={e=>setNp({...np,sub:e.target.value})}/>
            <Inp label={t.sku} placeholder="BP-HS-009" value={np.sku} onChange={e=>setNp({...np,sku:e.target.value})}/>
            <Inp label={t.category} placeholder="Brakes" value={np.cat} onChange={e=>setNp({...np,cat:e.target.value})}/>
            <Inp label={t.bikeModel} placeholder="Honda Shine" value={np.bike} onChange={e=>setNp({...np,bike:e.target.value})}/>
            <Inp label={t.qty} type="number" placeholder="50" value={np.qty} onChange={e=>setNp({...np,qty:e.target.value})}/>
            <Inp label={t.minAlert} type="number" placeholder="10" value={np.minQty} onChange={e=>setNp({...np,minQty:e.target.value})}/>
            <Inp label={`${t.buyPrice} ₹`} type="number" placeholder="100" value={np.buy} onChange={e=>setNp({...np,buy:e.target.value})}/>
            <Inp label={`${t.sellPrice} ₹`} type="number" placeholder="160" value={np.sell} onChange={e=>setNp({...np,sell:e.target.value})}/>
          </div>
          {np.buy&&np.sell&&<div style={{background:T.greenLt,borderRadius:10,padding:"10px 12px",marginBottom:10,textAlign:"center",fontSize:"0.8rem",color:T.green,fontWeight:700}}>Margin: {Math.round(((+np.sell-+np.buy)/+np.buy)*100)}%</div>}
          <Btn full color={T.saffron} onClick={addPart}>✅ {t.save}</Btn>
        </Sheet>
      </div>
    );
  };

  /* ── ANALYTICS ────────────────────────────────────────────── */
  const Analytics=()=>{
    const catData=stock.reduce((acc,s)=>{const ex=acc.find(a=>a.name===s.cat);if(ex)ex.value+=s.qty*s.sell;else acc.push({name:s.cat,value:s.qty*s.sell});return acc;},[]);
    const topParts=[...stock].sort((a,b)=>(b.sold||0)-(a.sold||0)).slice(0,5);
    return(
      <div>
        <Card>
          <SecHead title={`${t.revenue} & ${t.profit} — 6 Months`}/>
          <div style={{display:"flex",gap:14,marginBottom:10}}>
            {[[t.revenue,T.saffron],[t.profit,T.green]].map(([l,c])=>(<div key={l} style={{display:"flex",alignItems:"center",gap:5}}><div style={{width:10,height:10,borderRadius:2,background:c}}/><span style={{fontSize:"0.67rem",color:T.textMd}}>{l}</span></div>))}
          </div>
          <ResponsiveContainer width="100%" height={150}>
            <BarChart data={MONTHLY} barGap={3}>
              <XAxis dataKey="month" tick={{fill:T.textSm,fontSize:10}} axisLine={false} tickLine={false}/>
              <YAxis tick={{fill:T.textSm,fontSize:9}} axisLine={false} tickLine={false} tickFormatter={v=>`₹${v/1000}k`}/>
              <Tooltip contentStyle={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:10,fontSize:"0.8rem"}}/>
              <Bar dataKey="revenue" fill={T.saffron} radius={[5,5,0,0]}/>
              <Bar dataKey="profit" fill={T.green} radius={[5,5,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <SecHead title="Stock by Category"/>
          <div style={{display:"flex",alignItems:"center",gap:12}}>
            <ResponsiveContainer width={130} height={130}>
              <PieChart><Pie data={catData} cx="50%" cy="50%" innerRadius={35} outerRadius={58} paddingAngle={3} dataKey="value">{catData.map((_,i)=><Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>)}</Pie><Tooltip contentStyle={{background:T.surface,border:`1px solid ${T.border}`,borderRadius:10,fontSize:"0.8rem"}}/></PieChart>
            </ResponsiveContainer>
            <div style={{flex:1}}>{catData.map((c,i)=>(<div key={i} style={{display:"flex",alignItems:"center",gap:7,marginBottom:7}}><div style={{width:8,height:8,borderRadius:"50%",background:PIE_COLORS[i%PIE_COLORS.length],flexShrink:0}}/><div style={{flex:1,fontSize:"0.72rem",color:T.text}}>{c.name}</div><div style={{fontSize:"0.69rem",color:T.textMd,fontWeight:600}}>₹{(c.value/1000).toFixed(1)}k</div></div>))}</div>
          </div>
        </Card>
        <Card>
          <SecHead title={t.bestSellers}/>
          {topParts.map((p,i)=>{const pct=Math.round(((p.sold||0)/Math.max(topParts[0].sold||1,1))*100);return(
            <div key={p.id} style={{marginBottom:11}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{fontSize:"0.78rem",fontWeight:700,color:T.text}}>{p.name} <span style={{color:T.textSm,fontWeight:400}}>· {p.sub}</span></span><span style={{fontSize:"0.71rem",color:PIE_COLORS[i],fontWeight:700}}>{p.sold||0} sold</span></div>
              <div style={{height:6,borderRadius:3,background:T.hi2,overflow:"hidden"}}><div style={{height:"100%",width:`${pct}%`,borderRadius:3,background:PIE_COLORS[i]}}/></div>
            </div>
          );})}
        </Card>
        <Card>
          <SecHead title={t.dailySales}/>
          <div style={{display:"flex",gap:5,alignItems:"flex-end"}}>
            {DAILY.map((d,i)=>{const pct=d.sales/3200;const isMax=d.sales===3200;return(
              <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:3}}>
                <div style={{fontSize:"0.57rem",color:isMax?T.saffronDk:T.textSm,fontWeight:isMax?700:400}}>₹{(d.sales/1000).toFixed(1)}k</div>
                <div style={{width:"100%",height:Math.round(pct*65+12),borderRadius:5,background:isMax?T.saffron:`${T.saffron}35`,transition:"height 0.4s"}}/>
                <div style={{fontSize:"0.61rem",color:T.textMd,fontWeight:isMax?700:400}}>{d.day}</div>
              </div>
            );})}
          </div>
        </Card>
        <Card>
          <SecHead title={t.exportExcel}/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            {[["📊 Sales",T.green,()=>exportToExcel(sales.map(s=>({Date:s.date,Item:s.item,Customer:s.customer,Qty:s.qty,"Total ₹":s.total,"Profit ₹":s.profit,Status:s.paid?(s.upi?"UPI":"Cash"):"Pending"})),"Sales","Sales")],["📦 Stock",T.blue,()=>exportToExcel(stock.map(s=>({SKU:s.sku,Name:s.name,Category:s.cat,Bike:s.bike,Qty:s.qty,Min:s.minQty,"Buy ₹":s.buy,"Sell ₹":s.sell,"Sold":s.sold||0,"Margin %":s.buy?Math.round(((s.sell-s.buy)/s.buy)*100):0})),"Stock","Stock")],["👥 Customers",T.purple,()=>exportToExcel(customers.map(c=>({Name:c.name,Phone:c.phone,Bike:c.bike,Area:c.area,"Purchase ₹":c.totalBuy,Visits:c.visits,"Points":c.points})),"Customers","Customers")],["💰 Ledger",T.saffron,()=>exportToExcel(ledger.map(l=>({Date:l.date,Type:l.type,Dir:l.dir==="in"?"In":"Out","Amount ₹":l.amt,Note:l.note})),"Ledger","Ledger")]].map(([l,c,fn])=>(
              <button key={l} onClick={fn} style={{display:"flex",alignItems:"center",gap:8,padding:"11px 12px",borderRadius:11,border:`1.5px solid ${c}25`,background:`${c}08`,cursor:"pointer",textAlign:"left"}}>
                <div style={{width:30,height:30,borderRadius:8,background:`${c}18`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.95rem",flexShrink:0}}>{l.split(" ")[0]}</div>
                <div><div style={{fontSize:"0.72rem",fontWeight:700,color:c}}>{l.slice(3)}</div><div style={{fontSize:"0.6rem",color:T.textSm}}>Export .xlsx</div></div>
              </button>
            ))}
          </div>
        </Card>
      </div>
    );
  };

  /* ── FINANCE ──────────────────────────────────────────────── */
  const Finance=()=>{
    const[sub,setSub]=useState("ledger");
    const[sheet,setSheet]=useState(false);
    const[nl,setNl]=useState({type:"cash",dir:"in",amt:"",note:""});
    const[selMonth,setSelMonth]=useState("Jun");
    const[expenses,setExpenses]=useState({rent:8000,electricity:1200,salary:15000,misc:1500});
    const addEntry=()=>{if(!nl.amt)return;const u=[...ledger,{id:Date.now(),date:new Date().toISOString().slice(0,10),...nl,amt:+nl.amt}];setLedger(u);save({ledger:u});setSheet(false);setNl({type:"cash",dir:"in",amt:"",note:""});};
    const m=MONTHLY.find(h=>h.month===selMonth)||MONTHLY[5];
    const totalExp=Object.values(expenses).reduce((a,v)=>a+Number(v),0);
    const cogs=Math.round(m.revenue*0.55);const grossProfit=m.revenue-cogs;const netProfit=grossProfit-totalExp;
    return(
      <div>
        <div style={{display:"flex",gap:7,marginBottom:14,overflowX:"auto",scrollbarWidth:"none"}}>
          {[[t.ledger,"ledger"],[t.plReport,"pl"],["UPI","upi"]].map(([l,v])=>(<Pill key={v} active={sub===v} color={T.saffron} onClick={()=>setSub(v)}>{l}</Pill>))}
        </div>
        {sub==="ledger"&&(<>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
            <StatCard label={t.cashInHand} value={`₹${cashBal.toLocaleString("en-IN")}`} icon="💵" color={T.green} lt={T.greenLt}/>
            <StatCard label={t.bankBalance} value={`₹${bankBal.toLocaleString("en-IN")}`} icon="🏦" color={T.blue} lt={T.blueLt}/>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><SecHead title={t.ledger}/><Btn sm color={T.saffron} onClick={()=>setSheet(true)}>+ {t.addEntry}</Btn></div>
          {[...ledger].reverse().map(l=>(<div key={l.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:T.surface,border:`1px solid ${T.border}`,borderRadius:12,padding:"11px 14px",marginBottom:8,boxShadow:T.shadow}}>
            <div style={{display:"flex",gap:10,alignItems:"center"}}>
              <div style={{width:37,height:37,borderRadius:10,background:l.dir==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.95rem"}}>{l.type==="cash"?l.dir==="in"?"💵":"💸":l.dir==="in"?"📲":"🏧"}</div>
              <div><div style={{fontWeight:600,fontSize:"0.82rem",color:T.text}}>{l.note}</div><div style={{fontSize:"0.63rem",color:T.textSm,marginTop:2,display:"flex",gap:5,alignItems:"center"}}>{l.date.slice(5)} · <Badge color={l.type==="cash"?T.saffron:T.blue} lt={l.type==="cash"?T.saffronLt:T.blueLt}>{l.type==="cash"?t.cash:t.bank}</Badge></div></div>
            </div>
            <div style={{fontWeight:800,fontSize:"0.98rem",color:l.dir==="in"?T.green:T.red}}>{l.dir==="in"?"+":"−"}₹{l.amt.toLocaleString("en-IN")}</div>
          </div>))}
          <Sheet title={t.addEntry} open={sheet} onClose={()=>setSheet(false)}>
            <div style={{display:"flex",gap:8,marginBottom:12}}><Pill active={nl.type==="cash"} color={T.saffron} onClick={()=>setNl({...nl,type:"cash"})}>💵 {t.cash}</Pill><Pill active={nl.type==="bank"} color={T.blue} onClick={()=>setNl({...nl,type:"bank"})}>🏦 {t.bank}</Pill></div>
            <div style={{display:"flex",gap:8,marginBottom:14}}><Pill active={nl.dir==="in"} color={T.green} onClick={()=>setNl({...nl,dir:"in"})}>📥 {t.moneyIn}</Pill><Pill active={nl.dir==="out"} color={T.red} onClick={()=>setNl({...nl,dir:"out"})}>📤 {t.moneyOut}</Pill></div>
            <Inp label={`${t.amount} ₹ *`} type="number" placeholder="500" value={nl.amt} onChange={e=>setNl({...nl,amt:e.target.value})}/>
            <Inp label={t.note} placeholder="e.g. Brake pad sale" value={nl.note} onChange={e=>setNl({...nl,note:e.target.value})}/>
            <Btn full color={T.saffron} onClick={addEntry} style={{marginTop:4}}>✅ {t.save}</Btn>
          </Sheet>
        </>)}
        {sub==="pl"&&(<>
          <div style={{display:"flex",gap:6,overflowX:"auto",marginBottom:14,scrollbarWidth:"none"}}>{MONTHLY.map(h=><Pill key={h.month} active={selMonth===h.month} color={T.saffron} onClick={()=>setSelMonth(h.month)}>{h.month}</Pill>)}</div>
          <div style={{background:`linear-gradient(135deg,${netProfit>=0?T.green:T.red},${netProfit>=0?T.greenMd||T.green:T.red}cc)`,borderRadius:20,padding:"20px",marginBottom:14,textAlign:"center",boxShadow:`0 8px 24px ${netProfit>=0?T.green:T.red}30`}}>
            <div style={{fontSize:"0.6rem",color:"rgba(255,255,255,0.8)",textTransform:"uppercase",letterSpacing:"2px",marginBottom:8}}>{t.netProfit} — {selMonth} 2026</div>
            <div style={{fontSize:"2rem",fontWeight:900,color:"#fff",letterSpacing:"-1px"}}>₹{netProfit.toLocaleString("en-IN")}</div>
            <div style={{fontSize:"0.78rem",color:"rgba(255,255,255,0.85)",marginTop:5}}>{t.margin}: {m.revenue?Math.round((netProfit/m.revenue)*100):0}%</div>
          </div>
          <Card accent={T.saffron}>
            {[[t.grossRevenue,m.revenue,T.text,false],[t.cogs,cogs,T.red,true],[t.grossProfit,grossProfit,T.green,false]].map(([l,v,c,minus])=>(
              <div key={l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 0",borderBottom:`1px solid ${T.hi2}`}}>
                <span style={{fontSize:"0.82rem",fontWeight:600,color:T.text}}>{l}</span>
                <span style={{fontWeight:700,color:c,fontSize:"0.88rem"}}>{minus?"− ":""}₹{v.toLocaleString("en-IN")}</span>
              </div>
            ))}
          </Card>
          <Card>
            <SecHead title={t.totalExp}/>
            {[[t.rent,"rent"],[t.electricity,"electricity"],[t.salary,"salary"],[t.misc,"misc"]].map(([l,k])=>(
              <div key={k} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:`1px solid ${T.hi2}`}}>
                <span style={{fontSize:"0.81rem",color:T.text}}>{l}</span>
                <div style={{display:"flex",alignItems:"center",gap:5}}><span style={{fontSize:"0.7rem",color:T.textSm}}>₹</span>
                  <input type="number" value={expenses[k]} onChange={e=>setExpenses({...expenses,[k]:e.target.value})} style={{width:74,padding:"5px 8px",borderRadius:8,border:`1px solid ${T.border}`,background:T.hi2,color:T.text,fontSize:"0.81rem",textAlign:"right",outline:"none"}}/>
                </div>
              </div>
            ))}
            <div style={{display:"flex",justifyContent:"space-between",padding:"9px 0",fontWeight:700}}><span>Total</span><span style={{color:T.red}}>₹{totalExp.toLocaleString("en-IN")}</span></div>
          </Card>
          <Btn full color={T.saffron} onClick={()=>{const txt=`P&L ${selMonth} 2026 — Kesri Nandan Spares\n${t.grossRevenue}: ₹${m.revenue.toLocaleString("en-IN")}\n${t.cogs}: ₹${cogs.toLocaleString("en-IN")}\n${t.grossProfit}: ₹${grossProfit.toLocaleString("en-IN")}\n${t.totalExp}: ₹${totalExp.toLocaleString("en-IN")}\n${t.netProfit}: ₹${netProfit.toLocaleString("en-IN")}`;navigator.clipboard.writeText(txt).then(()=>alert("Copied!")).catch(()=>{});}}>📋 {t.copyWA}</Btn>
        </>)}
        {sub==="upi"&&(<>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
            <StatCard label="UPI In" value={`₹${ledger.filter(l=>l.type==="bank"&&l.dir==="in").reduce((a,l)=>a+l.amt,0).toLocaleString("en-IN")}`} icon="📲" color={T.green} lt={T.greenLt}/>
            <StatCard label="UPI Out" value={`₹${ledger.filter(l=>l.type==="bank"&&l.dir==="out").reduce((a,l)=>a+l.amt,0).toLocaleString("en-IN")}`} icon="💸" color={T.red} lt={T.redLt}/>
          </div>
          {ledger.filter(l=>l.type==="bank").reverse().map(u=>(<div key={u.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:T.surface,border:`1px solid ${T.border}`,borderRadius:12,padding:"11px 14px",marginBottom:8,boxShadow:T.shadow}}>
            <div style={{display:"flex",gap:10,alignItems:"center"}}><div style={{width:37,height:37,borderRadius:10,background:u.dir==="in"?T.greenLt:T.redLt,display:"flex",alignItems:"center",justifyContent:"center"}}>{u.dir==="in"?"📲":"🏧"}</div><div><div style={{fontWeight:600,fontSize:"0.82rem",color:T.text}}>{u.note}</div><div style={{fontSize:"0.63rem",color:T.textSm,marginTop:2}}>{u.date.slice(5)}</div></div></div>
            <div style={{fontWeight:800,fontSize:"0.98rem",color:u.dir==="in"?T.green:T.red}}>{u.dir==="in"?"+":"−"}₹{u.amt.toLocaleString("en-IN")}</div>
          </div>))}
        </>)}
      </div>
    );
  };

  /* ── TOOLS ────────────────────────────────────────────────── */
  const Tools=()=>{
    const[sub,setSub]=useState("whatsapp");
    const[custMsg,setCustMsg]=useState("");
    const[scenario,setScenario]=useState("general");
    const[waTone,setWaTone]=useState("hinglish");
    const[reply,setReply]=useState("");
    const[waLoading,setWaLoading]=useState(false);
    const[remSheet,setRemSheet]=useState(false);
    const[nr,setNr]=useState({title:"",cat:"Payment",time:"10:00",repeat:"once"});
    const today=new Date().toISOString().slice(0,10);
    const CAT_C={Payment:T.red,Reorder:T.saffron,Supplier:T.blue,Customer:T.green,Other:T.purple};
    const CAT_I={Payment:"💸",Reorder:"📦",Supplier:"🚛",Customer:"👤",Other:"🔔"};
    const genReply=async()=>{
      if(!custMsg.trim())return;setWaLoading(true);setReply("");
      const tMap={hinglish:"natural Hinglish (Hindi+English mix)",hindi:"simple pure Hindi",english:"clear English",formal:"formal professional English"};
      const sMap={general:`Customer: "${custMsg}". Reply helpfully.`,available:`"${custMsg}" — say YES available, invite to visit Kesri Nandan Spares.`,notavail:`"${custMsg}" — not available, will inform, ask number.`,price:`"${custMsg}" — best prices, ask bike model.`,complaint:`"${custMsg}" — empathy, apology, resolve.`,promo:`Promo about: "${custMsg}". Attractive with urgency.`};
      try{const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:350,messages:[{role:"user",content:`WhatsApp assistant for Kesri Nandan Spares, India bike parts shop. Reply in ${tMap[waTone]}.\n${sMap[scenario]}\n3-5 lines. End with shop name. No markdown.`}]})});
      const d=await res.json();setReply(d.content?.[0]?.text||"Error");}catch{setReply("Network error.");}
      setWaLoading(false);
    };
    const addReminder=()=>{if(!nr.title)return;const u=[...reminders,{id:Date.now(),...nr,done:false}];setReminders(u);save({reminders:u});setRemSheet(false);setNr({title:"",cat:"Payment",time:"10:00",repeat:"once"});};
    const markAttendance=(sid,status)=>{const u=staff.map(s=>s.id===sid?{...s,attendance:{...s.attendance,[today]:status}}:s);setStaff(u);save({staff:u});};
    return(
      <div>
        <div style={{display:"flex",gap:6,marginBottom:14,overflowX:"auto",scrollbarWidth:"none"}}>
          {[["📱 "+t.whatsapp,"wa"],["🔔 "+t.reminders,"rem"],["👥 "+t.customers,"cust"],["👨‍💼 "+t.staff,"stf"]].map(([l,v])=>(<Pill key={v} active={sub===v} color={T.saffron} onClick={()=>setSub(v)}>{l}</Pill>))}
        </div>
        {sub==="wa"&&(<>
          <div style={{display:"flex",gap:6,overflowX:"auto",marginBottom:12,scrollbarWidth:"none"}}>
            {[{id:"general",l:"💬 General"},{id:"available",l:"✅ Available"},{id:"notavail",l:"❌ Not Avail"},{id:"price",l:"💰 Price"},{id:"complaint",l:"😤 Complaint"},{id:"promo",l:"🎉 Offer"}].map(s=>(<Pill key={s.id} active={scenario===s.id} color={T.green} onClick={()=>setScenario(s.id)}>{s.l}</Pill>))}
          </div>
          <div style={{display:"flex",gap:6,marginBottom:14,flexWrap:"wrap"}}>
            {[{id:"hinglish",l:"🇮🇳 Hinglish"},{id:"hindi",l:"🙏 Hindi"},{id:"english",l:"🌐 English"},{id:"formal",l:"👔 Formal"}].map(tn=>(<Pill key={tn.id} active={waTone===tn.id} color={T.blue} onClick={()=>setWaTone(tn.id)}>{tn.l}</Pill>))}
          </div>
          <Card>
            <div style={{fontSize:"0.62rem",color:T.textMd,textTransform:"uppercase",letterSpacing:1,fontWeight:700,marginBottom:8}}>Customer Message / Topic</div>
            <textarea value={custMsg} onChange={e=>setCustMsg(e.target.value)} placeholder="e.g. Bhai Pulsar 150 ka air filter milega?"
              style={{width:"100%",padding:"11px 13px",borderRadius:10,border:`1.5px solid ${T.border}`,background:T.surface,color:T.text,fontSize:"0.83rem",resize:"none",height:76,outline:"none",marginBottom:10,boxSizing:"border-box",fontFamily:"inherit",transition:"border-color 0.15s"}}
              onFocus={e=>e.target.style.borderColor=T.green} onBlur={e=>e.target.style.borderColor=T.border}/>
            <Btn full color={T.green} onClick={genReply} disabled={waLoading||!custMsg.trim()}>
              {waLoading?<span style={{display:"flex",alignItems:"center",justifyContent:"center",gap:8}}><span style={{display:"inline-block",width:14,height:14,border:`2px solid ${T.green}40`,borderTopColor:T.green,borderRadius:"50%",animation:"spin 0.6s linear infinite"}}/>Generating...</span>:"🤖 Generate Reply"}
            </Btn>
          </Card>
          {reply&&(<Card accent={T.green}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}><span style={{fontWeight:700,color:T.green,fontSize:"0.81rem"}}>✅ Reply Ready</span><Btn sm color={T.green} onClick={()=>navigator.clipboard.writeText(reply).then(()=>alert("Copied!")).catch(()=>{})}>📋 Copy</Btn></div>
            <div style={{background:T.greenLt,borderRadius:10,padding:"13px",whiteSpace:"pre-wrap",fontSize:"0.84rem",lineHeight:1.7,color:"#065f46"}}>{reply}</div>
            <button onClick={()=>window.open(`https://wa.me/?text=${encodeURIComponent(reply)}`,"_blank")} style={{width:"100%",marginTop:10,padding:"10px",borderRadius:10,background:"#25D366",border:"none",color:"#fff",fontWeight:700,fontSize:"0.81rem",cursor:"pointer",boxShadow:"0 3px 10px #25D36635"}}>💬 Open in WhatsApp</button>
          </Card>)}
        </>)}
        {sub==="rem"&&(<>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
            <StatCard label="Pending" value={reminders.filter(r=>!r.done).length} icon="🔔" color={T.saffron} lt={T.saffronLt}/>
            <StatCard label="Done" value={reminders.filter(r=>r.done).length} icon="✅" color={T.green} lt={T.greenLt}/>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}><SecHead title={t.reminders}/><Btn sm color={T.saffron} onClick={()=>setRemSheet(true)}>+ Add</Btn></div>
          {reminders.filter(r=>!r.done).map(r=>(<div key={r.id} style={{display:"flex",gap:11,background:T.surface,border:`1px solid ${CAT_C[r.cat]||T.saffron}25`,borderRadius:13,padding:"13px 14px",marginBottom:9,boxShadow:T.shadow}}>
            <button onClick={()=>{const u=reminders.map(x=>x.id===r.id?{...x,done:true}:x);setReminders(u);save({reminders:u});}} style={{width:22,height:22,borderRadius:6,border:`2px solid ${CAT_C[r.cat]||T.saffron}`,background:"transparent",cursor:"pointer",flexShrink:0,marginTop:2}}/>
            <div style={{flex:1}}><div style={{fontWeight:600,fontSize:"0.84rem",color:T.text,marginBottom:5}}>{r.title}</div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap"}}><Badge color={CAT_C[r.cat]||T.saffron} lt={`${CAT_C[r.cat]||T.saffron}15`}>{CAT_I[r.cat]} {r.cat}</Badge><Badge color={T.textSm} lt={T.hi2}>⏰ {r.time}</Badge></div></div>
            <button onClick={()=>{const u=reminders.filter(x=>x.id!==r.id);setReminders(u);save({reminders:u});}} style={{background:T.redLt,border:"none",color:T.red,width:26,height:26,borderRadius:7,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>×</button>
          </div>))}
          {reminders.filter(r=>!r.done).length===0&&<div style={{textAlign:"center",padding:"28px 0",color:T.textSm}}><div style={{fontSize:"2rem",marginBottom:8}}>✅</div>All tasks done!</div>}
          <Sheet title={t.addReminder} open={remSheet} onClose={()=>setRemSheet(false)}>
            <Inp label="Title *" placeholder="e.g. Collect payment from Ajay" value={nr.title} onChange={e=>setNr({...nr,title:e.target.value})}/>
            <div style={{marginBottom:12}}><div style={{fontSize:"0.62rem",color:T.textMd,textTransform:"uppercase",letterSpacing:1,fontWeight:700,marginBottom:8}}>Category</div><div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{Object.entries(CAT_C).map(([c,col])=>(<Pill key={c} active={nr.cat===c} color={col} onClick={()=>setNr({...nr,cat:c})}>{CAT_I[c]} {c}</Pill>))}</div></div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <Inp label={t.time} type="time" value={nr.time} onChange={e=>setNr({...nr,time:e.target.value})}/>
              <div><div style={{fontSize:"0.62rem",color:T.textMd,textTransform:"uppercase",letterSpacing:1,fontWeight:700,marginBottom:8}}>{t.repeat}</div><select value={nr.repeat} onChange={e=>setNr({...nr,repeat:e.target.value})} style={{width:"100%",padding:"10px 12px",borderRadius:10,border:`1.5px solid ${T.border}`,background:T.surface,color:T.text,fontSize:"0.83rem",outline:"none"}}><option value="once">Once</option><option value="daily">Daily</option><option value="weekly">Weekly</option></select></div>
            </div>
            <Btn full color={T.saffron} onClick={addReminder} style={{marginTop:8}}>🔔 Set Reminder</Btn>
          </Sheet>
        </>)}
        {sub==="cust"&&(<>
          <SecHead title={`${t.customers} (${customers.length})`} action={<Btn sm ghost color={T.purple}>⭐ Loyalty</Btn>}/>
          {[...customers].sort((a,b)=>b.totalBuy-a.totalBuy).map(c=>(<Card key={c.id} style={{padding:"14px 16px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
              <div style={{display:"flex",gap:11,alignItems:"center"}}>
                <div style={{width:40,height:40,borderRadius:11,background:T.blueLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1.2rem",flexShrink:0}}>👤</div>
                <div><div style={{fontWeight:700,fontSize:"0.89rem",color:T.text}}>{c.name}</div><div style={{fontSize:"0.65rem",color:T.textSm,marginTop:2}}>🏍️ {c.bike} · 📍 {c.area}</div>
                <div style={{marginTop:4,display:"flex",gap:5}}><Badge color={T.purple} lt={T.purpleLt}>⭐ {c.points} pts</Badge>{c.points>=100&&<Badge color={T.green} lt={T.greenLt}>Reward!</Badge>}</div></div>
              </div>
              <div style={{textAlign:"right"}}><div style={{fontWeight:800,color:T.saffronDk,fontSize:"0.89rem"}}>₹{c.totalBuy.toLocaleString("en-IN")}</div><div style={{fontSize:"0.63rem",color:T.textSm}}>{c.visits} visits</div></div>
            </div>
            <div style={{display:"flex",gap:8}}>
              <a href={`tel:${c.phone}`} style={{flex:1,padding:"8px",borderRadius:9,background:T.blueLt,border:`1px solid ${T.blue}20`,color:T.blue,textDecoration:"none",textAlign:"center",fontSize:"0.73rem",fontWeight:600}}>📞 Call</a>
              <a href={`https://wa.me/91${c.phone}`} target="_blank" rel="noreferrer" style={{flex:1,padding:"8px",borderRadius:9,background:T.greenLt,border:`1px solid ${T.green}20`,color:T.green,textDecoration:"none",textAlign:"center",fontSize:"0.73rem",fontWeight:600}}>💬 WhatsApp</a>
            </div>
          </Card>))}
        </>)}
        {sub==="stf"&&(<>
          <SecHead title={t.staffManager}/>
          {staff.map(s=>{const ts=s.attendance[today];return(<Card key={s.id} style={{padding:"14px 16px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
              <div style={{display:"flex",gap:11,alignItems:"center"}}>
                <div style={{width:40,height:40,borderRadius:11,background:T.purpleLt,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1.2rem",flexShrink:0}}>👨‍💼</div>
                <div><div style={{fontWeight:700,fontSize:"0.89rem",color:T.text}}>{s.name}</div><div style={{fontSize:"0.65rem",color:T.textSm,marginTop:2}}>{s.role} · 📱 {s.phone}</div><Badge color={T.saffron} lt={T.saffronLt} style={{marginTop:4}}>₹{s.salary.toLocaleString("en-IN")}/mo</Badge></div>
              </div>
              <div>{ts?<Badge color={ts==="present"?T.green:T.red} lt={ts==="present"?T.greenLt:T.redLt}>{ts==="present"?t.present:t.absent}</Badge>:<Badge color={T.textSm} lt={T.hi2}>Not marked</Badge>}</div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:7,marginBottom:8}}>
              <button onClick={()=>markAttendance(s.id,"present")} style={{padding:"8px",borderRadius:9,border:`1.5px solid ${T.green}`,background:ts==="present"?T.greenLt:"transparent",color:T.green,fontSize:"0.72rem",fontWeight:600,cursor:"pointer"}}>✅ {t.present}</button>
              <button onClick={()=>markAttendance(s.id,"absent")} style={{padding:"8px",borderRadius:9,border:`1.5px solid ${T.red}`,background:ts==="absent"?T.redLt:"transparent",color:T.red,fontSize:"0.72rem",fontWeight:600,cursor:"pointer"}}>❌ {t.absent}</button>
              <a href={`tel:${s.phone}`} style={{padding:"8px",borderRadius:9,border:`1.5px solid ${T.blue}`,background:T.blueLt,color:T.blue,textDecoration:"none",textAlign:"center",fontSize:"0.72rem",fontWeight:600}}>📞 Call</a>
            </div>
            <div style={{fontSize:"0.7rem",color:T.textSm}}>This month: {Object.values(s.attendance).filter(v=>v==="present").length} present / {Object.values(s.attendance).filter(v=>v==="absent").length} absent</div>
          </Card>);})}
        </>)}
      </div>
    );
  };

  const SCREENS=[<Dashboard/>,<Analytics/>,<SalesTab/>,<StockTab/>,<Finance/>,<Tools/>];

  return(
    <>
      <style>{`*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent}body{background:${T.bg}}input,textarea,select{-webkit-appearance:none;font-family:inherit}input::placeholder,textarea::placeholder{color:${T.textSm}}::-webkit-scrollbar{width:3px;height:3px}::-webkit-scrollbar-thumb{background:${T.border};border-radius:2px}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      <div style={{fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif",minHeight:"100vh",background:T.bg,color:T.text}}>
        {showSearch&&<SearchOverlay/>}
        {/* HEADER */}
        <div style={{background:T.surface,borderBottom:`1px solid ${T.border}`,padding:"12px 16px",position:"sticky",top:0,zIndex:50,boxShadow:"0 1px 8px rgba(0,0,0,0.05)"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",maxWidth:520,margin:"0 auto"}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:35,height:35,borderRadius:9,background:`linear-gradient(135deg,${T.saffron},${T.saffronDk})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"1.05rem",boxShadow:`0 3px 10px ${T.saffron}40`}}>🏍️</div>
              <div>
                <div style={{fontWeight:900,fontSize:"0.92rem",color:T.text,lineHeight:1.1}}>{t.appName}</div>
                <div style={{fontSize:"0.56rem",color:T.saffronDk,fontWeight:700,letterSpacing:"0.5px"}}>{t.appSub}</div>
              </div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:7}}>
              <button onClick={()=>{const nl=lang==="en"?"hi":"en";setLang(nl);save({lang:nl});}} style={{padding:"5px 11px",borderRadius:20,border:`1.5px solid ${T.saffron}40`,background:T.saffronXLt,color:T.saffronDk,fontSize:"0.67rem",fontWeight:700,cursor:"pointer"}}>
                {lang==="en"?"🇮🇳 हिंदी":"🌐 ENG"}
              </button>
              <button onClick={()=>setShowSearch(true)} style={{width:32,height:32,borderRadius:9,background:T.hi2,border:`1px solid ${T.border}`,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",fontSize:"0.9rem",color:T.textMd}}>🔍</button>
              {lowStockItems.length>0&&<div style={{background:T.redLt,border:`1px solid ${T.red}25`,borderRadius:20,padding:"3px 9px",fontSize:"0.62rem",color:T.red,fontWeight:700}}>⚠️ {lowStockItems.length}</div>}
            </div>
          </div>
        </div>
        {/* CONTENT */}
        <div style={{padding:"14px 14px 84px",maxWidth:520,margin:"0 auto"}} key={tab}>{SCREENS[tab]}</div>
        {/* BOTTOM NAV */}
        <div style={{position:"fixed",bottom:0,left:0,right:0,background:T.surface,borderTop:`1px solid ${T.border}`,boxShadow:"0 -4px 20px rgba(0,0,0,0.05)",zIndex:100}}>
          <div style={{display:"flex",maxWidth:520,margin:"0 auto"}}>
            {[["🏠",t.home],["📊",t.analytics],["⚡",t.pos],["📦",t.stock],["💰",t.finance],["🛠️",t.tools]].map(([ic,lb],i)=>(
              <button key={i} onClick={()=>setTab(i)} style={{flex:1,padding:"9px 4px 11px",background:"none",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
                <div style={{width:28,height:28,borderRadius:8,background:tab===i?T.saffronLt:"transparent",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"0.92rem",transition:"all 0.14s"}}>{ic}</div>
                <span style={{fontSize:"0.51rem",fontWeight:tab===i?700:500,color:tab===i?T.saffronDk:T.textSm,letterSpacing:"0.2px"}}>{lb}</span>
                {tab===i&&<div style={{width:14,height:2.5,borderRadius:2,background:T.saffron}}/>}
              </button>
            ))}
          </div>
        </div>
      </div>
    </>
  );
}
