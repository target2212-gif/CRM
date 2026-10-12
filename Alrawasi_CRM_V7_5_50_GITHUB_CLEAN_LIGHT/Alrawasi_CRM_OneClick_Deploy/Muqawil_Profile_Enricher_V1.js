(async()=>{
'use strict';
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
const delay=()=>3500+Math.floor(Math.random()*3000);
const pickFile=()=>new Promise((resolve,reject)=>{const i=document.createElement('input');i.type='file';i.accept='.json,application/json';i.onchange=async()=>{try{resolve(JSON.parse(await i.files[0].text()))}catch(e){reject(e)}};i.click()});
const txt=s=>(s||'').replace(/\s+/g,' ').trim();
const valueAfter=(root,label)=>{
 const els=[...root.querySelectorAll('div,p,span,li,td,th,strong,b,label')];
 for(const e of els){const t=txt(e.innerText); if(t===label||t.startsWith(label+':')||t.startsWith(label+'：')){
   let v=t.replace(label,'').replace(/^\s*[:：]\s*/,'').trim(); if(v)return v;
   const n=e.nextElementSibling; if(n&&txt(n.innerText))return txt(n.innerText);
   const p=e.parentElement; if(p){const all=txt(p.innerText); const z=all.replace(label,'').replace(/^\s*[:：]\s*/,'').trim(); if(z&&z!==all)return z;}
 }} return '';
};
const rx=(body,re)=>{const m=body.match(re);return m?txt(m[1]):''};
const parse=html=>{
 const doc=new DOMParser().parseFromString(html,'text/html'), body=txt(doc.body.innerText);
 const email=(body.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)||[])[0]||'';
 const phones=[...new Set((body.match(/(?:\+?966\s*[- ]?)?(?:0?1\d|0?5\d)(?:[\s-]?\d){7,8}/g)||[]).map(txt))];
 const mobile=phones.find(x=>/^(?:\+?966\s*[- ]?)?0?5/.test(x.replace(/\s/g,'')))||'';
 const land=phones.find(x=>/^(?:\+?966\s*[- ]?)?0?1/.test(x.replace(/\s/g,'')))||'';
 const grade=rx(body,/مصنف\s*(?:-|–|—)?\s*درجة\s*([^\s،|]+)/i);
 return {
  membershipType:valueAfter(doc,'العضوية')||rx(body,/العضوية\s*[:：]?\s*([^\n|•]{2,40})/),
  membershipSince:valueAfter(doc,'عضو منذ')||rx(body,/عضو منذ\s*[:：]?\s*([0-9/\-]{6,12})/),
  companySize:valueAfter(doc,'حجم المنشأة')||rx(body,/حجم المنشأة\s*[:：]?\s*([^\n|•]{2,60})/),
  accountStatus:valueAfter(doc,'حالة الحساب')||'',
  trainingHours:(valueAfter(doc,'عدد الساعات التدريبية')||rx(body,/عدد الساعات التدريبية\s*[:：]?\s*([0-9]+)/)).replace(/[^0-9]/g,''),
  region:valueAfter(doc,'المنطقة'), city:valueAfter(doc,'المدينة'), address:valueAfter(doc,'عنوان')||valueAfter(doc,'العنوان'),
  email, phoneMobile:mobile, phoneLandline:land,
  classificationStatus: grade?'مصنفة':(body.includes('غير مصنف')?'غير مصنفة':''), classificationGrade:grade
 };
};
if(location.hostname!=='muqawil.org'&&location.hostname!=='www.muqawil.org'){alert('شغّل هذا الكود من صفحة muqawil.org بعد تسجيل الدخول العادي.');return}
const q=await pickFile(); const list=q.companies||[]; if(!list.length){alert('ملف قائمة الإثراء فارغ');return}
let out=[],fail=[]; console.log('بدء إثراء',list.length,'شركة. لا تغلق الصفحة.');
for(let i=0;i<list.length;i++){
 const c=list[i], url=c.muqawil_profile_url||c.source_url; if(!url)continue;
 try{
  const u=new URL(url,location.origin); const r=await fetch(u.pathname+u.search,{credentials:'include',cache:'no-store'});
  if(r.status===403||r.status===429){console.warn('توقف آمن HTTP',r.status,'عند',i+1);break}
  if(!r.ok)throw new Error('HTTP '+r.status);
  const data=parse(await r.text()); out.push({muqawilKey:c.muqawil_key,profileUrl:u.href,...data});
  console.log(`[${i+1}/${list.length}]`,c.name_ar||c.muqawil_key,data.phoneLandline||data.phoneMobile||'بدون هاتف');
 }catch(e){fail.push({muqawilKey:c.muqawil_key,url,error:String(e)});console.warn('فشل',i+1,e)}
 if((i+1)%200===0){const blob=new Blob([JSON.stringify({type:'muqawil_profile_enrichment',companies:out,failed:fail},null,2)],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=`muqawil_enrichment_backup_${out.length}.json`;a.click();URL.revokeObjectURL(a.href)}
 await sleep(delay());
}
const blob=new Blob([JSON.stringify({type:'muqawil_profile_enrichment',exportedAt:new Date().toISOString(),companies:out,failed:fail},null,2)],{type:'application/json'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=`muqawil_enrichment_${out.length}.json`;a.click();URL.revokeObjectURL(a.href);console.log('انتهى',out.length,'فشل',fail.length);
})();
