const SHEETS = {
  SETTINGS:'Settings', VALUES:'Values', DEPARTMENTS:'Departments', DEVICES:'Devices', INSURANCE:'Insurance', PARTNERS:'Partners', USERS:'Users'
};

function doGet(e) {
  const p=e && e.parameter ? e.parameter : {};
  if (String(p.admin||'')==='1') return HtmlService.createTemplateFromFile('Admin').evaluate().setTitle('إدارة موقع أطباء الستين').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  const payload={ok:true,data:getPublicData()};
  const body=JSON.stringify(payload);
  if(p.callback) return ContentService.createTextOutput(`${p.callback}(${body})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e){
  try{
    const p=e.parameter||{};
    if(p.action==='login') return json_({ok:checkLogin_(p.username,p.password)});
    if(p.action==='save'){
      if(!checkToken_(p.token)) return json_({ok:false,error:'Unauthorized'});
      const payload=JSON.parse(p.payload||'{}'); saveAll_(payload); return json_({ok:true});
    }
    return json_({ok:false,error:'Unknown action'});
  }catch(err){return json_({ok:false,error:String(err.message||err)})}
}
function json_(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)}
function ss_(){return SpreadsheetApp.getActiveSpreadsheet()}
function sheet_(name){return ss_().getSheetByName(name)}
function rows_(name){const s=sheet_(name);if(!s)return[];const v=s.getDataRange().getValues();if(v.length<2)return[];const h=v.shift().map(String);return v.filter(r=>r.some(x=>x!==''&&x!==null)).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]])))}
function getPublicData(){
 const set=Object.fromEntries(rows_(SHEETS.SETTINGS).map(r=>[r.key,r.value]));
 return {
   site:{whatsapp:String(set.whatsapp||'966540394555'),address_ar:String(set.address_ar||'الطائف - شارع الستين'),address_en:String(set.address_en||'Taif - 60th Street'),map_url:String(set.map_url||'https://maps.app.goo.gl/5cZ4MrckUV79ZdHf9'),map_embed:String(set.map_embed||'https://www.google.com/maps?q=21.4427328,40.483089&z=17&output=embed')},
   values:rows_(SHEETS.VALUES),
   departments:rows_(SHEETS.DEPARTMENTS),
   devices:rows_(SHEETS.DEVICES),
   insurance:rows_(SHEETS.INSURANCE).filter(r=>String(r.active)!=='false').map(r=>String(r.name)),
   partners:rows_(SHEETS.PARTNERS).filter(r=>String(r.active)!=='false')
 };
}
function setupSheets(){
 const ss=ss_();
 const defs={
  Settings:[['key','value'],['whatsapp','966540394555'],['address_ar','الطائف - شارع الستين'],['address_en','Taif - 60th Street'],['map_url','https://maps.app.goo.gl/5cZ4MrckUV79ZdHf9'],['map_embed','https://www.google.com/maps?q=21.4427328,40.483089&z=17&output=embed']],
  Values:[['icon','ar','en','ar_desc','en_desc'],['✦','الجودة','Quality','تقديم الرعاية الصحية بمستوى مهني متميز.','Professional healthcare delivered to a high standard.'],['♥','المريض أولاً','Patient First','احتياجات المريض محور الخدمة والرعاية.','Patient needs are at the center of our care.'],['◎','العمل كفريق','Teamwork','تكامل الخبرات لتقديم تجربة صحية أفضل.','Integrated expertise for better care.'],['◇','الاحترام','Respect','احترام جميع المرضى والزائرين والمحافظة على كرامتهم.','Respect for every patient and visitor.']],
  Departments:[['ar','en','desc_ar','desc_en','active'],['عيادة الباطنة','Internal Medicine','تشخيص وعلاج ومتابعة الأمراض الباطنية والمزمنة.','Diagnosis and follow-up of internal and chronic conditions.',true],['عيادة الأسنان','Dental Clinic','خدمات أسنان علاجية وتجميلية متكاملة.','Integrated restorative and cosmetic dental care.',true],['عيادة الطب العام','General Medicine','رعاية أولية وخدمات طبية عامة للمرضى.','Primary and general medical care.',true],['عيادة طب الأطفال','Pediatrics','رعاية الطفل ومتابعة صحته ونموه.','Child healthcare and growth follow-up.',true],['الطوارئ','Emergency','استقبال الحالات الطارئة وتقديم الرعاية الأولية.','Initial care for urgent and emergency cases.',true],['عيادة الجلدية','Dermatology','الجلدية والتجميل والليزر بأجهزة وتقنيات حديثة.','Dermatology, aesthetics and laser technologies.',true],['عيادة النساء والولادة','Obstetrics & Gynecology','متابعة صحة المرأة والحمل والولادة.','Women’s health, pregnancy and obstetric care.',true],['فحص العمالة','Medical Screening','الفحص الطبي للعمالة والفحوصات المطلوبة.','Medical screening and required examinations.',true],['قسم المختبر','Laboratory','تحاليل مخبرية باستخدام أجهزة حديثة.','Laboratory testing with modern equipment.',true],['قسم الأشعة','Radiology','خدمات الأشعة والتصوير الطبي.','Diagnostic imaging and radiology services.',true],['الصيدليات','Pharmacies','أدوية ومستلزمات طبية وخدمة حسب الفرع.','Medicines and medical supplies.',true],['شركات التأمين','Insurance','خدمة حاملي بطاقات شركات التأمين المعتمدة.','Support for approved insurance card holders.',true]],
  Devices:[['name','ar','en','active'],['Spectra XT','تقنيات الليزر والعناية بالبشرة','Laser and skin-care technology',true],['GentleMax Pro Plus','تقنيات متقدمة لإزالة الشعر والليزر','Advanced hair removal and laser technology',true],['Action II','تقنيات تجميل وتجديد البشرة','Aesthetic and skin rejuvenation technology',true],['Clarity II','ليزر ثنائي الموجة لتطبيقات الجلدية','Dual-wavelength dermatology laser',true]],
  Insurance:[['name','active'],['GlobeMed',true],['Saudi Enaya',true],['Aljazira Takaful',true],['MedGulf',true],['Malath',true],['Walaa',true],['NextCare',true],['Tawuniya',true],['TUC',true],['AXA',true],['TCS',true],['Liva',true],['Cigna',true],['Bupa Arabia',true]],
  Partners:[['ar','en','logo','active'],['شركة أبناء عبدالله باغفار المحدودة','Abdullah A. Baghffar Sons Co., Ltd.','assets/partners/partner-01.jpg',true],['شركة مكتب الكمال للاستيراد المحدودة','Alkamal Import Office Co.','assets/partners/partner-02.jpg',true],['شركة إمداد للأعمال الطبية','Imdad Medical Business','assets/partners/partner-03.jpg',true],['الشركة الرابية الحديثة للتجارة','Modern Rabiyah Trading Company','assets/partners/partner-04.jpg',true],['الشركة السعودية الخليجية سيبكو','SEPCO Environment','assets/partners/partner-05.jpg',true],['شركة بسام بن محمد الشقراوي وشركاؤه للاستشارات المهنية','BMB & Co Advisory & Assurance','assets/partners/partner-06.jpg',true],['شركة الخدمات الطبية والصيدلانية بشرى شكيب','Medical & Pharmaceutical Services','assets/partners/partner-07.jpg',true],['شركة مختبر الجودة الطبي','Aljawadah Medical Laboratory Company','assets/partners/partner-08.jpg',true],['شركة صفوة التشخيص الطبية','Diagnostics Elite','assets/partners/partner-09.jpg',true],['مجمع جمجوم','Jamjoom Medical Solutions','assets/partners/partner-10.jpg',true],['شركة العرفج للتجارة والاستيراد','Alarfaj Commercial & Imports Co.','assets/partners/partner-11.jpg',true],['شركة المرمر الأبيض للخدمات الطبية','Al Marmar Dental Co.','assets/partners/partner-12.jpg',true]],
  Users:[['username','password_hash'],['admin',hash_('ChangeMe123!')]]
 };
 Object.keys(defs).forEach(name=>{let s=ss.getSheetByName(name);if(!s)s=ss.insertSheet(name);s.clear();s.getRange(1,1,defs[name].length,defs[name][0].length).setValues(defs[name]);s.setFrozenRows(1);s.autoResizeColumns(1,defs[name][0].length)});
 PropertiesService.getScriptProperties().setProperty('ADMIN_TOKEN', Utilities.getUuid());
}
function hash_(s){const b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s));return b.map(x=>(x+256)%256).map(x=>x.toString(16).padStart(2,'0')).join('')}
function checkLogin_(u,p){const users=rows_(SHEETS.USERS);const ok=users.some(x=>String(x.username)===String(u)&&String(x.password_hash)===hash_(p));return ok?PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN'):false}
function checkToken_(t){return t&&t===PropertiesService.getScriptProperties().getProperty('ADMIN_TOKEN')}
function saveAll_(payload){
 const map={values:SHEETS.VALUES,departments:SHEETS.DEPARTMENTS,devices:SHEETS.DEVICES,insurance:SHEETS.INSURANCE,partners:SHEETS.PARTNERS};
 Object.keys(map).forEach(k=>{if(Array.isArray(payload[k])) writeObjects_(map[k],payload[k])});
 if(payload.site){const s=sheet_(SHEETS.SETTINGS);const obj=payload.site;const rows=[['key','value'],...Object.keys(obj).map(k=>[k,obj[k]])];s.clear();s.getRange(1,1,rows.length,2).setValues(rows);s.setFrozenRows(1)}
}
function writeObjects_(name,arr){const s=sheet_(name);const current=s.getDataRange().getValues();const headers=current[0].map(String);const values=[headers,...arr.map(o=>headers.map(h=>o[h]??''))];s.clear();s.getRange(1,1,values.length,headers.length).setValues(values);s.setFrozenRows(1)}


function upgradeV2(){
 const ss=ss_();
 let settings=ss.getSheetByName(SHEETS.SETTINGS);
 if(settings){
   const vals=settings.getDataRange().getValues();
   const keys=vals.slice(1).map(r=>String(r[0]));
   const add=[];
   if(!keys.includes('map_url')) add.push(['map_url','https://maps.app.goo.gl/5cZ4MrckUV79ZdHf9']);
   if(!keys.includes('map_embed')) add.push(['map_embed','https://www.google.com/maps?q=21.4427328,40.483089&z=17&output=embed']);
   if(add.length) settings.getRange(settings.getLastRow()+1,1,add.length,2).setValues(add);
 }
 let p=ss.getSheetByName(SHEETS.PARTNERS);
 if(!p){
   p=ss.insertSheet(SHEETS.PARTNERS);
   const rows=[['ar','en','logo','active'],['شركة أبناء عبدالله باغفار المحدودة','Abdullah A. Baghffar Sons Co., Ltd.','assets/partners/partner-01.jpg',true],['شركة مكتب الكمال للاستيراد المحدودة','Alkamal Import Office Co.','assets/partners/partner-02.jpg',true],['شركة إمداد للأعمال الطبية','Imdad Medical Business','assets/partners/partner-03.jpg',true],['الشركة الرابية الحديثة للتجارة','Modern Rabiyah Trading Company','assets/partners/partner-04.jpg',true],['الشركة السعودية الخليجية سيبكو','SEPCO Environment','assets/partners/partner-05.jpg',true],['شركة بسام بن محمد الشقراوي وشركاؤه للاستشارات المهنية','BMB & Co Advisory & Assurance','assets/partners/partner-06.jpg',true],['شركة الخدمات الطبية والصيدلانية بشرى شكيب','Medical & Pharmaceutical Services','assets/partners/partner-07.jpg',true],['شركة مختبر الجودة الطبي','Aljawadah Medical Laboratory Company','assets/partners/partner-08.jpg',true],['شركة صفوة التشخيص الطبية','Diagnostics Elite','assets/partners/partner-09.jpg',true],['مجمع جمجوم','Jamjoom Medical Solutions','assets/partners/partner-10.jpg',true],['شركة العرفج للتجارة والاستيراد','Alarfaj Commercial & Imports Co.','assets/partners/partner-11.jpg',true],['شركة المرمر الأبيض للخدمات الطبية','Al Marmar Dental Co.','assets/partners/partner-12.jpg',true]];
   p.getRange(1,1,rows.length,rows[0].length).setValues(rows); p.setFrozenRows(1); p.autoResizeColumns(1,rows[0].length);
 }
}
