const DEFAULT_TEXT = "مع التقدم التكنولوجي، والانتشار الكبير لعالم الإنترنت ومواقع التواصل الاجتماعي، أصبح تعلم استعمال الحاسوب ضرورة لا يمكن الاستغناء عنها في جميع أقطار العالم، وهو ما يترتب عليه ضرورة تعلم استخدام ملحقات الحاسوب كالفأرة ولوحة المفاتيح.\n\nاستخدام الفأرة سهل ولا يتطلب مهارة خاصة، لكن على العكس تماما تتطلب لوحة المفاتيح تعلما وتدريبا خاصا لإتقان الكتابة عليها، وذلك لما تحتويه من المفاتيح من جهة، ولضرورة التمرن على الطباعة بشكل أسرع من جهة أخرى.\n\nولعل البعض يتساءل عن أهمية التمرن على الكتابة باللمس والفائدة من ذلك، والجواب هنا هو أن الكثير من الأفراد بحاجة ماسة لهذه المهارة نظرا لأن طبيعة عملهم تتطلب السرعة في الكتابة على لوحة المفاتيح مثل المحررين الصحفيين، أو المدونين، أو الكتاب وغيرهم. كما أنها تعد عاملا أساسيا في عملية إدخال المعلومات والنصوص والرسائل دون الحاجة إلى النظر بشكل مستمر في لوحة المفاتيح، وبالتالي اختصار الوقت والجهد.\n\nعلى هذا الأساس، ظهرت بعض المواقع العربية التي تهدف إلى تمكين المستخدمين من فن الطباعة باللمس، ولعل أبرز المواقع العربية التي تقدم هذه.";
const DEFAULT_QUESTIONS = [["ما الهدف الأساسي من إدخال البيانات؟", ["تنظيم المعلومات", "حذف الملفات", "إيقاف الجهاز", "تغيير الشاشة"], 0], ["أي مما يلي يساعد على تقليل أخطاء الكتابة؟", ["السرعة دون مراجعة", "التركيز والمراجعة", "استخدام مفاتيح عشوائية", "إغلاق البرنامج"], 1], ["أين يفضل حفظ الملفات؟", ["في مجلدات منظمة", "على سطح المكتب فقط", "دون أسماء", "في سلة المحذوفات"], 0], ["أي برنامج يستخدم لمعالجة النصوص؟", ["برنامج تحرير النصوص", "الآلة الحاسبة", "مشغل الصوت", "عارض الصور"], 0], ["ما فائدة التدريب المستمر؟", ["تقليل المهارة", "تحسين السرعة والدقة", "حذف البيانات", "إيقاف الحفظ"], 1], ["تنظيم الملفات يسهل الوصول إليها.", true], ["مراجعة البيانات تزيد احتمالية اكتشاف الأخطاء.", true], ["يمكن مشاركة كلمات المرور مع الجميع.", false], ["الدقة مهمة عند إدخال البيانات.", true], ["حفظ الملفات دون أسماء واضحة هو أفضل أسلوب.", false]];
const KEY='typingExamApp_v1';
const $=id=>document.getElementById(id);
const clone=x=>JSON.parse(JSON.stringify(x));
const initial={text:DEFAULT_TEXT,questions:DEFAULT_QUESTIONS,pin:'1234',officeConfig:{word:{totalMarks:100,duration:20},excel:{totalMarks:100,duration:20}},officeContent:{word:'اكتب هنا نص اختبار Word المطلوب تنسيقه.',excel:'المنتج | الكمية | السعر\nقلم | 10 | 5\nدفتر | 7 | 12'},excelFiles:[],excelModels:[],officeAssets:{word:{template:[],solution:[]},excel:{template:[],solution:[]}},officeTasks:[{"id": "w1", "app": "word", "title": "كتابة عبارة محددة", "prompt": "اكتب عبارة التقنية الحديثة داخل المستند.", "kind": "containsText", "target": "", "expected": "التقنية الحديثة", "points": 20}, {"id": "w2", "app": "word", "title": "تنسيق نص غامق", "prompt": "اجعل جزءاً من النص بخط غامق.", "kind": "bold", "target": "", "expected": "", "points": 20}, {"id": "w3", "app": "word", "title": "إدراج جدول", "prompt": "أدرج جدولاً في المستند.", "kind": "table", "target": "", "expected": "", "points": 20}, {"id": "w4", "app": "word", "title": "عنوان المستند", "prompt": "اكتب عنوان مهارات الحاسوب في المستند.", "kind": "containsText", "target": "", "expected": "مهارات الحاسوب", "points": 20}, {"id": "w5", "app": "word", "title": "توسيط فقرة", "prompt": "وسّط عنواناً أو فقرة في المستند.", "kind": "centered", "target": "", "expected": "", "points": 20}, {"id": "e1", "app": "excel", "title": "إدخال قيمة", "prompt": "أدخل القيمة 120 في الخلية B2.", "kind": "cellEquals", "target": "B2", "expected": "120", "points": 25}, {"id": "e2", "app": "excel", "title": "إدخال قيمة", "prompt": "أدخل القيمة 180 في الخلية B3.", "kind": "cellEquals", "target": "B3", "expected": "180", "points": 25}, {"id": "e3", "app": "excel", "title": "حساب المجموع", "prompt": "استخدم دالة SUM في الخلية E2 لجمع B2 إلى B5.", "kind": "formulaEquals", "target": "E2", "expected": "=SUM(B2:B5)", "points": 25}, {"id": "e4", "app": "excel", "title": "حساب المتوسط", "prompt": "استخدم دالة AVERAGE في الخلية E3 لحساب متوسط B2 إلى B5.", "kind": "formulaEquals", "target": "E3", "expected": "=AVERAGE(B2:B5)", "points": 25}],settings:{title:'منصة اختبار إدخال البيانات ومعالجة النصوص',primary:'#176b50',bg:'#f3f7f5',logo:''},results:[],unlockedTrainees:{}};
let db=loadDB();
if(!db.excelModels.length && db.excelFiles.length){
 db.excelModels=db.excelFiles.map((f,i)=>({id:'model_'+f.id,name:'نموذج '+(i+1)+' - '+f.name,fileId:f.id,duration:20,totalMarks:100}));
}
if(!db.excelModels.length && db.officeTasks.some(t=>t.app==='excel')){
 db.excelModels=[{id:'model_default',name:'النموذج الافتراضي',fileId:'',duration:db.officeConfig.excel.duration||20,totalMarks:db.officeConfig.excel.totalMarks||100}];
}
if(db.excelModels.length){
 const fallbackModel=db.excelModels[0];
 db.officeTasks.forEach(t=>{if(t.app==='excel'&&!t.formId)t.formId=fallbackModel.id});
}
let activeExcelModelId=db.excelModels[0]?.id||'';
let pendingExcelModelSlot=null;
let selectedExcelExamModel=null;
let studentName='',section='typing',timeLeft=600,interval=null,startedAt=0,typingScore=0,quizScore=0,lastResult=null,typingPassed=false;
function loadDB(){try{let saved=JSON.parse(localStorage.getItem(KEY)||'{}');const data={...clone(initial),...saved,officeConfig:{word:{...initial.officeConfig.word,...(saved.officeConfig?.word||{})},excel:{...initial.officeConfig.excel,...(saved.officeConfig?.excel||{})}},officeContent:{...initial.officeContent,...(saved.officeContent||{})},excelFiles:Array.isArray(saved.excelFiles)?saved.excelFiles:[],excelModels:Array.isArray(saved.excelModels)?saved.excelModels:[],officeAssets:saved.officeAssets||{word:{template:[],solution:[]},excel:{template:[],solution:[]}},unlockedTrainees:saved.unlockedTrainees||{}};
 data.typingTexts=Array.isArray(saved.typingTexts)?saved.typingTexts.filter(x=>x&&typeof x.text==='string'&&x.text.trim()):[];
 if(!data.typingTexts.length){const legacy=String(saved.text||initial.text).trim();data.typingTexts=[{id:'text_legacy',name:'النص الأساسي',text:legacy,createdAt:new Date().toISOString()}];}
 if(!data.typingTexts.some(x=>x.text===data.text)){data.typingTexts.unshift({id:'text_current',name:'النص الحالي',text:String(data.text||initial.text),createdAt:new Date().toISOString()});}
 return data
 }catch{const data=clone(initial);data.typingTexts=[{id:'text_default',name:'النص الأساسي',text:DEFAULT_TEXT,createdAt:new Date().toISOString()}];return data}}
let trainerPreviewMode=false,trainerPreviewPreviousName='';
let adminHasUnsavedChanges=false;
function setAdminSaveState(dirty){adminHasUnsavedChanges=!!dirty;const el=$('adminSaveState');if(el){el.textContent=dirty?'● توجد تعديلات غير محفوظة':'✓ جميع التعديلات محفوظة';el.className=dirty?'save-state is-dirty':'save-state is-saved';}}
function saveDB(){localStorage.setItem(KEY,JSON.stringify(db));setAdminSaveState(false)}
function show(id){document.querySelectorAll('.screen').forEach(x=>x.classList.remove('active'));$(id).classList.add('active')}
function applySettings(){document.documentElement.style.setProperty('--primary',db.settings.primary);document.documentElement.style.setProperty('--bg',db.settings.bg);$('brand').textContent=db.settings.title;$('titleEdit').value=db.settings.title;$('colorEdit').value=db.settings.primary;$('bgEdit').value=db.settings.bg;if(db.settings.logo)$('logo').src=db.settings.logo}
function formatTime(s){return String(Math.floor(s/60)).padStart(2,'0')+':'+String(s%60).padStart(2,'0')}
function tokens(s){return s.trim().split(/\s+/).filter(Boolean)}
function renderPassage(){const target=tokens(db.text),typed=tokens($('typed').value);let index=0;const paragraphs=db.text.trim().split(/\\n\\s*\\n/).filter(Boolean);$('passage').innerHTML=paragraphs.map(par=>'<p>'+tokens(par).map(w=>{let cls='pending';if(index<typed.length)cls=typed[index]===w?'good':'bad';index++;return `<span class="word ${cls}">${escapeHtml(w)}</span>`}).join(' ')+'</p>').join('');let correct=0,wrong=0;typed.forEach((w,i)=>{if(w===target[i])correct++;else wrong++});wrong+=Math.max(0,typed.length-target.length);$('correct').textContent=correct;$('wrong').textContent=wrong;let accuracy=typed.length?Math.round(correct/typed.length*100):0;$('accuracy').textContent=accuracy+'%';let elapsed=Math.max(1,(Date.now()-startedAt)/60000);$('speed').textContent=Math.round(typed.length/elapsed);$('progress').style.width=(target.length?Math.min(100,typed.length/target.length*100):0)+'%';$('typingGateProgress').style.width=Math.min(100,correct/160*100)+'%';$('typingGateMessage').textContent=correct>=160?'أكملت 160 كلمة صحيحة. اضغط إنهاء لتفعيل بقية الاختبارات.':'الكلمات الصحيحة: '+correct+' من 160 كلمة مطلوبة.';}
function escapeHtml(s){return s.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function stopTimer(){if(interval)clearInterval(interval);interval=null}
function startTimer(seconds,onEnd){stopTimer();timeLeft=seconds;$('timer').textContent=formatTime(timeLeft);interval=setInterval(()=>{timeLeft--;$('timer').textContent=formatTime(timeLeft);if(timeLeft<=0){stopTimer();onEnd()}},1000)}
function beginTyping(){const library=Array.isArray(db.typingTexts)&&db.typingTexts.length?db.typingTexts:[{text:db.text}];const chosen=library[Math.floor(Math.random()*library.length)];db.text=chosen.text;$('toQuiz').onclick=completeTypingGate;$('toQuiz').textContent='إنهاء الكتابة والتحقق من الاجتياز';section='typing';$('heading').textContent='اختبار الكتابة';$('part').textContent='القسم الأول';$('typing').hidden=false;$('quiz').hidden=true;$('result').hidden=true;$('typed').value='';$('typingGateMessage').textContent='الكلمات الصحيحة: 0 من 160 كلمة مطلوبة.';startedAt=Date.now();renderPassage();startTimer(600,()=>completeTypingGate(true))}
function completeTypingGate(timeExpired=false){
 if(section!=='typing')return;
 const target=tokens(db.text),typed=tokens($('typed').value);
 let correct=0;
 typed.forEach((w,i)=>{if(w===target[i])correct++});
 stopTimer();
 if(trainerPreviewMode){
  section='home';
  $('typed').value='';
  updateTraineeDashboard();
  show('traineeDashboard');
  return;
 }
 if(correct<160){
  section='home';
  $('typed').value='';
  show('home');
  alert('نعتذر، نرجو المحاولة مرة أخرى.');
  return;
 }
 typingPassed=true;
 db.unlockedTrainees[studentName.trim().toLocaleLowerCase('ar')]=true;
 saveDB();
 typingScore=30;
 $('typingGateMessage').textContent='مبروك! اجتزت اختبار الكتابة وفتحت بقية الاختبارات.';
 updateTraineeDashboard();
 show('traineeDashboard');
}
function toQuiz(){completeTypingGate(false)}
function renderQuestions(){const wrap=$('questions');wrap.innerHTML='';db.questions.forEach((q,i)=>{let d=document.createElement('div');d.className='question';let p=document.createElement('p');p.textContent=(i+1)+'. '+q[0];d.appendChild(p);let opts=document.createElement('div');opts.className='options';let choices=i<5?q[1]:['صح','خطأ'];choices.forEach((v,j)=>{let label=document.createElement('label');label.className='option';let input=document.createElement('input');input.type='radio';input.name='q'+i;input.value=j;label.append(input,document.createTextNode(v));opts.appendChild(label)});d.appendChild(opts);wrap.appendChild(d)})}
function submitQuiz(){if(section!=='quiz')return;stopTimer();let correct=0;db.questions.slice(0,10).forEach((q,i)=>{let selected=document.querySelector(`input[name="q${i}"]:checked`);let answer=i<5?q[2]:(q[1]===true?0:1);if(selected&&Number(selected.value)===answer)correct++});quizScore=Math.round(correct*2);let total=quizScore;lastResult={name:studentName,date:new Date().toLocaleString('ar-SA'),typing:typingScore,quiz:quizScore,total,category:'اختبار المعلومات'};$('total').textContent=total;$('details').innerHTML=`<div class="result-row">اجتياز الكتابة: <b>160 كلمة صحيحة</b></div><div class="result-row">درجة الأسئلة: <b>${quizScore} / 20</b></div><div class="result-row">اسم المتدرب: <b>${escapeHtml(studentName)}</b></div>`;$('typing').hidden=true;$('quiz').hidden=true;$('result').hidden=false;$('heading').textContent='النتيجة النهائية';$('timer').textContent='00:00';section='result'}
$('studentBtn').onclick=()=>{show('student');$('name').value=''};
$('trainerBtn').onclick=()=>{show('trainer');$('login').hidden=false;$('admin').hidden=true;$('pin').value=''};
document.querySelectorAll('[data-home]').forEach(b=>b.onclick=()=>{stopTimer();if(typeof stopOfficeTimer==='function')stopOfficeTimer();if(trainerPreviewMode){updateTraineeDashboard();show('traineeDashboard');}else show('home')});
function traineeKey(name){return String(name||'').trim().toLocaleLowerCase('ar')}
function updateTraineeDashboard(){renderTraineeExcelFiles();$('previewBanner').classList.toggle('active',trainerPreviewMode);const unlocked=trainerPreviewMode?true:!!db.unlockedTrainees[traineeKey(studentName)];$('welcomeTrainee').textContent=trainerPreviewMode?'هذه معاينة للمدرب — هكذا ستظهر بوابة الاختبارات للمتدرب.':'مرحباً '+studentName+' — اختر الاختبار المطلوب.';$('typingGateStatus').textContent=unlocked?'تم الاجتياز':'إلزامي أولاً';$('typingGateStatus').className=unlocked?'status-open':'status-locked';['word','excel','quiz'].forEach(k=>{$(k+'GateStatus').textContent=unlocked?'متاح':'مقفل';$(k+'GateStatus').className=unlocked?'status-open':'status-locked';$('launch'+k.charAt(0).toUpperCase()+k.slice(1)).disabled=!unlocked});$('launchTyping').textContent=unlocked?'إعادة اختبار الكتابة':'بدء الاختبار';if(trainerPreviewMode){$('typingGateStatus').textContent='معاينة';$('wordGateStatus').textContent='معاينة';$('excelGateStatus').textContent='معاينة';$('quizGateStatus').textContent='معاينة';$('previewReturn').hidden=false;}else{$('previewReturn').hidden=true;}}
$('start').onclick=()=>{studentName=$('name').value.trim();if(!studentName){alert('يرجى كتابة اسم المتدرب أولاً.');$('name').focus();return}typingPassed=!!db.unlockedTrainees[traineeKey(studentName)];updateTraineeDashboard();show('traineeDashboard')};
$('launchTyping').onclick=()=>{show('exam');beginTyping()};
$('launchWord').onclick=()=>{if(!trainerPreviewMode&&!db.unlockedTrainees[traineeKey(studentName)])return alert('اجتز اختبار الكتابة أولاً.');officeTrainee=studentName;openNativeOffice('word');};
$('launchExcel').onclick=()=>{
 if(!trainerPreviewMode&&!db.unlockedTrainees[traineeKey(studentName)])return alert('اجتز اختبار الكتابة أولاً.');
 const models=(db.excelModels||[]).slice(0,4).filter(m=>db.excelFiles.some(f=>f.id===m.fileId)&&(db.officeTasks||[]).some(t=>t.app==='excel'&&t.formId===m.id));
 if(!models.length)return alert('لا توجد نماذج Excel جاهزة. اطلب من المدرب إضافة نموذج وأسئلته.');
 selectedExcelExamModel=models[Math.floor(Math.random()*models.length)];
 officeTrainee=studentName;
 openNativeOffice('excel');
};
$('launchQuiz').onclick=()=>{if(!trainerPreviewMode&&!db.unlockedTrainees[traineeKey(studentName)])return alert('اجتز اختبار الكتابة أولاً.');section='quiz';show('exam');$('typing').hidden=true;$('quiz').hidden=false;$('result').hidden=true;$('heading').textContent='اختبار المعلومات';$('part').textContent='القسم الرابع';renderQuestions();startTimer(900,submitQuiz)};
$('typed').addEventListener('input',renderPassage);$('toQuiz').onclick=toQuiz;$('submit').onclick=submitQuiz;
$('save').onclick=()=>{if(lastResult&&!trainerPreviewMode){db.results.push(lastResult);saveDB()}stopTimer();if(trainerPreviewMode){lastResult=null;updateTraineeDashboard();show('traineeDashboard');return}show('home')};
$('unlock').onclick=()=>{if($('pin').value===db.pin){$('login').hidden=true;$('admin').hidden=false;loadAdmin();setAdminSaveState(false)}else alert('الرقم السري غير صحيح.')};
$('previewTrainee').onclick=()=>{if(adminHasUnsavedChanges){alert('لديك تعديلات لم تحفظ بعد. احفظ التعديلات أولاً ثم افتح المعاينة.');return;}trainerPreviewPreviousName=studentName;trainerPreviewMode=true;studentName='معاينة المدرب';updateTraineeDashboard();show('traineeDashboard');};
$('previewReturn').onclick=()=>{trainerPreviewMode=false;studentName=trainerPreviewPreviousName;show('trainer');$('login').hidden=true;$('admin').hidden=false;};
document.querySelectorAll('#admin input, #admin textarea, #admin select').forEach(el=>{el.addEventListener('input',()=>setAdminSaveState(true));el.addEventListener('change',()=>setAdminSaveState(true));});
document.querySelectorAll('[data-tab]').forEach(b=>b.onclick=()=>{document.querySelectorAll('.tab').forEach(x=>x.hidden=true);$(b.dataset.tab).hidden=false;document.querySelectorAll('[data-tab]').forEach(x=>x.classList.toggle('active',x===b));if(b.dataset.tab==='results')renderResults()});
function loadAdmin(){$('textEdit').value=db.text;renderTypingLibrary();renderQuestionEditor();renderOfficeTaskEditor();renderOfficeAssets();$('wordAdminSection').hidden=false;$('excelAdminSection').hidden=false;document.querySelectorAll('[data-tab]').forEach((b,i)=>b.classList.toggle('active',i===0));document.querySelectorAll('.tab').forEach((x,i)=>x.hidden=i!==0)}
function renderQuestionEditor(){let root=$('editQuestions');root.innerHTML='';db.questions.forEach((q,i)=>{let d=document.createElement('div');d.className='edit-q';let label=document.createElement('label');label.textContent=(i+1)+'. نص السؤال';let input=document.createElement('input');input.value=q[0];input.dataset.i=i;input.dataset.field='question';d.append(label,input);if(i<5){let grid=document.createElement('div');grid.className='grid';q[1].forEach((v,j)=>{let x=document.createElement('input');x.value=v;x.placeholder='الخيار '+(j+1);x.dataset.i=i;x.dataset.field='option';x.dataset.j=j;grid.appendChild(x)});let sel=document.createElement('select');sel.dataset.i=i;sel.dataset.field='answer';['الخيار الأول','الخيار الثاني','الخيار الثالث','الخيار الرابع'].forEach((v,j)=>{let o=document.createElement('option');o.value=j;o.textContent=v;sel.appendChild(o)});sel.value=q[2];let l=document.createElement('label');l.textContent='الإجابة الصحيحة';d.append(grid,l,sel)}else{let sel=document.createElement('select');sel.dataset.i=i;sel.dataset.field='tf';[['true','صح'],['false','خطأ']].forEach(([v,t])=>{let o=document.createElement('option');o.value=v;o.textContent=t;sel.appendChild(o)});sel.value=String(q[1]);let l=document.createElement('label');l.textContent='الإجابة الصحيحة';d.append(l,sel)}root.appendChild(d)})}
$('importTypingWord').onclick=async()=>{
 if(!window.desktopApp?.importTypingWord){alert('استيراد ملفات Word متاح من نسخة البرنامج المثبتة على الكمبيوتر.');return;}
 const button=$('importTypingWord');button.disabled=true;button.textContent='جارٍ استيراد الملف...';
 try {
  const result=await window.desktopApp.importTypingWord();
  if(result.canceled)return;
  if(!result.ok){alert(result.message||'تعذر استيراد ملف Word.');return;}
  const imported=String(result.content||'').replace(/\r/g,'').trim();
  const count=tokens(imported).length;
  if(!imported){alert('لم يتم العثور على نص داخل الملف.');return;}
  if(count!==160){$('typingImportStatus').textContent='تمت قراءة '+count+' كلمة من '+result.fileName+'؛ المطلوب 160 كلمة بالضبط.';alert('عدد كلمات الملف '+count+' كلمة. يجب أن يحتوي النص على 160 كلمة بالضبط.');return;}
  $('textEdit').value=imported;
  $('typingImportStatus').textContent='تم استيراد '+result.fileName+' بنجاح — 160 كلمة. راجع النص ثم اضغط «حفظ نص الكتابة». ';
  alert('تم استيراد النص بنجاح. تمت معاينته في مربع النص؛ اضغط حفظ نص الكتابة لاعتماده.');
 }catch(error){alert('تعذر استيراد الملف: '+error.message)}
 finally{button.disabled=false;button.textContent='📄 استيراد نص من Word';}
};
function renderTypingLibrary(){
 const list=$('typingTextLibrary'),count=$('typingTextCount');
 if(!list||!count)return;
 const texts=Array.isArray(db.typingTexts)?db.typingTexts:[];
 count.textContent=texts.length;
 list.innerHTML='';
 if(!texts.length){list.innerHTML='<p class="hint">لا توجد نصوص محفوظة حتى الآن.</p>';return;}
 texts.forEach((item,index)=>{
  const row=document.createElement('div');row.className='typing-library-item';
  const info=document.createElement('div');info.className='typing-library-info';
  const name=document.createElement('b');name.textContent=item.name||('نص كتابة '+(index+1));
  const meta=document.createElement('small');meta.textContent=tokens(item.text||'').length+' كلمة';
  info.append(name,meta);
  const actions=document.createElement('div');actions.className='typing-library-actions';
  const preview=document.createElement('button');preview.type='button';preview.className='secondary';preview.textContent='معاينة';
  preview.onclick=()=>{ $('textEdit').value=item.text; $('typingTextName').value=item.name||('نص كتابة '+(index+1)); $('typingImportStatus').textContent='تم تحميل النص للمعاينة. اضغط حفظ لإضافة نسخة محدثة إلى المكتبة.'; };
  const remove=document.createElement('button');remove.type='button';remove.className='secondary';remove.textContent='حذف';remove.onclick=()=>{
   if(texts.length<=1){alert('يجب الاحتفاظ بنص واحد على الأقل في المكتبة.');return;}
   if(!confirm('هل تريد حذف النص «'+(item.name||'نص كتابة')+'»؟'))return;
   db.typingTexts=db.typingTexts.filter(x=>x.id!==item.id);
   if(db.text===item.text)db.text=db.typingTexts[0].text;
   saveDB();renderTypingLibrary();$('textEdit').value=db.text;
  };
  actions.append(preview,remove);row.append(info,actions);list.appendChild(row);
 });
}
$('saveTypingText').onclick=()=>{
 const newText=$('textEdit').value.trim()||DEFAULT_TEXT;
 if(tokens(newText).length!==160){alert('يجب أن يحتوي كل نص محفوظ على 160 كلمة بالضبط.');return}
 const name=$('typingTextName').value.trim()||('نص كتابة '+((db.typingTexts||[]).length+1));
 const item={id:'text_'+Date.now()+'_'+Math.random().toString(36).slice(2,7),name,text:newText,createdAt:new Date().toISOString()};
 db.typingTexts=Array.isArray(db.typingTexts)?db.typingTexts:[];
 db.typingTexts.push(item);db.text=newText;saveDB();renderTypingLibrary();
 $('typingTextName').value='';$('typingImportStatus').textContent='تم حفظ «'+name+'» في مكتبة النصوص. إجمالي النصوص: '+db.typingTexts.length;
 alert('تمت إضافة النص إلى مكتبة النصوص بنجاح.');
};
$('saveTheoryQuestions').onclick=()=>{document.querySelectorAll('#editQuestions [data-field]').forEach(el=>{let i=Number(el.dataset.i),f=el.dataset.field;if(f==='question')db.questions[i][0]=el.value;if(f==='option')db.questions[i][1][Number(el.dataset.j)]=el.value;if(f==='answer')db.questions[i][2]=Number(el.value);if(f==='tf')db.questions[i][1]=el.value==='true'});saveDB();alert('تم حفظ أسئلة الجزء النظري.')};
$('logoFile').onchange=e=>{let file=e.target.files[0];if(!file)return;let reader=new FileReader();reader.onload=()=>{db.settings.logo=reader.result;saveDB();applySettings()};reader.readAsDataURL(file)};
$('saveSettings').onclick=()=>{db.settings.title=$('titleEdit').value.trim()||initial.settings.title;db.settings.primary=$('colorEdit').value;db.settings.bg=$('bgEdit').value;if($('pinEdit').value.trim())db.pin=$('pinEdit').value.trim();saveDB();applySettings();$('pinEdit').value='';alert('تم حفظ الإعدادات.')};

const officeTaskKinds={
 word:[['containsText','وجود نص محدد'],['bold','وجود تنسيق غامق'],['table','إدراج جدول'],['centered','توسيط نص']],
 excel:[['cellEquals','قيمة خلية محددة'],['formulaEquals','معادلة خلية محددة']]
};
function renderOfficeTaskEditor(){
 $('wordContentEdit').value=db.officeContent?.word||'';
 renderExcelFilesAdmin();
 $('wordTotalMarks').value=db.officeConfig.word.totalMarks;
 $('excelTotalMarks').value=(db.officeTasks||[]).filter(t=>t.app==='excel'&&t.formId===activeExcelModelId).reduce((n,t)=>n+(Number(t.points)||0),0);
 $('wordDuration').value=db.officeConfig.word.duration;
 $('excelDuration').value=(db.excelModels.find(m=>m.id===activeExcelModelId)?.duration)||db.officeConfig.excel.duration;
 renderOfficeTasksFor('word');renderOfficeTasksFor('excel');
}

function renderExcelFilesAdmin(){
 const list=$('excelFilesAdminList');if(!list)return;list.innerHTML='';
 (db.excelFiles||[]).forEach(file=>{
  const row=document.createElement('div');row.className='excel-file-row';
  const name=document.createElement('span');name.textContent=file.name;
  const model=(db.excelModels||[]).find(m=>m.fileId===file.id);
  const action=document.createElement('button');action.type='button';action.className='primary';
  if(model){
   const count=(db.officeTasks||[]).filter(t=>t.app==='excel'&&t.formId===model.id).length;
   action.textContent=count?'إضافة / تعديل أسئلة هذا النموذج':'＋ إضافة أسئلة لهذا النموذج';
   action.onclick=()=>{commitOfficeEditor('excel');activeExcelModelId=model.id;renderOfficeTaskEditor();$('addExcelTask')?.click();$('excelTaskEditor')?.scrollIntoView({behavior:'smooth',block:'start'})};
  }else{
   action.textContent='إنشاء نموذج وإضافة أسئلته';
   action.disabled=db.excelModels.length>=4;
   action.onclick=()=>{createExcelModel(file);$('addExcelTask')?.click();$('excelTaskEditor')?.scrollIntoView({behavior:'smooth',block:'start'})};
  }
  const preview=document.createElement('button');preview.type='button';preview.className='secondary';preview.textContent='معاينة الملف';
  preview.onclick=async()=>{try{if(!window.desktopApp?.previewExcelFile){alert('ميزة المعاينة تتوفر في تطبيق سطح المكتب.');return}const result=await window.desktopApp.previewExcelFile(file.path);if(!result?.ok)alert(result?.message||'تعذرت معاينة الملف.')}catch(error){alert('تعذرت معاينة الملف: '+error.message)}};
  const del=document.createElement('button');del.type='button';del.className='secondary';del.textContent='حذف الملف';
  del.onclick=()=>{if(model){if(!confirm('هذا الملف مرتبط بنموذج. هل تريد حذف النموذج وأسئلته والملف معًا؟'))return;db.officeTasks=db.officeTasks.filter(t=>!(t.app==='excel'&&t.formId===model.id));db.excelModels=db.excelModels.filter(m=>m.id!==model.id);if(activeExcelModelId===model.id)activeExcelModelId=db.excelModels[0]?.id||'';}db.excelFiles=db.excelFiles.filter(x=>x.id!==file.id);saveDB();renderExcelFilesAdmin();renderOfficeTaskEditor();renderTraineeExcelFiles()};
  row.append(name,preview,action,del);list.append(row);
 });
 renderExcelModelsAdmin();
}
function createExcelModel(file){
 if(db.excelModels.length>=4){alert('الحد الأقصى أربعة نماذج Excel.');return}
 if(db.excelModels.some(m=>m.fileId===file.id)){alert('هذا الملف مرتبط بنموذج موجود بالفعل.');return}
 const model={id:'model_'+Date.now()+'_'+Math.random().toString(16).slice(2,6),name:'نموذج '+((db.excelModels||[]).length+1)+' - '+file.name,fileId:file.id,duration:20,totalMarks:100};
 db.excelModels.push(model);
 const oldTasks=db.officeTasks.filter(t=>t.app==='excel'&&!t.formId);
 oldTasks.forEach(t=>t.formId=model.id);
 activeExcelModelId=model.id;saveDB();renderExcelFilesAdmin();renderOfficeTaskEditor();
 
}
function renderExcelModelsAdmin(){
 const select=$('excelModelSelect');if(!select)return;select.innerHTML='';
 (db.excelModels||[]).forEach(model=>{const o=document.createElement('option');o.value=model.id;o.textContent=model.name;select.append(o)});
 if(!db.excelModels.length){const o=document.createElement('option');o.value='';o.textContent='لا توجد نماذج بعد';select.append(o);activeExcelModelId='';}
 else if(!db.excelModels.some(m=>m.id===activeExcelModelId))activeExcelModelId=db.excelModels[0].id;
 select.value=activeExcelModelId;
 const nameEdit=$('excelModelNameEdit');if(nameEdit)nameEdit.value=db.excelModels.find(m=>m.id===activeExcelModelId)?.name||'';
 const cards=$('excelModelCards');
 if(cards){
  cards.innerHTML='';
  for(let slot=0;slot<4;slot++){
   const model=db.excelModels[slot];
   const card=document.createElement('div');card.className='excel-model-card'+(model&&model.id===activeExcelModelId?' active':'');
   const title=document.createElement('h4');title.textContent=model?model.name:'النموذج '+(slot+1)+' — فارغ';
   const file=(db.excelFiles||[]).find(f=>f.id===model?.fileId);
   const detail=document.createElement('p');detail.textContent=model?(file?file.name:'ملف غير مرتبط')+' • '+(db.officeTasks||[]).filter(t=>t.app==='excel'&&t.formId===model.id).length+' أسئلة':'أضف ملف Excel لإنشاء هذا النموذج';
   card.append(title,detail);
   if(model){
    const choose=document.createElement('button');choose.type='button';choose.className='secondary';choose.textContent='اختيار وتعديل الأسئلة';
    choose.onclick=()=>{commitOfficeEditor('excel');activeExcelModelId=model.id;renderOfficeTaskEditor();$('excelTaskEditor')?.scrollIntoView({behavior:'smooth',block:'start'})};
    const add=document.createElement('button');add.type='button';add.className='primary';add.textContent='＋ إضافة أسئلة لهذا النموذج';
    add.onclick=()=>{commitOfficeEditor('excel');activeExcelModelId=model.id;renderOfficeTaskEditor();$('addExcelTask')?.click();$('excelTaskEditor')?.scrollIntoView({behavior:'smooth',block:'start'})};
    card.append(choose,add);
   }else{
    const add=document.createElement('button');add.type='button';add.className='secondary';add.textContent='إضافة ملف وإنشاء النموذج';
    add.disabled=!(db.excelFiles||[]).some(f=>!(db.excelModels||[]).some(m=>m.fileId===f.id)) || db.excelModels.length>=4;
    add.onclick=()=>{
     if(db.excelModels.length>=4){alert('اكتملت النماذج الأربعة.');return}
     pendingExcelModelSlot=slot;
     const picker=$('excelFilesPicker');
     if(!picker){alert('تعذر فتح اختيار الملفات. أعد تشغيل التطبيق.');return}
     picker.multiple=false;
     picker.click();
    };
    card.append(add);
   }
   cards.append(card);
  }
 }
 const addModelBtn=$('addExcelModel');if(addModelBtn)addModelBtn.disabled=db.excelModels.length>=4 || !(db.excelFiles||[]).some(f=>!(db.excelModels||[]).some(m=>m.fileId===f.id));
}
function renderTraineeExcelFiles(){
 const label=$('randomExcelModelLabel');if(label)label.textContent=(db.excelModels||[]).length?'سيُختار نموذج عشوائي عند بدء الاختبار':'لا توجد نماذج Excel مضافة بعد';
}
function renderTrainerExcelSheet(serialized=''){
 const root=$('excelSheetEditor');if(!root)return;
 const columns=18,rows=28,letters=Array.from({length:columns},(_,i)=>String.fromCharCode(65+i));
 const source=String(serialized||'').split(/\r?\n/).map(line=>line.split('|'));
 const toolbar=document.createElement('div');toolbar.className='excel-trainer-toolbar';
 const selectBtn=document.createElement('button');selectBtn.type='button';selectBtn.className='secondary';selectBtn.textContent='تحديد كل البيانات';
 const clearBtn=document.createElement('button');clearBtn.type='button';clearBtn.className='secondary';clearBtn.textContent='حذف البيانات المحددة';
 const hint=document.createElement('span');hint.textContent='انسخ البيانات من Excel ثم الصقها هنا بـ Ctrl+V. يمكنك أيضاً الضغط Ctrl+A ثم Delete لمسح الكل.';
 selectBtn.onclick=()=>{root.querySelectorAll('input[data-row]').forEach(cell=>cell.classList.add('excel-cell-selected'));root.dataset.allSelected='1'};
 toolbar.append(selectBtn,clearBtn,hint);
 const table=document.createElement('table');let head=document.createElement('tr');
 const corner=document.createElement('th');corner.className='excel-corner';corner.textContent='';head.append(corner);
 letters.forEach(letter=>{const th=document.createElement('th');th.textContent=letter;th.className='excel-column-header';head.append(th)});table.append(head);
 for(let r=0;r<rows;r++){
  const tr=document.createElement('tr');tr.className='excel-grid-row';const rowHead=document.createElement('th');rowHead.textContent=String(r+1);rowHead.className='excel-row-header';tr.append(rowHead);
  for(let c=0;c<columns;c++){
   const td=document.createElement('td'),input=document.createElement('input');
   if(r===0&&c===0)td.classList.add('excel-active-cell');
   input.type='text';input.autocomplete='off';input.spellcheck=false;input.setAttribute('aria-label',letters[c]+(r+1));
   input.value=source[r]?.[c]||'';input.dataset.row=r;input.dataset.col=c;
   input.addEventListener('input',()=>{db.officeContent.excel=readTrainerExcelSheet();saveDB()});
   input.addEventListener('keydown',event=>{
    if((event.ctrlKey||event.metaKey)&&event.key.toLowerCase()==='a'){
     event.preventDefault();
     root.querySelectorAll('input[data-row]').forEach(cell=>cell.classList.add('excel-cell-selected'));
     root.dataset.allSelected='1';
    }else if(event.key==='Delete' && root.dataset.allSelected==='1'){
     event.preventDefault();
     root.querySelectorAll('input[data-row]').forEach(cell=>{cell.value='';cell.classList.remove('excel-cell-selected')});
     root.dataset.allSelected='';db.officeContent.excel='';saveDB();
    }else if(event.key==='Escape'){
     root.querySelectorAll('.excel-cell-selected').forEach(cell=>cell.classList.remove('excel-cell-selected'));
     root.dataset.allSelected='';
    }else{
     root.querySelectorAll('.excel-cell-selected').forEach(cell=>cell.classList.remove('excel-cell-selected'));
     root.dataset.allSelected='';
    }
   });
   input.addEventListener('focus',()=>{root.querySelectorAll('.excel-cell-selected').forEach(cell=>cell.classList.remove('excel-cell-selected'));root.dataset.allSelected=''});
   input.addEventListener('paste',event=>{
    const clipboard=(event.clipboardData||window.clipboardData)?.getData('text');
    if(!clipboard || !/[\t\r\n]/.test(clipboard))return;
    event.preventDefault();
    const matrix=clipboard.replace(/\r\n/g,'\n').replace(/\r/g,'\n').replace(/\n$/,'').split('\n').map(line=>line.split('\t'));
    const startRow=Number(input.dataset.row),startCol=Number(input.dataset.col);
    matrix.forEach((line,ri)=>line.forEach((value,ci)=>{
     const target=root.querySelector(`input[data-row="${startRow+ri}"][data-col="${startCol+ci}"]`);
     if(target)target.value=value;
    }));
    db.officeContent.excel=readTrainerExcelSheet();saveDB();
   });
   td.append(input);tr.append(td);
  }table.append(tr);
 }
 clearBtn.onclick=()=>{
  if(!confirm('هل تريد حذف جميع البيانات المحددة من جدول Excel؟'))return;
  root.querySelectorAll('input[data-row]').forEach(input=>{input.value='';input.classList.remove('excel-cell-selected')});
  root.dataset.allSelected='';db.officeContent.excel='';saveDB();
 };
 root.replaceChildren(toolbar,table);
}
function readTrainerExcelSheet(){
 const root=$('excelSheetEditor');if(!root)return '';
 const rows=[...root.querySelectorAll('tr')].slice(1).map(tr=>[...tr.querySelectorAll('input')].map(input=>input.value));
 while(rows.length&&rows[rows.length-1].every(v=>!v.trim()))rows.pop();
 return rows.map(row=>row.join('|')).join('\n');
}

function updateOfficeSummary(which){
 const entries=(db.officeTasks||[]).filter(t=>t.app===which && (which!=='excel'||t.formId===activeExcelModelId));
 const count=$(which+'TaskCount'); const total=$(which+'TotalMarks');
 if(count)count.value=entries.length;
 if(total)total.value=entries.reduce((sum,t)=>sum+Math.max(0,Number(t.points)||0),0);
}
function renderOfficeTasksFor(which){
 const root=$(which+'TaskEditor');root.innerHTML='';
 const entries=(db.officeTasks||[]).map((t,i)=>({t,i})).filter(x=>x.t.app===which && (which!=='excel'||x.t.formId===activeExcelModelId));
 entries.forEach(({t,i},localIndex)=>{
  const card=document.createElement('div');card.className='edit-q office-task-edit';
  const heading=document.createElement('h4');heading.className='office-question-heading';heading.textContent='السؤال '+(localIndex+1);
  const title=document.createElement('input');title.value=t.title||'';title.dataset.i=i;title.dataset.f='title';title.placeholder='عنوان مختصر للسؤال';
  const prompt=document.createElement('textarea');prompt.rows=3;prompt.value=t.prompt||'';prompt.dataset.i=i;prompt.dataset.f='prompt';prompt.placeholder='اكتب السؤال بوضوح، مثال: اجعل عنوان المستند غامقاً ووسط الصفحة.';
  const points=document.createElement('input');points.type='number';points.min='0';points.max='1000';points.value=t.points??1;points.dataset.i=i;points.dataset.f='points';points.className='office-points-input';
  points.addEventListener('input',()=>updateOfficeSummary(which));
  const row=(label,node)=>{let wrap=document.createElement('label');wrap.textContent=label;wrap.append(node);return wrap};
  const advanced=document.createElement('details');advanced.className='office-advanced';
  const summary=document.createElement('summary');summary.textContent='إعدادات التصحيح المتقدمة (اختياري)';
  const kind=document.createElement('select');kind.dataset.i=i;kind.dataset.f='kind';
  (officeTaskKinds[which]||[]).forEach(([v,l])=>{let o=document.createElement('option');o.value=v;o.textContent=l;kind.append(o)});kind.value=t.kind;
  const target=document.createElement('input');target.value=t.target||'';target.placeholder=which==='excel'?'مثال: B2':'النص المستهدف (اختياري)';target.dataset.i=i;target.dataset.f='target';
  const expected=document.createElement('input');expected.value=t.expected||'';expected.placeholder='الإجابة أو معيار التنفيذ (اختياري)';expected.dataset.i=i;expected.dataset.f='expected';
  advanced.append(summary,row('طريقة التحقق',kind),row(which==='excel'?'الخلية المستهدفة':'النص المستهدف',target),row('الإجابة أو معيار التنفيذ',expected));
  const del=document.createElement('button');del.className='secondary';del.textContent='حذف السؤال';del.onclick=()=>{commitOfficeEditor(which);const current=db.officeTasks.findIndex(x=>x.id===t.id);if(current>=0)db.officeTasks.splice(current,1);saveDB();renderOfficeTaskEditor()};
  card.append(heading,row('عنوان السؤال (اختياري)',title),row('نص السؤال / المطلوب من المتدرب',prompt),row('درجة هذا السؤال',points),advanced,del);root.append(card);
 });
 // حفظ التعديلات فور تغيير أي حقل، حتى لا تضيع عند الانتقال بين الأقسام أو إغلاق التطبيق.
 root.querySelectorAll('[data-f]').forEach(el=>{
  const persist=()=>{
   const task=db.officeTasks[Number(el.dataset.i)];
   if(!task)return;
   const field=el.dataset.f;
   task[field]=field==='points'?Math.max(0,Number(el.value)||0):el.value;
   saveDB();
   if(field==='points')updateOfficeSummary(which);
  };
  el.addEventListener('input',persist);
  el.addEventListener('change',persist);
 });
 updateOfficeSummary(which);
}
function commitOfficeEditor(which){
 const root=$(which+'TaskEditor');
 if(root){
  root.querySelectorAll('[data-f]').forEach(el=>{
   const task=db.officeTasks[Number(el.dataset.i)];
   if(!task)return;
   const field=el.dataset.f;
   task[field]=field==='points'?Math.max(1,Number(el.value)||1):el.value;
  });
 }
 if(which==='word' && $('wordContentEdit')) db.officeContent.word=$('wordContentEdit').value;
 const scopedTasks=(db.officeTasks||[]).filter(t=>t.app===which && (which!=='excel'||t.formId===activeExcelModelId));
 db.officeConfig[which].totalMarks=scopedTasks.reduce((sum,t)=>sum+Math.max(0,Number(t.points)||0),0);
 db.officeConfig[which].duration=Math.max(1,Number($(which+'Duration')?.value)||db.officeConfig[which].duration||20);
 if(which==='excel'){
  const model=db.excelModels.find(m=>m.id===activeExcelModelId);
  if(model){model.totalMarks=db.officeConfig.excel.totalMarks;model.duration=db.officeConfig.excel.duration}
 }
 saveDB();
}
function saveOfficeTasksFor(which){
 commitOfficeEditor(which);
 renderOfficeTaskEditor();
 alert('تم حفظ محتوى وأسئلة وإعدادات '+(which==='word'?'Word':'Excel')+'.');
}

$('addWordTask').onclick=()=>{commitOfficeEditor('word');db.officeTasks.push({id:'word'+Date.now(),app:'word',title:'مهمة Word جديدة',prompt:'اكتب تعليمات التنسيق المطلوبة هنا.',kind:'bold',target:'',expected:'',points:10});saveDB();renderOfficeTaskEditor()};
$('addExcelTask').onclick=()=>{if(!activeExcelModelId){alert('أضف ملف Excel وأنشئ نموذجًا أولاً.');return}commitOfficeEditor('excel');db.officeTasks.push({id:'excel'+Date.now(),app:'excel',formId:activeExcelModelId,title:'مهمة Excel جديدة',prompt:'اكتب تعليمات إدخال البيانات أو التنسيق أو المعادلة المطلوبة هنا.',kind:'cellEquals',target:'B2',expected:'',points:10});saveDB();renderOfficeTaskEditor()};
$('excelModelSelect').onchange=()=>{commitOfficeEditor('excel');activeExcelModelId=$('excelModelSelect').value;renderOfficeTaskEditor()};
$('saveExcelModelName').onclick=()=>{const model=db.excelModels.find(m=>m.id===activeExcelModelId);if(!model)return alert('اختر نموذجًا أولاً.');model.name=$('excelModelNameEdit').value.trim()||model.name;saveDB();renderExcelModelsAdmin();renderExcelFilesAdmin();alert('تم حفظ اسم النموذج.')};
$('addExcelModel').onclick=()=>{if(db.excelModels.length>=4){alert('يمكنك إضافة أربعة نماذج فقط.');return}const available=(db.excelFiles||[]).filter(f=>!(db.excelModels||[]).some(m=>m.fileId===f.id));if(!available.length){alert('أضف ملف Excel جديدًا أولاً، أو أنشئ نموذجًا من أحد الملفات المستوردة.');return}createExcelModel(available[0])};
$('deleteExcelModel').onclick=()=>{const model=db.excelModels.find(m=>m.id===activeExcelModelId);if(!model)return alert('لا يوجد نموذج محدد.');if(!confirm('هل تريد حذف النموذج وأسئلته؟'))return;db.officeTasks=db.officeTasks.filter(t=>!(t.app==='excel'&&t.formId===model.id));db.excelModels=db.excelModels.filter(m=>m.id!==model.id);activeExcelModelId=db.excelModels[0]?.id||'';saveDB();renderOfficeTaskEditor();renderExcelFilesAdmin()};
$('saveWordTasks').onclick=()=>saveOfficeTasksFor('word');
$('saveExcelTasks').onclick=()=>saveOfficeTasksFor('excel');
$('addExcelFiles').onclick=()=>{if(!window.desktopApp?.isDesktop||!window.desktopApp.importExcelFiles){alert('إضافة الملفات متاحة من تطبيق سطح المكتب المثبت فقط.');return}pendingExcelModelSlot=null;$('excelFilesPicker').multiple=true;$('excelFilesPicker').click()};
$('excelFilesPicker').onchange=async()=>{
 try{
  const result=await window.desktopApp.importExcelFiles();
  if(result?.ok&&Array.isArray(result.files)&&result.files.length){
   db.excelFiles=[...(db.excelFiles||[]),...result.files];
   const remaining=Math.max(0,4-db.excelModels.length);
   const toModel=result.files.slice(0,remaining);
   toModel.forEach(file=>createExcelModel(file));
   saveDB();renderExcelFilesAdmin();renderTraineeExcelFiles();
   if(toModel.length){
    activeExcelModelId=db.excelModels.find(m=>m.fileId===toModel[0].id)?.id||activeExcelModelId;
    renderOfficeTaskEditor();
    $('addExcelTask')?.click();
    $('excelTaskEditor')?.scrollIntoView({behavior:'smooth',block:'start'});
   }
   alert('تمت إضافة '+result.files.length+' ملف Excel وربط '+toModel.length+' منها بنماذج. أضف أسئلة كل نموذج من الزر الموجود بجانب الملف.');
  }
 }catch(error){alert('تعذرت إضافة الملفات: '+error.message)}
 $('excelFilesPicker').value='';
 pendingExcelModelSlot=null;
 $('excelFilesPicker').multiple=true;
};
$('returnToTrainee').onclick=()=>{
 // عند مغادرة لوحة المدرب نعيد المتدرب إلى شاشة إدخال الاسم
 // حتى لا تبقى جلسة الاسم السابق عالقة على لوحة المتدرب.
 stopTimer();
 if(typeof stopOfficeTimer==='function') stopOfficeTimer();
 studentName='';
 typingPassed=false;
 $('name').value='';
 show('student');
 $('name').focus();
};

function renderResults(){
 let root=$('resultList');root.innerHTML='';
 if(!db.results.length){root.textContent='لا توجد نتائج محفوظة حتى الآن.';return}
 db.results.slice().reverse().forEach(r=>{
  let card=document.createElement('div');card.className='result-row';
  let head=document.createElement('div');head.textContent=`${r.name} — ${r.category||'الاختبار الأساسي'} — ${r.earned??r.total??0} / ${r.possible??(r.category==='Word'||r.category==='Excel'?100:50)} — ${r.date}`;card.appendChild(head);
  if(r.modelName){let model=document.createElement('div');model.textContent='النموذج: '+r.modelName;card.appendChild(model)}
  if(Array.isArray(r.details)&&r.details.length){let list=document.createElement('ul');list.style.margin='8px 0 0';r.details.forEach(x=>{let li=document.createElement('li');li.textContent=`${x.title}: ${x.correct?'مُنجز':'غير محلول / غير صحيح'} — ${x.earned} من ${x.possible} درجة`;list.appendChild(li)});card.appendChild(list)}
  root.appendChild(card);
 });
}
$('csv').onclick=()=>{
 const esc=v=>String(v??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
 const byName=new Map();
 db.results.forEach(r=>{
  const name=String(r.name||'').trim()||'بدون اسم';
  if(!byName.has(name))byName.set(name,{name,word:'',excel:'',theory:'',total:0,dates:[]});
  const s=byName.get(name); if(r.date)s.dates.push(r.date);
  if(r.category==='Word')s.word=Number(r.earned??r.total??0);
  else if(r.category==='Excel')s.excel=Number(r.earned??r.total??0);
  else if(r.category==='اختبار المعلومات'){s.theory=Number(r.typing||0)+Number(r.quiz||0);}
 });
 const summaries=[...byName.values()].map(s=>{
  const w=Number(s.word)||0,x=Number(s.excel)||0,t=Number(s.theory)||0;
  s.total=w+x+t;return s;
 });
 const cell=(v,style='Cell')=>'<Cell ss:StyleID="'+style+'"><Data ss:Type="'+(typeof v==='number'?'Number':'String')+'">'+esc(v)+'</Data></Cell>';
 const row=(vals,style='Cell')=>'<Row>'+vals.map((v,i)=>cell(v,i===0&&style==='Cell'?'Name':style)).join('')+'</Row>';
 const summaryRows=[
  row(['اسم المتدرب','درجة Word','درجة Excel','درجة النظري','المجموع'],'Header'),
  ...summaries.map(s=>row([s.name,s.word,s.excel,s.theory,s.total]))
 ];
 const detailRows=[
  row(['اسم المتدرب','نوع الاختبار','النموذج','الدرجة المحصلة','الدرجة الكاملة','التاريخ','تفاصيل المهام'],'Header'),
  ...db.results.map(r=>row([r.name||'',r.category||'الاختبار الأساسي',r.modelName||'',Number(r.earned??r.total??0),Number(r.possible??(r.category==='Word'||r.category==='Excel'?100:50)),r.date||'',Array.isArray(r.details)?r.details.map(d=>(d.correct?'✓ ':'✗ ')+d.title+' ('+d.earned+'/'+d.possible+')').join(' | '):'']))
 ];
 const xml='<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>'+
 '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'+
 '<Styles><Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11"/><Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2F3"/></Borders></Style>'+
 '<Style ss:ID="Header"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#176B50" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#0F5132"/></Borders></Style>'+
 '<Style ss:ID="Name"><Alignment ss:Vertical="Center" ss:Horizontal="Right"/><Font ss:FontName="Arial" ss:Size="11"/><Interior ss:Color="#F1F7F4" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2F3"/></Borders></Style>'+
 '<Style ss:ID="Cell"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2F3"/></Borders></Style></Styles>'+
 '<Worksheet ss:Name="ملخص النتائج"><Table><Column ss:Width="190"/><Column ss:Width="110"/><Column ss:Width="110"/><Column ss:Width="110"/><Column ss:Width="110"/>'+summaryRows.join('')+'</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><DisplayRightToLeft/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions></Worksheet>'+
 '<Worksheet ss:Name="تفاصيل الاختبارات"><Table><Column ss:Width="170"/><Column ss:Width="130"/><Column ss:Width="160"/><Column ss:Width="100"/><Column ss:Width="100"/><Column ss:Width="170"/><Column ss:Width="420"/>'+detailRows.join('')+'</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><DisplayRightToLeft/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>1</SplitHorizontal><TopRowBottomPane>1</TopRowBottomPane><ActivePane>2</ActivePane></WorksheetOptions></Worksheet></Workbook>';
 const blob=new Blob(['\uFEFF',xml],{type:'application/vnd.ms-excel;charset=utf-8'});
 download(blob,'نتائج_المتدربين.xls');
}; 
$('backup').onclick=()=>download(new Blob([JSON.stringify(db,null,2)],{type:'application/json'}),'نسخة_احتياطية_للمنصة.json');
$('restore').onchange=e=>{let f=e.target.files[0];if(!f)return;let r=new FileReader();r.onload=()=>{try{let incoming=JSON.parse(r.result);if(!incoming.text||!incoming.questions||!incoming.settings)throw Error();db={...clone(initial),...incoming};saveDB();applySettings();loadAdmin();renderResults();alert('تم استيراد النسخة الاحتياطية بنجاح.')}catch{alert('ملف النسخة الاحتياطية غير صالح.')}};r.readAsText(f)};
function download(blob,name){let a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}

// Word / Excel interactive training and assessment simulator
let officeApp='word', officeStarted=false, officeTrainee='';
const wordTaskHTML='<b>مهمة Word:</b><p>أنشئ مستنداً بعنوان <b>مهارات الحاسوب</b>. اجعل العنوان غامقاً وفي الوسط، واكتب فقرة تتضمن عبارة «التقنية الحديثة»، ثم أدرج جدولاً من عمودين وصفّين على الأقل.</p>';
const excelTaskHTML='<b>مهمة Excel:</b><p>أدخل في الخلايا B2 إلى B5 الأرقام: 120، 180، 200، 100. اكتب في E2 المعادلة <b>=SUM(B2:B5)</b>، وفي E3 المعادلة <b>=AVERAGE(B2:B5)</b>. يجب أن تظهر النتيجتان 600 و150.</p>';
function setOfficeApp(which){if(typeof stopOfficeTimer==='function')stopOfficeTimer();officeStarted=false;$('officeExamBar').hidden=true;$('officeTask').hidden=false;officeApp=which;$('wordWorkspace').hidden=which!=='word';$('excelWorkspace').hidden=which!=='excel';$('wordTab').classList.toggle('active',which==='word');$('excelTab').classList.toggle('active',which==='excel');$('officeHeading').textContent=which==='word'?'تدريب واختبار Word':'تدريب واختبار Excel';$('officeTask').innerHTML='<b>اختر وضع التدريب أو الاختبار ثم اضغط بدء النشاط.</b><p>يمكنك استخدام أدوات المحاكاة مباشرة للتجربة.</p>';$('officeFeedback').hidden=true;officeStarted=false}
$('wordTab').onclick=()=>setOfficeApp('word');$('excelTab').onclick=()=>setOfficeApp('excel');
let officeRunTasks=[],officeQuestionIndex=0,officeSeconds=0,officeTimer=null;
function renderOfficeQuestion(){
 const task=officeRunTasks[officeQuestionIndex];if(!task)return;
 $('officeExamProgress').textContent='السؤال '+(officeQuestionIndex+1)+' من '+officeRunTasks.length+' · '+Number(task.points||0)+' درجة';
 $('officeExamQuestion').textContent=task.prompt;
 $('officePrev').disabled=officeQuestionIndex===0;$('officeNext').disabled=officeQuestionIndex===officeRunTasks.length-1;
 $('officeExamTimer').textContent=formatTime(officeSeconds);
}
function stopOfficeTimer(){if(officeTimer){clearInterval(officeTimer);officeTimer=null}}
function startOfficeTimer(){
 stopOfficeTimer();officeSeconds=(Number(db.officeConfig[officeApp].duration)||20)*60;
 $('officeExamTimer').textContent=formatTime(officeSeconds);
 if($('officeMode').value!=='test')return;
 officeTimer=setInterval(()=>{officeSeconds=Math.max(0,officeSeconds-1);$('officeExamTimer').textContent=formatTime(officeSeconds);if(officeSeconds===0){stopOfficeTimer();finishOfficeExam(true)}},1000);
}
function renderStudentOfficeTasks(which){let tasks=(db.officeTasks||[]).filter(t=>t.app===which);return tasks.length?tasks.map(t=>t.prompt).join(' • '):''}
$('officeStart').onclick=()=>{
 officeTrainee=$('officeName').value.trim();
 if(!db.unlockedTrainees[traineeKey(officeTrainee)]){alert('يجب اجتياز 160 كلمة صحيحة في اختبار الكتابة أولاً.');return}
 if(!officeTrainee){alert('يرجى كتابة اسم المتدرب أولاً.');$('officeName').focus();return}
 officeRunTasks=(db.officeTasks||[]).filter(t=>t.app===officeApp&&(officeApp!=='excel'||t.formId===selectedExcelExamModel?.id)).map(t=>({...t}));
 if(!officeRunTasks.length){alert('لا توجد مهام لهذا القسم. اطلب من المدرب إضافة المهام من لوحة التحكم.');return}
 officeStarted=true;officeQuestionIndex=0;$('officeFeedback').hidden=true;$('nativeOfficeStatus').hidden=true;
 $('officeExamBar').hidden=false;$('officeTask').hidden=true;
 if(officeApp==='word'){if($('officeMode').value==='practice')$('wordEditor').innerHTML='<h2>مهارات الحاسوب</h2><p>اكتب وعدّل النص ثم جرّب أدوات التنسيق.</p>'}
 else resetExcelGrid();
 renderOfficeQuestion();startOfficeTimer();
};
$('officeReset').onclick=()=>{stopOfficeTimer();officeRunTasks=[];officeStarted=false;$('officeExamBar').hidden=true;$('officeTask').hidden=false;$('officeFeedback').hidden=true;if(officeApp==='word')$('wordEditor').innerHTML='<h2>مهارات الحاسوب</h2><p>اكتب هنا وتدرّب على تنسيق النصوص والفقرات في المستندات.</p>';else resetExcelGrid()};
$('officePrev').onclick=()=>{if(officeQuestionIndex>0){officeQuestionIndex--;renderOfficeQuestion()}};
$('officeNext').onclick=()=>{if(officeQuestionIndex<officeRunTasks.length-1){officeQuestionIndex++;renderOfficeQuestion()}};
$('officeFinish').onclick=()=>finishOfficeExam(false);
document.querySelectorAll('[data-cmd]').forEach(b=>b.onclick=()=>{ $('wordEditor').focus();document.execCommand(b.dataset.cmd,false,null)});
document.querySelectorAll('[data-align]').forEach(b=>b.onclick=()=>{$('wordEditor').focus();document.execCommand('justify'+b.dataset.align,false,null)});
$('wordFontSize').onchange=e=>{$('wordEditor').focus();document.execCommand('fontSize',false,e.target.value)};
$('wordInsertTable').onclick=()=>{$('wordEditor').focus();document.execCommand('insertHTML',false,'<table><tr><th>العنصر</th><th>الوصف</th></tr><tr><td>الحاسوب</td><td>جهاز لمعالجة البيانات</td></tr></table><p></p>')};
$('wordCheck').onclick=()=>{if(!officeStarted){alert('ابدأ النشاط أولاً.');return}finishOfficeExam(false)};
function resetExcelGrid(){let t=$('excelGrid');let letters=['A','B','C','D','E'];let html='<tr><th></th>'+letters.map(x=>'<th>'+x+'</th>').join('')+'</tr>';for(let r=1;r<=8;r++){html+='<tr><th>'+r+'</th>';for(let c=0;c<5;c++){let val='';if(r===1)val=['البيان','القيمة','','','النتيجة'][c];if(c===0&&r>1)val='بند '+(r-1);if(c===1&&r>=2&&r<=5)val=[120,180,200,100][r-2];html+='<td><input id="x-'+r+'-'+c+'" value="'+val+'" data-row="'+r+'" data-col="'+c+'" aria-label="'+letters[c]+r+'"><span class="cell-value" id="v-'+r+'-'+c+'"></span></td>'}html+='</tr>'}t.innerHTML=html;document.querySelectorAll('#excelGrid input').forEach(inp=>{inp.onfocus=()=>selectExcelCell(inp);inp.oninput=()=>{if($('excelFormulaBar').dataset.sync!=='1')$('excelFormulaBar').value=inp.value;calculateExcel()};inp.onblur=calculateExcel});$('excelFormulaBar').value='';$('excelCellName').textContent='A1';calculateExcel()}
function selectExcelCell(inp){document.querySelectorAll('#excelGrid td').forEach(td=>td.classList.remove('selected'));inp.parentElement.classList.add('selected');let r=Number(inp.dataset.row),c=Number(inp.dataset.col);$('excelCellName').textContent='ABCDE'[c]+r;$('excelFormulaBar').dataset.row=r;$('excelFormulaBar').dataset.col=c;$('excelFormulaBar').value=inp.value}
$('excelFormulaBar').onchange=e=>{let r=Number(e.target.dataset.row),c=Number(e.target.dataset.col);if(!r)return;let cell=$('x-'+r+'-'+c);cell.value=e.target.value;calculateExcel()};
function getExcelVal(r,c){let el=$('x-'+r+'-'+c);if(!el)return 0;let v=el.value.trim();if(v.startsWith('='))return evalExcelFormula(v);let n=Number(v);return Number.isFinite(n)?n:0}
function evalExcelFormula(f){let m=f.trim().match(/^=(SUM|AVERAGE|COUNT)\(([A-E])(\d+):([A-E])(\d+)\)$/i);if(!m)return NaN;let fn=m[1].toUpperCase(),c1='ABCDE'.indexOf(m[2].toUpperCase()),r1=Number(m[3]),c2='ABCDE'.indexOf(m[4].toUpperCase()),r2=Number(m[5]);if(c1<0||c2<0)return NaN;let vals=[];for(let r=Math.min(r1,r2);r<=Math.max(r1,r2);r++)for(let c=Math.min(c1,c2);c<=Math.max(c1,c2);c++){let raw=$('x-'+r+'-'+c)?.value.trim()||'';if(raw!==''&&!raw.startsWith('=')&&Number.isFinite(Number(raw)))vals.push(Number(raw))}if(fn==='SUM')return vals.reduce((a,b)=>a+b,0);if(fn==='COUNT')return vals.length;return vals.length?vals.reduce((a,b)=>a+b,0)/vals.length:0}
function calculateExcel(){if(!$('excelGrid'))return;document.querySelectorAll('#excelGrid input').forEach(inp=>{let v=inp.value.trim(),out=$('v-'+inp.dataset.row+'-'+inp.dataset.col);if(v.startsWith('=')){let n=evalExcelFormula(v);out.textContent=Number.isFinite(n)?'='+n:'معادلة غير مدعومة'}else out.textContent=''});}
$('excelCheck').onclick=()=>{if(!officeStarted){alert('ابدأ النشاط أولاً.');return}finishOfficeExam(false)};
function evaluateOfficeTask(task){
 if(officeApp==='word'){
  let el=$('wordEditor');
  if(task.kind==='containsText')return (el.innerText||'').includes(task.expected||'');
  if(task.kind==='bold')return !!el.querySelector('b,strong');
  if(task.kind==='table')return !!el.querySelector('table');
  if(task.kind==='centered')return [...el.querySelectorAll('*')].some(n=>(n.getAttribute('style')||'').includes('center')||n.align==='center');
  return false;
 }
 let cell=(task.target||'').toUpperCase().match(/^([A-E])([1-8])$/);if(!cell)return false;
 let c='ABCDE'.indexOf(cell[1]),r=Number(cell[2]),value=$('x-'+r+'-'+c)?.value.trim()||'';
 if(task.kind==='cellEquals')return value.toLowerCase()===(task.expected||'').trim().toLowerCase();
 if(task.kind==='formulaEquals')return value.replace(/\s/g,'').toUpperCase()===(task.expected||'').replace(/\s/g,'').toUpperCase();
 return false;
}
function finishOfficeExam(autoFinished){
 if(!officeStarted||!officeRunTasks.length)return;
 stopOfficeTimer();
 let checks=officeRunTasks.map(evaluateOfficeTask);
 let rawPossible=officeRunTasks.reduce((a,t)=>a+Math.max(1,Number(t.points)||1),0);
 let rawEarned=officeRunTasks.reduce((a,t,i)=>a+(checks[i]?Math.max(1,Number(t.points)||1):0),0);
 let percent=rawPossible?Math.round(rawEarned/rawPossible*100):0;
 let totalMarks=Math.max(1,Number(officeApp==='excel'&&selectedExcelExamModel?selectedExcelExamModel.totalMarks:db.officeConfig[officeApp].totalMarks)||100);
 let earned=Math.round(totalMarks*rawEarned/rawPossible);
 // Distribute the final scaled score per task; every missed task is explicitly recorded as zero.
 let details=officeRunTasks.map((t,i)=>({title:t.title||t.prompt,correct:!!checks[i],possible:Math.round(totalMarks*Math.max(1,Number(t.points)||1)/rawPossible),earned:checks[i]?Math.round(totalMarks*Math.max(1,Number(t.points)||1)/rawPossible):0}));
 // Correct rounding so the task breakdown always matches the displayed total.
 let delta=earned-details.reduce((n,x)=>n+x.earned,0);if(delta>0){let last=details.map((x,i)=>x.correct?i:-1).filter(i=>i>=0).pop();if(last!==undefined)details[last].earned+=delta;}
 $('officeFeedback').hidden=false;
 $('officeFeedback').innerHTML='<div><b>انتهى اختبار '+(officeApp==='word'?'Word':'Excel')+'</b></div><div>الدرجة: '+earned+' من '+totalMarks+' ('+percent+'%)</div><ul>'+details.map(x=>'<li>'+(x.correct?'✓ ':'✗ ')+escapeHtml(x.title)+' — '+x.earned+' من '+x.possible+' درجة</li>').join('')+'</ul>'+(autoFinished?'<p>انتهى الوقت وتم تسليم الاختبار تلقائياً.</p>':'');
 if($('officeMode').value==='test'){
  db.results.push({name:officeTrainee,date:new Date().toLocaleString('ar-SA'),typing:0,quiz:0,total:earned,category:officeApp==='word'?'Word':'Excel',modelName:officeApp==='excel'?(selectedExcelExamModel?.name||'نموذج Excel'):'',officeScore:percent,earned,possible:totalMarks,details});
  saveDB();$('officeFeedback').innerHTML+='<p>تم حفظ النتيجة في لوحة المدرب.</p>';
 }else $('officeFeedback').innerHTML+='<p>هذا تقييم تدريبي ويمكنك إعادة المحاولة.</p>';
 officeStarted=false;$('officeExamBar').hidden=true;$('officeTask').hidden=true;
 $('officeFeedback').hidden=false;
 $('officeFeedback').setAttribute('tabindex','-1');
 $('officeFeedback').focus({preventScroll:true});
 $('officeFeedback').scrollIntoView({behavior:'smooth',block:'center'});
}
resetExcelGrid();setOfficeApp('word');


async function openNativeOffice(kind){
 if(!trainerPreviewMode&&!db.unlockedTrainees[traineeKey(officeTrainee||$('officeName').value.trim())]){alert('يجب اجتياز اختبار الكتابة أولاً لفتح برامج Office.');return}
 const status=$('nativeOfficeStatus');status.hidden=false;status.textContent='جارٍ التحقق من وجود البرنامج وفتح ملف التدريب...';
 if(!window.desktopApp||!window.desktopApp.isDesktop||typeof window.desktopApp.openOffice!=='function'){
   status.textContent='أنت تشغّل نسخة المتصفح حالياً. أغلق هذه النافذة وافتح تطبيق «منصة اختبار إدخال البيانات» من اختصار سطح المكتب بعد تثبيته. لا تفتح index.html أو نسخة Chrome/Edge؛ فالمتصفح لا يستطيع تشغيل Word وExcel مباشرة.';alert(status.textContent);
   return;
 }
 try{
   const selectedModel=kind==='excel'?selectedExcelExamModel:null;
   const selectedExcel=selectedModel?(db.excelFiles||[]).find(f=>f.id===selectedModel.fileId):null;
   const tasks=kind==='excel'&&selectedModel?(db.officeTasks||[]).filter(t=>t.app==='excel'&&t.formId===selectedModel.id):(officeRunTasks.length?officeRunTasks:(db.officeTasks||[]).filter(t=>t.app===kind));
   const duration=selectedModel?.duration||db.officeConfig?.[kind]?.duration||20;
   const result=await window.desktopApp.openOffice(kind,tasks,officeTrainee||studentName,db.officeContent?.[kind]||'',duration,selectedExcel?.path||'');
   status.textContent=result.message|| (result.ok?'تم فتح البرنامج.':'تعذر فتح البرنامج.');
   status.style.background=result.ok?'#eaf6ee':'#fff3e8';alert(result.message|| (result.ok?'تم فتح البرنامج الأصلي.':'تعذر فتح البرنامج الأصلي.'));
 }catch(err){status.textContent='تعذر فتح البرنامج: '+err.message;status.style.background='#fff3e8'}
}

if (window.desktopApp && typeof window.desktopApp.onOfficeExamReturned === 'function') {
  window.desktopApp.onOfficeExamReturned((data) => {
    updateTraineeDashboard();
    show('traineeDashboard');
    const label = data?.kind === 'excel' ? 'Excel' : 'Word';
    const welcome = $('welcomeTrainee');
    if (welcome) welcome.textContent = 'تم إنهاء اختبار ' + label + '. يمكنك اختيار اختبار آخر من القائمة.';
  });
}

$('openNativeWord').onclick=()=>openNativeOffice('word');
$('openNativeExcel').onclick=()=>openNativeOffice('excel');


// تم السماح بالنسخ واللصق والسحب والإفلات في حقل الكتابة للمتدرب.


function ensureOfficeAssets(){
 if(!db.officeAssets||typeof db.officeAssets!=='object')db.officeAssets={};
 ['word','excel'].forEach(k=>{if(!db.officeAssets[k])db.officeAssets[k]={template:[],solution:[]};['template','solution'].forEach(r=>{if(!Array.isArray(db.officeAssets[k][r]))db.officeAssets[k][r]=[]})});
}
function renderOfficeAssets(){
 ensureOfficeAssets();
 [['word','template','wordTemplateAssets'],['word','solution','wordSolutionAssets'],['excel','template','excelTemplateAssets'],['excel','solution','excelSolutionAssets']].forEach(([kind,role,id])=>{
  const root=$(id);if(!root)return;root.innerHTML='';
  (db.officeAssets[kind][role]||[]).forEach(asset=>{
   const row=document.createElement('div');row.className='office-asset-row';
   const name=document.createElement('span');name.textContent=asset.name||asset.id;row.appendChild(name);
   const del=document.createElement('button');del.type='button';del.className='secondary';del.textContent='حذف';
   del.onclick=async()=>{
    if(!confirm('هل تريد حذف هذا الملف من مساحة المدرب؟'))return;
    const result=await window.desktopApp.deleteOfficeAsset(kind,role,asset.id);
    if(result?.ok){db.officeAssets[kind][role]=db.officeAssets[kind][role].filter(x=>x.id!==asset.id);saveDB();renderOfficeAssets()}
    else alert(result?.message||'تعذر حذف الملف.');
   };
   row.appendChild(del);root.appendChild(row);
  });
  if(!root.children.length){const empty=document.createElement('small');empty.className='hint';empty.textContent='لم يتم رفع ملفات بعد.';root.appendChild(empty)}
 });
}
async function uploadOfficeAsset(kind,role){
 if(!window.desktopApp?.isDesktop||!window.desktopApp.importOfficeAsset){alert('رفع الملفات متاح من تطبيق سطح المكتب المثبت فقط.');return}
 const result=await window.desktopApp.importOfficeAsset(kind,role);
 if(result?.ok&&result.asset){ensureOfficeAssets();db.officeAssets[kind][role].push(result.asset);saveDB();renderOfficeAssets();setAdminSaveState(false)}
 else if(!result?.canceled)alert(result?.message||'تعذر رفع الملف.');
}
[['uploadWordTemplate','word','template'],['uploadWordSolution','word','solution'],['uploadExcelTemplate','excel','template'],['uploadExcelSolution','excel','solution']].forEach(([id,kind,role])=>{
 const button=$(id);if(button)button.onclick=()=>uploadOfficeAsset(kind,role);
});

ensureOfficeAssets();renderOfficeAssets();
applySettings();renderPassage();
