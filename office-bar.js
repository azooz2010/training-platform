const $=id=>document.getElementById(id);
let current={index:0,count:0,seconds:1200,kind:'word'};
function time(s){return String(Math.floor(s/60)).padStart(2,'0')+':'+String(s%60).padStart(2,'0')}
if(window.desktopApp?.examBar){
 window.desktopApp.examBar.onUpdate(data=>{
  current=data;
  $('prompt').textContent=data.prompt||'نفّذ المطلوب داخل البرنامج الأصلي';
  $('clock').textContent=time(data.seconds);
  $('prev').disabled=data.index<=0;$('next').disabled=data.index>=data.count-1;
 });
 window.desktopApp.examBar.onTimeUp(()=>{$('prompt').textContent='انتهى الوقت. احفظ عملك ثم اضغط إنهاء وإرسال الامتحان.';$('clock').textContent='00:00'});
 $('prev').onclick=()=>window.desktopApp.examBar.previous();
 $('next').onclick=()=>window.desktopApp.examBar.next();
 $('finish').onclick=()=>{if(confirm('هل تريد إنهاء الاختبار وإغلاق شريط التعليمات؟'))window.desktopApp.examBar.finish()};
}else{$('prompt').textContent='افتح الاختبار من تطبيق المنصة المثبت على Windows.'}
