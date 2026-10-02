'use strict';
(() => {
 const F=window.Filament,$=id=>document.getElementById(id);
 const now=new Date();$('date').value=[now.getFullYear(),String(now.getMonth()+1).padStart(2,'0'),String(now.getDate()).padStart(2,'0')].join('-');
 const previous=F.persistentLoad('filamentRegistrationPrefixV1',F.localLoad('filamentRegistrationPrefixV1','S'));if(F.places[previous])$('prefix').value=previous;
 let busy=false;
 function mode(){const custom=$('mode').value==='custom';$('autoFields').hidden=custom;$('customFields').hidden=!custom;$('date').disabled=custom;$('count').disabled=custom;$('customId').required=custom}
 $('mode').onchange=mode;$('prefix').onchange=()=>F.persistentSave('filamentRegistrationPrefixV1',$('prefix').value);mode();
 $('startForm').onsubmit=async event=>{
  event.preventDefault();if(busy||!$('startForm').reportValidity())return;
  busy=true;$('startFields').disabled=true;$('status').classList.remove('error');$('status').textContent='最新の管理番号を確認しています…';
  try {
   let id=F.clean($('customId').value).toUpperCase(),count=1;
   if($('mode').value==='auto'){
    const date=$('date').value,stem=$('prefix').value+date.slice(2).replaceAll('-','');
    if(!/^20\d{2}-\d{2}-\d{2}$/.test(date)||!F.validId(stem+'01'))throw Error('入庫日を確認してください。');
    const rows=await F.candidates();
    const idColumn=rows.length?Object.keys(rows[0])[0]:null;
    if(rows.length&&!/^管理番号(?:[（(]半角[）)])?$/.test(idColumn))throw Error('台帳の管理番号列を確認できません。自動採番を中止しました。');
    const ids=rows.map(row=>F.clean(row[idColumn]));
    // CSVの反映待ちでも、このタブで登録に成功したIDを再候補にしない。
    ids.push(...F.recentRegistrations().map(item=>item.id));
    const max=Math.max(0,...ids.filter(value=>F.validId(value)&&value.startsWith(stem)).map(value=>Number(value.slice(-2))));
    count=$('count').valueAsNumber;
    if(!Number.isInteger(count)||count<1||count>20)throw Error('一括登録は1〜20本で指定してください。');
    if(max+count>99)throw Error('この本数では連番99を超えます。本数または入庫日を確認してください。');
    id=stem+String(max+1).padStart(2,'0');
   }
   if(!F.validId(id))throw Error('管理番号はS/O/F＋日付YYMMDD＋連番2桁で入力してください。');
   F.persistentSave('filamentRegistrationPrefixV1',id[0]);
   location.href='./register.html?id='+encodeURIComponent(id)+(count>1?'&count='+count:'');
  }catch(error){$('status').textContent=error.message;$('status').classList.add('error')}
  finally{busy=false;$('startFields').disabled=false;mode()}
 };
})();
