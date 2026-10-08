'use strict';
(() => {
 const F=window.Filament,$=id=>document.getElementById(id),params=new URLSearchParams(location.search);let busy=false;
 const requestedReturn=params.get('return')||'';
 function safeReturn(value){if(!value.startsWith('register.html?'))return './register-start.html';const query=new URLSearchParams(value.slice(value.indexOf('?')+1)),id=F.clean(query.get('id')),count=Number(query.get('count')||1);if([...query.keys()].some(key=>!['id','count','returnIds'].includes(key))||!F.validId(id)||!Number.isInteger(count)||count<1||count>20)return './register-start.html';return './register.html?'+query.toString()}
 const returnUrl=safeReturn(requestedReturn);
 $('back').href=returnUrl;$('return').href=returnUrl;
 function status(message,error=false){$('status').textContent=message;$('status').classList.toggle('error',error)}
 function detail(data){$('detail').replaceChildren();[['フィラメントID',data.id],['メーカー',data.maker],['商品名',data.product||'—'],['素材',data.base],['サブカテゴリ',data.sub],['色',data.color]].forEach(([label,value])=>{const dt=document.createElement('dt'),dd=document.createElement('dd');dt.textContent=label;dd.textContent=value;$('detail').append(dt,dd)})}
 async function load(){try{const makers=await F.makerMasters();if(!makers.length)throw Error('メーカーマスターが空です。');$('makerId').replaceChildren(new Option('選択してください',''));makers.sort((a,b)=>a.name.localeCompare(b.name,'ja')).forEach(item=>$('makerId').add(new Option(item.name+'（'+item.id+'）',item.id)));$('fields').disabled=false;status('メーカーを選び、商品の情報を入力してください。')}catch(error){status(error.message,true)}}
 $('form').onsubmit=async event=>{event.preventDefault();if(busy||!$('form').reportValidity())return;busy=true;$('fields').disabled=true;status('重複を確認してマスターへ追加しています…');const item={makerId:$('makerId').value,product:F.clean($('product').value),base:F.clean($('base').value),sub:F.clean($('sub').value),color:F.clean($('color').value),aliases:F.clean($('aliases').value),purchaseUrl:F.clean($('purchaseUrl').value),note:F.clean($('note').value)};try{const data=await F.createFilamentMaster(item);detail(data);$('form').closest('.panel').hidden=true;$('complete').hidden=false;status('')}catch(error){status(error.message,true);$('fields').disabled=false}finally{busy=false}};
 load();
})();
