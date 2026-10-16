'use strict';
(() => {
 const F=window.Filament,$=id=>document.getElementById(id);let aliases=[],busy=false;
 function status(message,error=false){$('status').textContent=message;$('status').classList.toggle('error',error)}
 function renderAliases(){$('aliasList').replaceChildren(...aliases.map(value=>{const button=document.createElement('button');button.type='button';button.textContent=value+' ×';button.onclick=()=>{aliases=aliases.filter(item=>item!==value);renderAliases()};return button}))}
 $('addAlias').onclick=()=>{const value=F.clean($('aliasEntry').value);if(!value||aliases.includes(value))return;aliases.push(value);$('aliasEntry').value='';renderAliases()};
 $('form').onsubmit=async event=>{event.preventDefault();if(busy||!$('form').reportValidity())return;busy=true;$('fields').disabled=true;status('重複を確認してメーカーを登録しています…');try{const data=await F.createMakerMaster({name:F.clean($('name').value),aliases:aliases.join('｜'),spoolWeight:$('spoolWeight').value,provisional:false});$('form').closest('.panel').hidden=true;$('complete').hidden=false;$('detail').replaceChildren();[['メーカーID',data.id],['メーカー名',data.name],['標準空スプール重量',data.spoolWeight==null?'未設定':data.spoolWeight+'g'],['状態',data.status]].forEach(([label,value])=>{const dt=document.createElement('dt'),dd=document.createElement('dd');dt.textContent=label;dd.textContent=value;$('detail').append(dt,dd)});status('正式なメーカーとして登録しました。')}catch(error){status(error.message,true);$('fields').disabled=false}finally{busy=false}};
 renderAliases();
})();
