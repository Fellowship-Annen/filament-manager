'use strict';
window.Filament = (() => {
  const api = 'https://script.google.com/macros/s/AKfycbzCCejDFlAfJV5sjxhzH7wp1-4S7oVK6s0Gru_m4QXVmSRdnnOXDcCfqeogW5GA4PFR/exec';
  const home = 'https://fellowship-annen.github.io/filament-manager/';
  const csv = 'https://docs.google.com/spreadsheets/d/15UGqdqOwVQNx9qY02HtZIaqrD6Kd1dwmWuj4p3EGIGg/export?format=csv&gid=93776902';
  const places = {S:'さんらいず',O:'オーシャン',F:'フォージー'};
  const clean = value => String(value ?? '').trim();
  function validId(id) {
    if (!/^[SOF]\d{8}$/.test(id)) return false;
    const y=2000+Number(id.slice(1,3)),m=Number(id.slice(3,5)),d=Number(id.slice(5,7));
    const date=new Date(Date.UTC(y,m-1,d));
    return date.getUTCFullYear()===y && date.getUTCMonth()===m-1 && date.getUTCDate()===d;
  }
  async function request(url, options={}) {
    const controller=new AbortController(), timeout=setTimeout(()=>controller.abort(),45000);
    try {
      const response=await fetch(url,{...options,cache:'no-store',signal:controller.signal});
      if (!response.ok) throw Error('通信エラー（HTTP '+response.status+'）');
      let data;try{data=JSON.parse(await response.text())}catch{throw Error('GASの応答を読めません。公開設定を確認してください。')}
      if (!data || typeof data.ok!=='boolean') throw Error('GASの応答形式が一致しません。');
      return data;
    } catch(error) {if(error.name==='AbortError')throw Error('通信がタイムアウトしました。');throw error}
    finally {clearTimeout(timeout)}
  }
  async function lookup(id) {
    const url=new URL(api);url.searchParams.set('id',id);
    const data=await request(url);
    if (data.ok && clean(data.id)===id) return data;
    if (data.ok) throw Error('取得した管理番号が一致しません。');
    if (data.error==='ID_NOT_FOUND') return null;
    throw Error(data.message||data.error||'登録状況を確認できません。');
  }
  async function verify() {
    const url=new URL(api);url.searchParams.set('action','health');const data=await request(url);
    if(!data.ok||data.apiVersion!=='filament-form-history-v3'||data.historySheet!=='フォームの回答 1'||!data.actions?.includes('register'))throw Error('新規登録と履歴記録に対応したGASではありません。');
  }
  async function candidates() {
    if(typeof Papa==='undefined')throw Error('候補の読込機能を取得できません。');
    const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),30000);
    try {
      const response=await fetch(csv,{cache:'no-store',signal:controller.signal});if(!response.ok)throw Error('登録候補の取得に失敗しました。');
      const data=Papa.parse(await response.text(),{header:true,skipEmptyLines:true,transformHeader:v=>v.trim()});
      if(data.errors.length||!data.meta.fields?.includes('メーカー'))throw Error('台帳の候補を読み取れません。');
      return data.data;
    } catch(error) {
      if(location.protocol==='file:'&&error instanceof TypeError)throw Error('ファイルを直接開いたため、Googleの在庫データを取得できません。ローカル確認用の起動ファイル、または公開済みサイトから開いてください。');
      if(error.name==='AbortError')throw Error('在庫データの取得がタイムアウトしました。再度お試しください。');
      if(error instanceof TypeError)throw Error('在庫データに接続できません。通信状況・シートの公開設定を確認してください。');
      throw error;
    } finally {clearTimeout(timer)}
  }
  async function updateLabel(id,label) {
    const healthUrl=new URL(api);healthUrl.searchParams.set('action','health');const health=await request(healthUrl);
    if(!health.ok||!health.actions?.includes('updateLabel'))throw Error('公開中のGASはラベル注釈の保存にまだ対応していません。入力した文字は今回のQRには使えます。');
    const data=await request(api,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'updateLabel',id,label})});
    if(!data.ok)throw Error(data.message||data.error||'ラベル注釈を保存できません。');
    if(data.action!=='updateLabel'||clean(data.id)!==id)throw Error('ラベル注釈の保存結果が一致しません。');
    return data;
  }
  function localLoad(key,fallback){try{return JSON.parse(sessionStorage.getItem(key))??fallback}catch{return fallback}}
  function localSave(key,value){try{sessionStorage.setItem(key,JSON.stringify(value))}catch{}}
  function persistentLoad(key,fallback){try{return JSON.parse(localStorage.getItem(key))??fallback}catch{return fallback}}
  function persistentSave(key,value){try{localStorage.setItem(key,JSON.stringify(value))}catch{}}
  function itemUrl(id){const url=new URL(home);url.searchParams.set('id',id);return url.toString()}
  function recentRegistrations(){const items=localLoad('filamentRecentRegistrationsV1',[]);return Array.isArray(items)?items.filter(item=>item&&validId(item.id)):[]}
  function rememberRegistration(item){const items=recentRegistrations().filter(old=>old.id!==item.id);items.push({id:item.id,maker:item.maker,base:item.base,color:item.color});localSave('filamentRecentRegistrationsV1',items)}
  return {api,home,places,clean,validId,request,lookup,verify,candidates,updateLabel,localLoad,localSave,persistentLoad,persistentSave,itemUrl,recentRegistrations,rememberRegistration};
})();
