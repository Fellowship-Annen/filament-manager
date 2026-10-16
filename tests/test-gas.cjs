const fs=require('fs'),vm=require('vm'),assert=require('assert');
let sheets,failLedger=false,failLog=false;
class Sheet {
 constructor(name,rows=[]){this.name=name;this.rows=rows}
 getName(){return this.name}
 appendRow(row){this.getRange(this.rows.length+1,1,1,row.length).setValues([row])}
 getLastRow(){return this.rows.length} setFrozenRows(){}
 getDataRange(){return this.getRange(1,1,this.rows.length,Math.max(...this.rows.map(row=>row.length)))}
 getRange(r,c,n=1,m=1){const self=this;return {getValues(){return Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>self.rows[r-1+i]?.[c-1+j]??''))},getFormulas(){return Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>{const v=self.rows[r-1+i]?.[c-1+j];return typeof v==='string'&&v.startsWith('=')?v:''}))},setValues(a){if(self.name==='在庫台帳'&&failLedger)throw Error('write failed');if(self.name==='フォームの回答 1'&&failLog)throw Error('log failed');a.forEach((row,i)=>row.forEach((v,j)=>{self.rows[r-1+i]??=[];self.rows[r-1+i][c-1+j]=v}));return this},setValue(v){return this.setValues([[v]])}}}
}
const ctx={console,SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:n=>sheets[n],insertSheet:n=>sheets[n]=new Sheet(n)}),flush(){}},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},Utilities:{getUuid:()=>String(Math.random())},ContentService:{MimeType:{JSON:'json'},createTextOutput:s=>({setMimeType:()=>JSON.parse(s)})}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync('gas/qr_api.gs','utf8'),ctx);
function reset(){sheets={'在庫台帳':new Sheet('在庫台帳',[['管理番号','メーカー','母材','サブカテゴリ','色','保管場所','重量','担当','備考','フィラメントキー','印刷回数','ラベル注釈','検索別名'],['S26092801','A','PLA','','白','さんらいず',842,'旧担当','備考','','','','']])};sheets['フォームの回答 1']=new Sheet('フォームの回答 1',[['タイムスタンプ','入荷／使用／保管場所移動','管理番号(半角)','重量(g)','担当','色','母材(ベース)','サブカテゴリ','メーカー','保管場所','本数','備考']]);sheets['メーカーマスター']=new Sheet('メーカーマスター',[['メーカーID','メーカー名','標準空スプール重量（g）','メーカー別名','状態'],['MK001','Master Maker',180,'A','正式']]);sheets['フィラメントマスター']=new Sheet('フィラメントマスター',[['フィラメントID','メーカーID','商品名・シリーズ名','素材','サブカテゴリ','色','検索別名','購入先URL','備考','状態'],['FL000001','MK001','Test PLA','PLA','標準','白','ホワイト','','','正式']]);failLedger=failLog=false}
reset();let r=ctx.updateWeight({id:'S26092801',weight:0,person:'新担当'});assert(r.ok);assert.equal(sheets['在庫台帳'].rows[1][6],0);assert.equal(sheets['フォームの回答 1'].rows[1].length,12);assert(sheets['フォームの回答 1'].rows[1][11].includes('842'));assert(sheets['フォームの回答 1'].rows[1][11].includes('完了'));
r=ctx.updateLocation({id:'S26092801',location:'オーシャン',person:'別担当'});assert(r.ok);assert.equal(sheets['在庫台帳'].rows[1][0],'S26092801');assert.equal(sheets['在庫台帳'].rows[1][6],0);
reset();r=ctx.updateWeight({id:'S26092801',weight:'',person:'新担当'});assert(!r.ok);assert.equal(sheets['在庫台帳'].rows[1][6],842);
r=ctx.updateWeight({id:'S26092801',weight:500});assert(!r.ok);
reset();r=ctx.updateWeight({id:'S26092801',weight:642,expectedWeight:800,person:'新担当'});assert(!r.ok);assert.equal(sheets['在庫台帳'].rows[1][6],842);assert.equal(sheets['フォームの回答 1'].rows.length,1);
reset();r=ctx.updateWeight({id:'S26092801',weight:642,expectedWeight:842,person:'新担当'});assert(r.ok);assert.equal(sheets['在庫台帳'].rows[1][6],642);
reset();failLedger=true;r=ctx.updateWeight({id:'S26092801',weight:500,person:'新担当'});assert.equal(r.error,'UPDATE_REQUIRES_REVIEW');assert(sheets['フォームの回答 1'].rows[1][11].includes('要確認'));
reset();failLog=true;r=ctx.updateWeight({id:'S26092801',weight:500,person:'新担当'});assert(!r.ok);assert.equal(sheets['在庫台帳'].rows[1][6],842);
reset();r=ctx.registerFilament({id:'O26092801',filamentKey:'FL000001',location:'オーシャン',person:'担当',weight:1000,note:'なし'});assert(r.ok);assert.equal(sheets['フォームの回答 1'].rows[1][1],'入荷');assert(sheets['在庫台帳'].rows[2][1].startsWith('='));assert(sheets['在庫台帳'].rows[2][1].includes('VLOOKUP(J3'));assert(sheets['在庫台帳'].rows[2][2].includes("'フィラメントマスター'!A:F,4,FALSE"));assert.equal(sheets['在庫台帳'].rows[2][9],'FL000001');assert.equal(sheets['在庫台帳'].rows[2][10],0);assert(sheets['在庫台帳'].rows[2][12].includes("'フィラメントマスター'!A:G,7,FALSE"));assert.equal(r.printCount,0);r=ctx.registerFilament({id:'O26092801',filamentKey:'FL000001',location:'オーシャン',person:'担当'});assert.equal(r.error,'ID_ALREADY_EXISTS');
reset();r=ctx.registerFilament({id:'O26092801',filamentKey:'FL999999',location:'オーシャン',person:'担当',weight:1000,note:'なし'});assert(!r.ok);assert.equal(sheets['在庫台帳'].rows.length,2);
const masterResponse=ctx.doGet({parameter:{action:'filamentMasters'}});assert(masterResponse.ok);assert.equal(masterResponse.items[0].id,'FL000001');assert.equal(masterResponse.items[0].maker,'Master Maker');
console.log('PASS: zero weight, history before/after, stable ID, location-only update, blank weight, missing person, ledger/log failures, registration and duplicate ID.');

reset();const makerResponse=ctx.doGet({parameter:{action:'makerMasters'}});assert(makerResponse.ok);assert.equal(makerResponse.items[0].id,'MK001');assert.equal(makerResponse.items[0].spoolWeight,180);
r=ctx.createFilamentMaster({makerId:'MK001',product:'Test ASA',base:'ASA',sub:'標準',color:'黒',aliases:'ブラック｜Black',purchaseUrl:'https://example.com/asa',note:'test'});assert(r.ok);assert.equal(r.id,'FL000002');assert.equal(sheets['フィラメントマスター'].rows[2][0],'FL000002');assert.equal(sheets['フィラメントマスター'].rows[2][6],'ブラック｜Black');
r=ctx.createFilamentMaster({makerId:'MK001',product:'Duplicate',base:'asa',sub:'標準',color:'黒'});assert(!r.ok);assert.equal(r.error,'MASTER_ALREADY_EXISTS');assert.equal(r.existingId,'FL000002');assert.equal(sheets['フィラメントマスター'].rows.length,3);
r=ctx.createFilamentMaster({makerId:'MK001',base:'PETG',sub:'標準',color:'透明',purchaseUrl:'http://example.com'});assert(!r.ok);assert.equal(sheets['フィラメントマスター'].rows.length,3);
r=ctx.createFilamentMaster({makerId:'MK999',base:'PETG',sub:'標準',color:'透明'});assert(!r.ok);assert.equal(sheets['フィラメントマスター'].rows.length,3);
console.log('PASS: maker list, master auto-numbering, duplicate rejection, URL and maker validation.');

const healthResponse=ctx.doGet({parameter:{action:'health'}});assert(healthResponse.ok);assert(healthResponse.features.includes('expectedWeightCheck'));assert(healthResponse.features.includes('masterBackedInventoryFormulas'));assert(healthResponse.features.includes('masterFinalization'));assert(healthResponse.actions.includes('finalizeFilamentMaster'));assert(healthResponse.features.includes('makerFinalization'));assert(healthResponse.actions.includes('finalizeMakerMaster'));assert(healthResponse.features.includes('unifiedMasterEditing'));assert(healthResponse.actions.includes('updateFilamentMaster'));assert(healthResponse.actions.includes('updateMakerMaster'));
console.log('PASS: health response advertises expected-weight protection.');

reset();sheets['在庫台帳'].rows[1][9]='FL000001';const tareLookup=ctx.doGet({parameter:{id:'S26092801'}});assert(tareLookup.ok);assert.equal(tareLookup.spoolWeight,180);assert(healthResponse.features.includes('tareWeightCheck'));
r=ctx.updateWeight({id:'S26092801',weight:1,grossWeight:830,expectedSpoolWeight:180,expectedWeight:842,person:'担当'});assert(r.ok);assert.equal(r.weight,650);assert.equal(sheets['在庫台帳'].rows[1][6],650);
reset();sheets['在庫台帳'].rows[1][9]='FL000001';r=ctx.updateWeight({id:'S26092801',weight:650,grossWeight:830,expectedSpoolWeight:200,expectedWeight:842,person:'担当'});assert(!r.ok);assert.equal(sheets['在庫台帳'].rows[1][6],842);assert.equal(sheets['フォームの回答 1'].rows.length,1);
reset();sheets['在庫台帳'].rows[1][9]='FL000001';r=ctx.updateWeight({id:'S26092801',weight:0,grossWeight:100,expectedSpoolWeight:180,expectedWeight:842,person:'担当'});assert(r.ok);assert.equal(r.weight,0);
reset();r=ctx.updateWeight({id:'S26092801',weight:650,grossWeight:830,expectedSpoolWeight:180,expectedWeight:842,person:'担当'});assert(!r.ok);assert.equal(sheets['在庫台帳'].rows[1][6],842);
console.log('PASS: tare lookup, server-side gross calculation, stale tare rejection, zero floor, and missing-master rejection.');

reset();r=ctx.createMakerMaster({name:'Quick Maker',provisional:true});assert(r.ok);assert.equal(r.id,'MK002');assert.equal(sheets['メーカーマスター'].rows[2][4],'要整備');
r=ctx.createFilamentMaster({makerId:'MK002',product:'Gold sample',provisional:true});assert(r.ok);assert.equal(r.id,'FL000002');assert.equal(r.base,'未設定');assert.equal(r.status,'要整備');assert.equal(sheets['フィラメントマスター'].rows[2][9],'要整備');
console.log('PASS: provisional maker and filament creation.');

reset();r=ctx.createMakerMaster({name:'Formal Maker',aliases:'正式別名',spoolWeight:210,provisional:false});assert(r.ok);assert.equal(r.status,'正式');assert.deepEqual(sheets['メーカーマスター'].rows[2].slice(1,5),['Formal Maker',210,'正式別名','正式']);
console.log('PASS: standalone formal maker creation.');

reset();sheets['フィラメントマスター'].rows[1]=['FL000001','MK001','届いた金色','未設定','未設定','未設定','','','','要整備'];sheets['在庫台帳'].rows[1][1]='Master Maker';sheets['在庫台帳'].rows[1][2]='未設定';sheets['在庫台帳'].rows[1][3]='未設定';sheets['在庫台帳'].rows[1][4]='未設定';sheets['在庫台帳'].rows[1][9]='FL000001';
r=ctx.masterMaintenancePreview_('FL000001');assert(r.ok);assert.equal(r.affectedCount,1);assert.equal(r.item.temporaryMemo,'届いた金色');
r=ctx.finalizeFilamentMaster({filamentId:'FL000001',makerId:'MK001',base:'PLA',sub:'silk',color:'金',aliases:'ゴールド',purchaseUrl:'https://example.com/gold',note:'正式化'});assert(r.ok);assert.equal(r.affectedCount,1);assert.equal(sheets['フィラメントマスター'].rows[1][2],'届いた金色');assert.equal(sheets['フィラメントマスター'].rows[1][3],'PLA');assert.equal(sheets['フィラメントマスター'].rows[1][9],'正式');assert.deepEqual(sheets['在庫台帳'].rows[1].slice(1,5),['Master Maker','PLA','silk','金']);assert.equal(sheets['フォームの回答 1'].rows.length,1);
r=ctx.finalizeFilamentMaster({filamentId:'FL000001',makerId:'MK001',base:'PLA',sub:'silk',color:'金'});assert(!r.ok);assert(String(r.error).includes('要整備ではありません'));
console.log('PASS: provisional master preview, finalization, linked inventory sync, and no form history row.');

reset();sheets['メーカーマスター'].rows[1]=['MK001','仮メーカー','','旧称','要整備'];sheets['フィラメントマスター'].rows[1]=['FL000001','MK001','仮メモ','PLA','標準','白','','','','正式'];sheets['在庫台帳'].rows[1][1]='仮メーカー';sheets['在庫台帳'].rows[1][9]='FL000001';
r=ctx.makerMaintenancePreview_('MK001');assert(r.ok);assert.equal(r.filamentCount,1);assert.equal(r.affectedCount,1);assert.equal(r.item.name,'仮メーカー');
r=ctx.finalizeMakerMaster({makerId:'MK001',name:'正式メーカー',aliases:'旧称｜英語名',spoolWeight:200});assert(r.ok);assert.equal(r.affectedCount,1);assert.equal(r.filamentCount,1);assert.deepEqual(sheets['メーカーマスター'].rows[1].slice(0,5),['MK001','正式メーカー',200,'旧称｜英語名','正式']);assert.equal(sheets['在庫台帳'].rows[1][1],'正式メーカー');assert.equal(sheets['フォームの回答 1'].rows.length,1);
r=ctx.finalizeMakerMaster({makerId:'MK001',name:'正式メーカー',spoolWeight:200});assert(!r.ok);assert(String(r.error).includes('要整備ではありません'));
console.log('PASS: provisional maker preview, finalization, linked inventory maker sync, and no form history row.');

reset();sheets['フィラメントマスター'].rows[1]=['FL000001','MK001','届いた色','未設定','未設定','未設定','','','','要整備'];sheets['在庫台帳'].rows[1][9]='FL000001';
r=ctx.updateFilamentMaster({filamentId:'FL000001',makerId:'MK001',base:'PLA',sub:'',color:'',aliases:'試作品',purchaseUrl:'',note:'途中'});assert(r.ok);assert.equal(r.status,'要整備');assert.equal(sheets['フィラメントマスター'].rows[1][3],'PLA');assert.equal(sheets['フィラメントマスター'].rows[1][4],'未設定');assert.equal(sheets['在庫台帳'].rows[1][2],'PLA');assert.equal(sheets['フォームの回答 1'].rows.length,1);
r=ctx.updateFilamentMaster({filamentId:'FL000001',makerId:'MK001',base:'PLA',sub:'標準',color:'白',aliases:'ホワイト',purchaseUrl:'',note:'完了'});assert(r.ok);assert.equal(r.status,'正式');assert.deepEqual(sheets['在庫台帳'].rows[1].slice(1,5),['Master Maker','PLA','標準','白']);
r=ctx.updateFilamentMaster({filamentId:'FL000001',makerId:'MK001',base:'',sub:'標準',color:'白'});assert(!r.ok);assert.equal(sheets['フィラメントマスター'].rows[1][9],'正式');
console.log('PASS: provisional partial save, automatic finalization, inventory sync, and formal required fields.');

reset();sheets['在庫台帳'].rows[1][9]='FL000001';r=ctx.updateFilamentMaster({filamentId:'FL000001',makerId:'MK001',base:'PLA',sub:'マット',color:'白',aliases:'白色',purchaseUrl:'https://example.com',note:'編集'});assert(r.ok);assert.equal(r.status,'正式');assert.equal(sheets['在庫台帳'].rows[1][3],'マット');
r=ctx.updateMakerMaster({makerId:'MK001',name:'Renamed Maker',aliases:'旧名',spoolWeight:190});assert(r.ok);assert.equal(sheets['メーカーマスター'].rows[1][1],'Renamed Maker');assert.equal(sheets['在庫台帳'].rows[1][1],'Renamed Maker');assert.equal(sheets['フォームの回答 1'].rows.length,1);
console.log('PASS: formal filament and maker editing with linked inventory sync and no form history row.');

reset();sheets['在庫台帳'].rows[1][9]='FL000001';sheets['在庫台帳'].rows[1].splice(1,4,...ctx.inventoryMasterFormulas_(2));const formulasBefore=sheets['在庫台帳'].rows[1].slice(1,5);
r=ctx.updateFilamentMaster({filamentId:'FL000001',makerId:'MK001',base:'PLA',sub:'マット',color:'白',aliases:'',purchaseUrl:'',note:'数式保持'});assert(r.ok);assert.deepEqual(sheets['在庫台帳'].rows[1].slice(1,5),formulasBefore);
r=ctx.updateMakerMaster({makerId:'MK001',name:'Formula Maker',aliases:'',spoolWeight:190});assert(r.ok);assert.deepEqual(sheets['在庫台帳'].rows[1].slice(1,5),formulasBefore);
console.log('PASS: master edits preserve formula-backed inventory cells.');

reset();const lookupBatch=ctx.doGet({parameter:{action:'lookupBatch',ids:'S26092801,O26092801'}});assert(lookupBatch.ok);assert.equal(lookupBatch.items.length,2);assert.equal(lookupBatch.items[0].found,true);assert.equal(lookupBatch.items[1].found,false);
r=ctx.registerFilamentBatch({items:[
 {id:'O26092801',filamentKey:'FL000001',location:'オーシャン',person:'担当',weight:1000,note:'一括'},
 {id:'O26092802',filamentKey:'FL000001',location:'オーシャン',person:'担当',weight:1000,note:'一括'}
]});assert(r.ok);assert.equal(r.action,'registerBatch');assert.deepEqual(Array.from(r.ids),['O26092801','O26092802']);assert.equal(sheets['在庫台帳'].rows.length,4);assert.equal(sheets['フォームの回答 1'].rows.length,3);assert(sheets['在庫台帳'].rows[2][1].startsWith('='));assert(sheets['在庫台帳'].rows[3][1].startsWith('='));assert(sheets['在庫台帳'].rows[2][1].includes('VLOOKUP(J3'));assert(sheets['在庫台帳'].rows[3][1].includes('VLOOKUP(J4'));assert.equal(sheets['在庫台帳'].rows[2][9],'FL000001');assert.equal(sheets['在庫台帳'].rows[3][10],0);assert(sheets['フォームの回答 1'].rows[1][11].includes('完了'));assert(sheets['フォームの回答 1'].rows[2][11].includes('完了'));
r=ctx.registerFilamentBatch({items:[
 {id:'O26092802',filamentKey:'FL000001',location:'オーシャン',person:'担当',weight:1000,note:'重複'},
 {id:'O26092803',filamentKey:'FL000001',location:'オーシャン',person:'担当',weight:1000,note:'重複'}
]});assert(!r.ok);assert(String(r.error).includes('ID_ALREADY_EXISTS'));assert.equal(sheets['在庫台帳'].rows.length,4);
console.log('PASS: batch lookup, two-item registration, history rows, print count and duplicate rejection.');

reset();failLedger=true;r=ctx.registerFilamentBatch({items:[
 {id:'F26092801',filamentKey:'FL000001',location:'フォージー',person:'担当',weight:1000,note:'一括'},
 {id:'F26092802',filamentKey:'FL000001',location:'フォージー',person:'担当',weight:1000,note:'一括'}
]});assert.equal(r.error,'UPDATE_REQUIRES_REVIEW');assert.equal(sheets['在庫台帳'].rows.length,2);assert(sheets['フォームの回答 1'].rows[1][11].includes('要確認'));assert(sheets['フォームの回答 1'].rows[2][11].includes('要確認'));
console.log('PASS: batch ledger failure requires review and does not silently retry.');

reset();delete sheets['フォームの回答 1'];r=ctx.updateWeight({id:'S26092801',weight:1,person:'担当'});assert(!r.ok);assert.equal(sheets['在庫台帳'].rows[1][6],842);assert(!sheets['操作ログ']);console.log('PASS: missing response sheet blocks mutation, no new sheet created');
