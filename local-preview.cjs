// ローカルの画面確認専用。外部公開せず、このPCの127.0.0.1だけで待ち受ける。
const http=require('http'),fs=require('fs'),path=require('path');
const root=__dirname;
const types={'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json','.png':'image/png','.ico':'image/x-icon'};
const server=http.createServer((req,res)=>{
 if(req.method!=='GET'&&req.method!=='HEAD'){res.writeHead(405);return res.end()}
 let file;try{const url=new URL(req.url,'http://localhost');file=path.resolve(root,'.'+decodeURIComponent(url.pathname==='/'?'/register-start.html':url.pathname))}catch{res.writeHead(400);return res.end()}
 if(!file.startsWith(root+path.sep)||!types[path.extname(file)]){res.writeHead(404);return res.end()}
 fs.readFile(file,(error,data)=>{if(error){res.writeHead(404);return res.end()}res.writeHead(200,{'Content-Type':types[path.extname(file)],'Cache-Control':'no-store'});res.end(req.method==='HEAD'?undefined:data)});
});
server.on('error',error=>{console.error(error.message);process.exitCode=1});
server.listen(0,'127.0.0.1',()=>{console.log('Open this URL in your browser:');console.log('http://127.0.0.1:'+server.address().port+'/');console.log('Keep this window open. Ctrl+C stops the preview.');console.log('Registration buttons connect to the real inventory.');});
