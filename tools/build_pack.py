"""사용법: python3 build_pack.py <template.html> <cube.json> <cube_i.json> <비밀번호> <출력 index.html>
데이터(결산·일보 큐브)를 비밀번호로 암호화(PBKDF2-SHA256 → AES-256-GCM, 압축 후)하여 index.html 한 파일로 조립합니다."""
import sys, os, json, zlib, base64
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
tpl,cj,ci,pw,out=sys.argv[1:6]
IT=250000
t=open(tpl,encoding='utf-8').read()
J=json.load(open(cj,encoding='utf-8')); I=json.load(open(ci,encoding='utf-8'))
raw=json.dumps({'J':J,'I':I},ensure_ascii=False,separators=(',',':')).encode('utf-8')
co=zlib.compressobj(9,zlib.DEFLATED,-15); comp=co.compress(raw)+co.flush()
salt=os.urandom(16); iv=os.urandom(12)
key=PBKDF2HMAC(algorithm=hashes.SHA256(),length=32,salt=salt,iterations=IT).derive(pw.encode('utf-8'))
ct=AESGCM(key).encrypt(iv,comp,None)
blob=base64.b64encode(salt+iv+ct).decode()
# 앱 스크립트를 부팅 함수로 변환
a="const DATA_J = __DATA_J__;\nconst DATA_I = __DATA_I__;\n(function(){"
assert a in t and t.count('\n})();\n</script>')==1
t=t.replace(a,"window.__boot=function(DATA_J,DATA_I){\n(function(){").replace('\n})();\n</script>','\n})();\n};\n</script>')
icons='''<link rel="icon" type="image/x-icon" href="./icons/favicon.ico">
<link rel="icon" type="image/png" sizes="32x32" href="./icons/favicon-32x32.png">
<link rel="icon" type="image/png" sizes="16x16" href="./icons/favicon-16x16.png">
<link rel="apple-touch-icon" sizes="180x180" href="./icons/apple-touch-icon.png">
<link rel="manifest" href="./icons/site.webmanifest">
<meta name="robots" content="noindex,nofollow">
'''
gate_css='''<style>
#gate{position:fixed;inset:0;z-index:99999;background:linear-gradient(160deg,#e8eeeb,#d9e4df);display:flex;align-items:center;justify-content:center;padding:16px}
#gate .gbox{background:#fff;border-radius:20px;padding:28px 24px;width:100%;max-width:340px;box-shadow:0 10px 40px rgba(20,40,60,.18);text-align:center;font-family:inherit}
#gate h1{font-size:18px;margin:0 0 4px;color:#1d2733}#gate p{font-size:13px;color:#4b5663;margin:0 0 16px}
#gate input{width:100%;box-sizing:border-box;padding:12px;border:1px solid #cfd8d5;border-radius:12px;font-size:16px;font:inherit;font-size:16px}
#gate button{margin-top:12px;width:100%;padding:12px;border:0;border-radius:12px;background:#34405a;color:#fff;font:inherit;font-weight:700;font-size:15px;cursor:pointer}
#gate .err{min-height:18px;margin-top:10px;font-size:12.5px;color:#b4501f}
</style>
'''
gate_html='''<div id="gate" role="dialog" aria-modal="true" aria-label="접속 확인"><form class="gbox" id="gform" autocomplete="off">
<h1>판매·손익 대시보드</h1><p>접속 비밀번호를 입력하세요</p>
<input id="gpw" type="password" placeholder="비밀번호" autocomplete="current-password" aria-label="비밀번호" autofocus>
<button type="submit" id="gbtn">확인</button><div class="err" id="gerr" role="alert"></div></form></div>
'''
gate_js='''<script>
(function(){
var BLOB="%s",IT=%d;
var $=function(i){return document.getElementById(i);};
function b64(s){var b=atob(s),u=new Uint8Array(b.length);for(var i=0;i<b.length;i++)u[i]=b.charCodeAt(i);return u;}
async function unlock(pw){
  var raw=b64(BLOB),salt=raw.slice(0,16),iv=raw.slice(16,28),ct=raw.slice(28);
  var km=await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveKey']);
  var key=await crypto.subtle.deriveKey({name:'PBKDF2',salt:salt,iterations:IT,hash:'SHA-256'},km,{name:'AES-GCM',length:256},false,['decrypt']);
  var pt=await crypto.subtle.decrypt({name:'AES-GCM',iv:iv},key,ct);
  var txt=await new Response(new Blob([pt]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text();
  return JSON.parse(txt);
}
async function go(pw,silent){
  var err=$('gerr'),btn=$('gbtn');
  if(!(window.crypto&&crypto.subtle)){err.textContent='보안 연결(https)에서만 열 수 있습니다.';return;}
  btn.disabled=true;btn.textContent='확인 중…';err.textContent='';
  try{
    var d=await unlock(pw);
    try{sessionStorage.setItem('dash_ok',pw);}catch(e){}
    $('gate').remove();window.__boot(d.J,d.I);
  }catch(e){
    try{sessionStorage.removeItem('dash_ok');}catch(x){}
    btn.disabled=false;btn.textContent='확인';
    if(!silent){err.textContent='비밀번호가 올바르지 않습니다.';$('gpw').value='';$('gpw').focus();}
  }
}
$('gform').addEventListener('submit',function(e){e.preventDefault();var v=$('gpw').value;if(v)go(v,false);});
var saved=null;try{saved=sessionStorage.getItem('dash_ok');}catch(e){}
if(saved)go(saved,true);
})();
</script>
''' % (blob,IT)
t=t.replace('<title>',icons+'<title>',1)
t=t.replace('</head>',gate_css+'</head>',1)
i=t.index('<body>')+len('<body>')
t=t[:i]+'\n'+gate_html+t[i:]
# 게이트 스크립트는 앱 스크립트 뒤(__boot 정의 후)에 실행
j=t.rindex('</script>')+len('</script>')
t=t[:j]+'\n'+gate_js+t[j:]
open(out,'w',encoding='utf-8').write(t)
print('index.html',round(os.path.getsize(out)/1e6,2),'MB; payload raw',round(len(raw)/1e6,2),'-> enc',round(len(blob)/1e6,2),'MB')
