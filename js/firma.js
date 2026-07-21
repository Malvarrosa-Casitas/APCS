
let canvas,ctx,drawing=false;
function pos(e){const r=canvas.getBoundingClientRect();const p=e.touches?e.touches[0]:e;return{x:p.clientX-r.left,y:p.clientY-r.top}}
function start(e){drawing=true;ctx.beginPath();const p=pos(e);ctx.moveTo(p.x,p.y);e.preventDefault()}
function move(e){if(!drawing)return;const p=pos(e);ctx.lineTo(p.x,p.y);ctx.stroke();e.preventDefault()}
window.getSignatureData=()=>canvas?canvas.toDataURL('image/png'):'';window.clearSignature=()=>{if(ctx)ctx.clearRect(0,0,canvas.width,canvas.height)};
document.addEventListener('DOMContentLoaded',()=>{canvas=document.querySelector('.signature');if(!canvas)return;function resize(){const old=canvas.toDataURL();canvas.width=canvas.offsetWidth;canvas.height=220;ctx=canvas.getContext('2d');ctx.lineWidth=3;ctx.lineCap='round';ctx.strokeStyle='#6c438b';const img=new Image();img.onload=()=>ctx.drawImage(img,0,0,canvas.width,canvas.height);img.src=old}resize();addEventListener('resize',resize);['mousedown','touchstart'].forEach(e=>canvas.addEventListener(e,start,{passive:false}));['mousemove','touchmove'].forEach(e=>canvas.addEventListener(e,move,{passive:false}));['mouseup','mouseleave','touchend','touchcancel'].forEach(e=>canvas.addEventListener(e,()=>drawing=false))});
