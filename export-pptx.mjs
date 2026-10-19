import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const here = path.dirname(fileURLToPath(import.meta.url));
const workspaceDir = path.resolve(here, '..');
const SKILL_DIR = 'C:\\Users\\sande\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.1011.11113\\skills\\presentations';
const RUNTIME_PYTHON = 'C:\\Users\\sande\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe';
const buildDir = path.join(workspaceDir, '.build-webdev');
const outputDir = here;
const heroPath = path.join(here, 'assets', 'web-development-network-hero.png');
const C = { bg:'#07152F', bg2:'#0D2450', panel:'#102D60', panel2:'#132F5F', white:'#F7FAFF', muted:'#B8C7E2', cyan:'#44D7FF', purple:'#A780FF', mint:'#77F3D0', gold:'#F4C96D', pink:'#FA8CFF', line:'#315B96' };
const FONT = 'Arial';

function shape(slide, geometry, position, fill = 'none', line = { style:'solid', fill:'none', width:0 }) {
  return slide.shapes.add({ geometry, position, fill, line });
}
function text(slide, value, left, top, width, height, opts = {}) {
  const box = shape(slide, 'textbox', { left, top, width, height });
  box.text = value;
  box.text.insets = opts.insets ?? 0;
  box.text.style = { fontSize: opts.size ?? 20, color: opts.color ?? C.white, bold: opts.bold ?? false, fontFamily: FONT, alignment: opts.align ?? 'left', ...opts.style };
  return box;
}
function round(slide, left, top, width, height, fill, radius = 'rounded-xl', line = { style:'solid', fill:'none', width:0 }) {
  return slide.shapes.add({ geometry:'roundRect', position:{left,top,width,height}, fill, line, borderRadius:radius });
}
function footer(slide, number) {
  text(slide, 'INTRODUCTION TO WEB DEVELOPMENT', 66, 685, 360, 16, { size:8, color:'#6E91C6', bold:true });
  text(slide, String(number).padStart(2,'0'), 1160, 685, 50, 16, { size:8, color:'#6E91C6', bold:true, align:'right' });
}
function title(slide, eyebrow, heading, accent) {
  text(slide, eyebrow.toUpperCase(), 68, 56, 720, 20, { size:10, color:C.cyan, bold:true });
  text(slide, heading, 68, 88, 1020, 64, { size:40, color:C.white, bold:true });
  shape(slide, 'rect', { left:68, top:159, width:74, height:4 }, C.cyan);
  if (accent) text(slide, accent, 890, 62, 320, 18, { size:10, color:C.muted, align:'right' });
}
function tag(slide, value, left, top, width, color = C.cyan) {
  const b = round(slide, left, top, width, 26, '#0C244E', 'rounded-full', { style:'solid', fill:color, width:1 });
  b.text = value.toUpperCase(); b.text.style = { fontSize:9, color, bold:true, fontFamily:FONT, alignment:'center' };
  return b;
}
function highlight(slide, value, left, top, width, height = 56) {
  round(slide, left, top, width, height, '#123C70', 'rounded-lg', { style:'solid', fill:'#2F83B9', width:1 });
  shape(slide, 'rect', { left, top, width:4, height }, C.cyan);
  text(slide, value, left+18, top+12, width-34, height-18, { size:16, color:'#E7FBFF', bold:true });
}
function bullet(slide, heading, detail, left, top, number, width = 485) {
  text(slide, String(number).padStart(2,'0'), left, top+3, 35, 18, { size:10, color:C.cyan, bold:true });
  text(slide, heading, left+45, top, width-45, 21, { size:16, color:C.white, bold:true });
  text(slide, detail, left+45, top+25, width-45, 38, { size:13, color:C.muted });
}
function addNode(slide, label, sublabel, left, top, width, height, color = C.cyan) {
  const n = round(slide, left, top, width, height, C.panel, 'rounded-xl', { style:'solid', fill:C.line, width:1 });
  shape(slide, 'ellipse', { left:left+width/2-18, top:top+16, width:36, height:36 }, '#133D72', { style:'solid', fill:color, width:2 });
  text(slide, label, left+10, top+62, width-20, 21, { size:15, color:C.white, bold:true, align:'center' });
  if (sublabel) text(slide, sublabel, left+10, top+87, width-20, height-95, { size:10, color:C.muted, align:'center' });
  return n;
}
function connect(slide, source, target, color = C.cyan, fromSide = 'right', toSide = 'left') {
  return slide.shapes.connect(source, target, { kind:'elbow', fromSide, toSide, line:{style:'solid', fill:color, width:2}, head:{type:'triangle',width:'sm',length:'sm'} });
}
async function imageBytes(imagePath) {
  const bytes = await fs.readFile(imagePath); return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}
async function newOutputPath() {
  const base = path.join(outputDir, 'web-development-introduction.pptx');
  try { await fs.access(base); } catch { return base; }
  for (let i = 2; i < 100; i += 1) {
    const candidate = path.join(outputDir, `web-development-introduction-v${i}.pptx`);
    try { await fs.access(candidate); } catch { return candidate; }
  }
  throw new Error('Could not find a free output name.');
}

async function buildDeck() {
  await fs.mkdir(buildDir, { recursive:true });
  const hero = await imageBytes(heroPath);
  const pres = Presentation.create({ slideSize:{ width:1280, height:720 } });
  const addSlide = () => { const s = pres.slides.add(); s.background.fill = C.bg; return s; };

  // 1. Cover
  { const s = addSlide();
    s.images.add({ blob:hero, contentType:'image/png', alt:'Web development network illustration', fit:'cover', position:{left:0,top:0,width:1280,height:720}, geometry:'rect' });
    shape(s,'rect',{left:0,top:0,width:1280,height:720},'#07152F/70'); shape(s,'rect',{left:0,top:0,width:720,height:720},'#07152F/90');
    tag(s,'BTech CSE classroom presentation',72,148,258); text(s,'Introduction to\nWeb Development',72,202,620,145,{size:56,color:C.white,bold:true});
    text(s,'Understanding the Internet, Browser, Client, and Server',72,375,560,58,{size:23,color:'#D6E6FF'}); shape(s,'rect',{left:72,top:458,width:110,height:4},C.cyan);
    text(s,'HTML · CSS · JavaScript · Client–Server Architecture',72,483,480,24,{size:12,color:'#C8D8F3',bold:true});
    tag(s,'12-slide learning guide',940,627,215,C.cyan); s.speakerNotes.text='Introduce the four ideas that explain how a browser reaches and displays a website.';
  }

  // 2. What is web development
  { const s = addSlide(); title(s,'01 · The big picture','What Is Web Development?');
    text(s,'Web development is the process of creating, building, and maintaining websites and web applications.',68,198,520,110,{size:24,color:C.muted});
    highlight(s,'Web development combines design, programming, and data management to create functional websites.',68,346,525,82);
    const f = addNode(s,'Frontend','What users see and interact with',710,192,150,150,C.cyan); const b = addNode(s,'Backend','Server logic and data processing',930,192,150,150,C.purple); const fs = addNode(s,'Full stack','Works across both areas',820,423,150,150,C.mint);
    connect(s,f,fs,C.cyan,'bottom','left'); connect(s,b,fs,C.purple,'bottom','right'); text(s,'USER EXPERIENCE',710,365,150,16,{size:9,color:C.cyan,bold:true,align:'center'}); text(s,'DATA + LOGIC',930,365,150,16,{size:9,color:C.purple,bold:true,align:'center'}); footer(s,2); s.speakerNotes.text='Describe the frontend as the visible part, backend as the service and data layer, and full stack as work across both.';
  }

  // 3. Internet
  { const s = addSlide(); title(s,'02 · Foundation','Internet: Definition and Overview');
    text(s,'A global network of interconnected networks that lets devices communicate and exchange information.',68,185,530,70,{size:21,color:C.muted});
    bullet(s,'Network of networks','Links local, private, academic, business, and government networks.',68,284,1); bullet(s,'Physical backbone','Uses cables, fiber-optic links, data centers, wireless, and satellites.',68,360,2); bullet(s,'Communication protocols','TCP/IP defines common rules for transferring data in packets.',68,436,3); bullet(s,'Decentralized infrastructure','Many network operators and providers maintain connected parts.',68,512,4);
    const pc=addNode(s,'Devices','Computers and phones',710,215,112,126,C.cyan); const router=addNode(s,'Router','Local network',875,165,112,126,C.purple); const isp=addNode(s,'ISP','Internet access',1030,240,112,126,C.cyan); const server=addNode(s,'Servers','Apps and data',865,425,145,126,C.mint); const fiber=addNode(s,'Fiber links','Physical backbone',1055,470,112,105,C.gold);
    connect(s,pc,router,C.cyan); connect(s,router,isp,C.purple); connect(s,router,server,C.cyan,'bottom','top'); connect(s,server,fiber,C.mint); highlight(s,'The Internet is a physical network infrastructure that connects billions of devices worldwide.',655,604,540,54); footer(s,3); s.speakerNotes.text='Emphasize that the Internet is physical infrastructure and protocols, not simply websites.';
  }

  // 4. How it works
  { const s = addSlide(); title(s,'03 · Request and response','How the Internet Works'); text(s,'Opening a website follows a clear request-and-response path.',68,184,640,35,{size:20,color:C.muted});
    const steps=[['1','Enter a web address','www.example.com',C.cyan],['2','DNS lookup','Finds the IP address',C.purple],['3','HTTP/HTTPS request','Travels across the network',C.cyan],['4','Server response','Returns web resources',C.mint],['5','Browser renders','Displays the webpage',C.gold]];
    const nodes=[]; steps.forEach((item,i)=>{ const left=72+i*230; const n=round(s,left,295,180,170,'#102D60','rounded-xl',{style:'solid',fill:C.line,width:1}); text(s,item[0],left+15,312,20,18,{size:11,color:item[3],bold:true}); shape(s,'ellipse',{left:left+68,top:340,width:44,height:44},'#153C73',{style:'solid',fill:item[3],width:2}); text(s,item[1],left+13,398,154,30,{size:14,color:C.white,bold:true,align:'center'}); text(s,item[2],left+14,439,152,26,{size:10,color:C.muted,align:'center'}); nodes.push(n); });
    nodes.slice(0,-1).forEach((n,i)=>connect(s,n,nodes[i+1],steps[i][3])); highlight(s,'DNS finds the IP address, HTTP/HTTPS transfers web resources, and the browser displays the webpage.',105,555,1070,56); footer(s,4); s.speakerNotes.text='Walk through the numbered request path from URL to the rendered page.';
  }

  // 5. Browser
  { const s = addSlide(); title(s,'04 · The user interface','Browser: Definition and Examples');
    text(s,'A web browser is software that lets users access, view, and interact with websites on the World Wide Web.',68,190,490,102,{size:22,color:C.muted});
    const journey=[['Request','Enter a URL or click a link'],['Lookup','DNS locates the server'],['Retrieval','HTTP/HTTPS gets resources'],['Rendering','HTML, CSS, and JavaScript become a page']]; journey.forEach((p,i)=>{const y=322+i*56; text(s,p[0],70,y,115,20,{size:14,color:i%2?C.purple:C.cyan,bold:true}); text(s,p[1],190,y,360,25,{size:13,color:C.muted}); shape(s,'rect',{left:70,top:y+29,width:482,height:1},'#284E82');});
    const mock=round(s,650,154,545,400,'#0D2C5B','rounded-xl',{style:'solid',fill:'#4986C8',width:1}); shape(s,'rect',{left:650,top:154,width:545,height:42},'#123A73'); ['#FA8CFF','#F4C96D','#77F3D0'].forEach((c,i)=>shape(s,'ellipse',{left:668+i*18,top:170,width:8,height:8},c)); round(s,746,166,365,18,'#27538C','rounded-full'); shape(s,'rect',{left:678,top:217,width:489,height:182},'#1B4D95'); text(s,'Website content',710,264,200,28,{size:25,color:C.white,bold:true}); text(s,'HTML · CSS · JavaScript',710,307,220,20,{size:13,color:'#D0E2FF'}); ['#264F88','#264F88','#264F88'].forEach((c,i)=>round(s,678+i*162,420,145,82,c,'rounded-lg')); text(s,'Popular browsers: Chrome, Firefox, Edge, Safari, Opera, Brave',658,574,535,28,{size:12,color:C.muted,align:'center'});
    highlight(s,'A web browser retrieves website resources and renders them into an interactive webpage.',68,597,1000,52); footer(s,5); s.speakerNotes.text='Explain rendering as the browser interpreting HTML, CSS, JavaScript, images, and other files.';
  }

  // 6. Browser vs search engine
  { const s = addSlide(); title(s,'05 · Clear the confusion','Browser vs Search Engine');
    const table=s.tables.add({rows:4,columns:2,left:120,top:205,width:1040,height:310,values:[['Browser','Search engine'],['Software application used on a device','Online service that helps users discover information and websites'],['Opens and interacts with websites','Uses an index to return relevant results'],['Chrome, Firefox, Safari, Edge','Google Search, Bing, DuckDuckGo']]});
    table.borders.assign({style:'solid',fill:'#315B96',width:1}); table.cells.block({row:0,column:0,rowCount:1,columnCount:2}).assign({fill:'#123D76',textStyle:{fontFamily:FONT,fontSize:19,bold:true,color:C.white}}); table.cells.block({row:1,column:0,rowCount:3,columnCount:2}).assign({fill:'#0C254F',textStyle:{fontFamily:FONT,fontSize:14,color:C.muted}});
    shape(s,'ellipse',{left:613,top:352,width:54,height:54},'#0B1D3E',{style:'solid',fill:C.purple,width:2}); text(s,'VS',622,370,36,20,{size:12,color:C.white,bold:true,align:'center'}); highlight(s,'Google Chrome is a browser; Google Search is a search engine.',205,565,870,55); footer(s,6); s.speakerNotes.text='Use the table to remove a common misconception. A browser opens websites. A search engine helps find websites and information.';
  }

  // 7. Client
  { const s = addSlide(); title(s,'06 · Request origin','Client: Definition and Role');
    text(s,'A client is a device or software application that sends requests to a server for data or services.',68,190,510,95,{size:23,color:C.muted});
    tag(s,'Laptop browser',68,321,140,C.cyan); tag(s,'Mobile app',220,321,124,C.purple); tag(s,'Email app',356,321,120,C.mint);
    text(s,'Client responsibilities',68,393,250,24,{size:16,color:C.white,bold:true}); ['Sends requests to servers','Displays received information','Collects user input','Communicates with backend APIs'].forEach((x,i)=>{shape(s,'ellipse',{left:72,top:432+i*35,width:9,height:9},C.cyan);text(s,x,94,426+i*35,380,22,{size:14,color:C.muted});});
    const user=addNode(s,'User','Chooses a page or action',700,250,130,140,C.cyan); const browser=addNode(s,'Browser client','Displays info and sends requests',900,205,185,185,C.purple); const server=addNode(s,'Server','Returns data or services',1120,270,105,125,C.mint); connect(s,user,browser,C.cyan);connect(s,browser,server,C.purple); text(s,'request',832,295,65,16,{size:10,color:C.cyan,align:'center'}); text(s,'API request',1085,292,62,16,{size:10,color:C.purple,align:'center'});
    highlight(s,'The client initiates requests; the server responds to those requests.',680,545,515,56); footer(s,7); s.speakerNotes.text='Point out that a browser commonly acts as the client when a student opens a website.';
  }

  // 8. Server
  { const s = addSlide(); title(s,'07 · The service provider','Server: Definition and Types'); text(s,'A server is a hardware system or software program that processes client requests and provides data, resources, or services over a network.',68,184,950,55,{size:20,color:C.muted});
    const cards=[['Web server','Delivers website resources and HTTP responses. Apache, Nginx.',C.cyan],['Database server','Stores and retrieves data. MySQL, PostgreSQL, MongoDB.',C.purple],['Application server','Runs application logic and business rules.',C.mint],['File server','Stores and shares files and backups.',C.gold],['Mail server','Handles email sending, receiving, and storage.',C.pink]];
    cards.forEach((c,i)=>{ const x=68+i*230; round(s,x,290,198,214,'#102D60','rounded-xl',{style:'solid',fill:C.line,width:1}); shape(s,'rect',{left:x,top:290,width:198,height:4},c[2]); shape(s,'ellipse',{left:x+18,top:322,width:36,height:36}, '#183D73',{style:'solid',fill:c[2],width:2}); text(s,c[0],x+18,380,164,40,{size:16,color:C.white,bold:true}); text(s,c[1],x+18,430,162,55,{size:11,color:C.muted}); });
    highlight(s,'A server receives requests, processes them, and returns appropriate responses.',154,575,970,54); footer(s,8); s.speakerNotes.text='Describe each server type as a focused role that can appear in a modern web application.';
  }

  // 9. Architecture
  { const s = addSlide(); title(s,'08 · Fundamental web model','Client–Server Architecture');
    const user=addNode(s,'User','',70,278,125,125,C.cyan); const browser=addNode(s,'Browser client','Sends request',290,255,145,165,C.cyan); const internet=addNode(s,'Internet','Transfers packets',525,255,145,165,C.purple); const server=addNode(s,'Web server','Processes request',760,255,145,165,C.mint); const data=addNode(s,'Application logic\n& database','Runs logic and gets data',995,242,190,190,C.gold);
    connect(s,user,browser,C.cyan);connect(s,browser,internet,C.cyan);connect(s,internet,server,C.purple);connect(s,server,data,C.mint); ['request','HTTPS','request','data'].forEach((t,i)=>text(s,t,202+i*232,232,70,16,{size:10,color:i===2?C.purple:C.cyan,align:'center'}));
    const response = shape(s,'connector',{left:158,top:483,width:932,height:56},'none',{style:'dashed',fill:C.purple,width:2}); response.text=''; text(s,'server response returns to the browser, then the browser renders it for the user',320,526,640,18,{size:11,color:'#D8C9FF',align:'center'}); highlight(s,'Most interactive websites depend on communication between clients, servers, and often databases.',150,595,975,52); footer(s,9); s.speakerNotes.text='Follow the arrows from user through browser, Internet, server, application logic, and back as a response.';
  }

  // 10. Shopping example
  { const s = addSlide(); title(s,'09 · In context','Online Shopping Example');
    text(s,'A product search shows how the Internet, browser, client, and server work together.',68,185,500,65,{size:20,color:C.muted});
    const steps=['User opens a browser and visits a shop.','The browser uses the Internet to reach the website.','The browser client requests product information.','The server gets product details from a database.','The browser displays products on screen.']; steps.forEach((p,i)=>{shape(s,'ellipse',{left:72,top:286+i*55,width:24,height:24},'#133E71',{style:'solid',fill:C.cyan,width:1}); text(s,String(i+1),72,291+i*55,24,14,{size:10,color:C.cyan,bold:true,align:'center'});text(s,p,112,283+i*55,470,32,{size:14,color:C.muted});});
    round(s,700,215,230,280,'#0F326B','rounded-xl',{style:'solid',fill:'#4E92D3',width:1}); shape(s,'rect',{left:700,top:215,width:230,height:36},'#1C4E98'); text(s,'MyStore',717,225,85,15,{size:10,color:C.white,bold:true}); text(s,'⌕',887,222,25,18,{size:16,color:C.white,bold:true,align:'right'}); for(let i=0;i<4;i+=1)round(s,720+(i%2)*94,275+Math.floor(i/2)*82,76,62,i%2? '#543C9B':'#2066B5','rounded-lg'); text(s,'Browser client',728,460,170,18,{size:11,color:C.muted,bold:true,align:'center'});
    const server=addNode(s,'Server','Gets product details',1035,264,130,135,C.purple); const db=addNode(s,'Database','Product data',1035,445,130,112,C.mint); connect(s,server,db,C.mint,'bottom','top'); connect(s,shape(s,'rect',{left:930,top:350,width:5,height:5},'none'),server,C.cyan); text(s,'client request',925,320,100,16,{size:10,color:C.cyan,align:'right'}); text(s,'product response',835,390,130,16,{size:10,color:C.purple,align:'right'});
    highlight(s,'The Internet connects the systems, the browser displays the website, the client sends requests, and the server processes them.',120,602,1040,50); footer(s,10); s.speakerNotes.text='Use a familiar shopping site to relate each technical concept to a visible user action.';
  }

  // 11. Revision
  { const s = addSlide(); title(s,'10 · Exam preparation','Quick Revision');
    const items=[['Internet','A global network of interconnected networks.',C.cyan],['Browser','Software used to access and display websites.',C.purple],['Client','A device or application that requests resources.',C.mint],['Server','A system that processes requests and provides resources.',C.gold]];
    items.forEach((item,i)=>{const x=68+i*288;round(s,x,200,260,176,'#102D60','rounded-xl',{style:'solid',fill:C.line,width:1});shape(s,'ellipse',{left:x+22,top:222,width:36,height:36},'#143C72',{style:'solid',fill:item[2],width:2});text(s,item[0],x+22,282,205,28,{size:19,color:C.white,bold:true});text(s,item[1],x+22,321,215,38,{size:13,color:C.muted});});
    round(s,68,452,1144,134,'#102C5B','rounded-xl',{style:'solid',fill:'#347AB5',width:1}); text(s,'Quick questions',95,477,165,22,{size:15,color:C.cyan,bold:true}); ['1. What is the difference between a browser and a search engine?','2. What is the role of TCP/IP?','3. Explain the client–server model.'].forEach((q,i)=>text(s,q,300,470+i*33,820,22,{size:14,color:C.white})); footer(s,11); s.speakerNotes.text='Pause on this slide and invite students to answer the questions before closing.';
  }

  // 12. Close
  { const s = addSlide();
    s.images.add({ blob:hero, contentType:'image/png', alt:'Web development network illustration', fit:'cover', position:{left:0,top:0,width:1280,height:720}, geometry:'rect' }); shape(s,'rect',{left:0,top:0,width:1280,height:720},'#07152F/72'); shape(s,'rect',{left:0,top:0,width:750,height:720},'#07152F/82'); tag(s,'End of presentation',82,182,165); text(s,'Thank You!',82,245,590,84,{size:66,color:C.white,bold:true}); text(s,'Questions and Discussion',82,353,420,34,{size:25,color:'#D7E6FF'}); shape(s,'rect',{left:82,top:428,width:96,height:4},C.cyan); text(s,'Internet  ·  Browser  ·  Client  ·  Server',82,455,390,22,{size:13,color:'#BBD0EF',bold:true}); s.speakerNotes.text='Thank the audience and invite questions about the Internet, browser, client, or server.';
  }
  return pres;
}

async function main() {
  const presentation = await buildDeck();
  const FINAL_PPTX = await newOutputPath();
  const { finalizePresentation } = await import(pathToFileURL(path.join(SKILL_DIR, 'container_tools', 'artifact_tool_utils.mjs')).href);
  const stagingDir = path.join(workspaceDir, '.codex-finalizer'); await fs.mkdir(stagingDir,{recursive:true});
  const candidatePath = path.join(stagingDir, 'web-development-candidate.pptx');
  await (await PresentationFile.exportPptx(presentation)).save(candidatePath);
  const requirements = { explicitTotalSlideCount:12, requiredNativeTableOwnerSlides:[6] };
  const result = await finalizePresentation({ ...requirements, workspaceDir, candidatePath, finalPath:FINAL_PPTX, pythonExecutable:RUNTIME_PYTHON, integrityValidatorPath:path.join(SKILL_DIR,'container_tools','inspect_presentation_package_integrity.py'), layoutValidatorPath:path.join(SKILL_DIR,'container_tools','inspect_presentation_layout_geometry.py'), layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-heading-fit','--require-native-table-slide','6'], requiredNativeTableOwnerSlides:[6], fontPolicy:{basis:'design',families:[FONT]}, verifyArtifactToolImport:true, receiptPath:path.join(stagingDir,`${path.basename(FINAL_PPTX)}.validation.json`) });
  console.log(JSON.stringify({ output: FINAL_PPTX, result }, null, 2));
}

main().catch(error => { console.error(error); process.exitCode = 1; });
