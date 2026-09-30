import fs from 'node:fs'
const index=fs.readFileSync('Index.html','utf8')
const style=fs.readFileSync('Styles.html','utf8')
const js=fs.readFileSync('JavaScript.html','utf8')
const counties={Mombasa:['Mvita','Nyali','Likoni','Changamwe','Jomvu','Kisauni'],Kwale:['Msambweni','Lunga Lunga','Matuga','Kinango'],Kilifi:['Kilifi North','Kilifi South','Kaloleni','Rabai','Ganze','Malindi','Magarini'],Lamu:['Lamu East','Lamu West'],'Tana River':['Bura','Galole','Garsen'],'Taita Taveta':['Voi','Mwatate','Wundanyi','Taveta']}
const demo=index.replace("<?!= include('Styles'); ?>",style+`<style>.source-download{position:fixed;right:18px;bottom:18px;z-index:15;background:#e8c980;color:#173c41;border-radius:9px;padding:11px 15px;font:700 12px Inter,sans-serif;text-decoration:none;box-shadow:0 5px 20px #0003}.source-download:hover{background:#f5d998}.preview-pill{position:fixed;top:18px;left:52%;z-index:12;background:#e9f4f1;color:#08776c;padding:6px 11px;border-radius:20px;font:700 10px Inter,sans-serif;letter-spacing:.5px}@media(max-width:720px){.preview-pill{display:none}}</style>`).replace('<?!= counties ?>',JSON.stringify(counties)).replace("<?!= include('JavaScript'); ?>",`<script src="/mock.js"></script>${js}<a class="source-download" href="/downloads/README.md" download="README.md">↓ &nbsp; Deployment guide</a><div class="preview-pill" id="previewPill">LIVE APPLICATION · GOOGLE SHEETS</div>`)
fs.mkdirSync('public/downloads',{recursive:true})
for(const file of ['Code.gs','Index.html','Styles.html','JavaScript.html','appsscript.json','README.md'])fs.copyFileSync(file,'public/downloads/'+file)
fs.writeFileSync('public/demo.html',demo)
console.log('Built preview and copied deployment files.')
