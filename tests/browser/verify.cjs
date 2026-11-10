const { chromium } = require('@playwright/test')
const assert = require('node:assert/strict')
const path = require('node:path')
;(async()=>{
 const browser = await chromium.launch({executablePath:process.env.CHROMIUM_PATH ?? '/usr/bin/chromium',headless:true,args:['--no-sandbox']})
 const page = await browser.newPage()
 await page.goto(new URL('/tests/browser/', process.env.SPOOL_PROOF_URL ?? 'http://localhost:5184').href)
 await page.waitForSelector('#report')
 const result = await page.evaluate(async()=>{
   const target=document.querySelector('#report')
   const canvas=document.createElement('canvas')
   canvas.width=16;canvas.height=16
   canvas.getContext('2d').fillRect(0,0,16,16)
   target.append(canvas)
   const snapshot=await window.spool.prepareDocument({target})
   const repeated=await window.spool.preparePrint({target:snapshot.target})
   if(!repeated.document.body.textContent.includes('Spool repeated header') || !repeated.document.body.textContent.includes('Spool repeated footer')) throw new Error('Reusable snapshot lost repeated print helpers')
   const copiedCanvas=repeated.target.querySelector('canvas')
   if(copiedCanvas.getContext('2d').getImageData(0,0,1,1).data[3]!==255) throw new Error('Reusable snapshot lost cross-document canvas pixels')
   repeated.dispose();snapshot.dispose();canvas.remove()
   const inherited=await window.spool.preparePrint({target})
   const color=inherited.frame.contentWindow.getComputedStyle(inherited.target).color
   let calls=0
   inherited.frame.contentWindow.print=()=>{calls++}
   inherited.print()
   if(calls!==1) throw new Error('Native frame print was not invoked')
   window.prepared=inherited
   const previewHost=document.createElement('section')
   document.body.append(previewHost)
   const visible=await window.spool.preparePrint({target:inherited.target,container:previewHost})
   Object.assign(visible.frame.style,{position:'static',left:'auto'})
   await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)))
   const previewPopulated=visible.frame.parentElement===previewHost && visible.document.querySelector('h1')?.textContent==='Native selectable report' && visible.frame.contentDocument===visible.document
   if(!previewPopulated) throw new Error('Create-in-place preview lost its document')
   const previewColor=visible.frame.contentWindow.getComputedStyle(visible.target).color
   if(previewColor!=='rgb(20, 60, 90)') throw new Error('Cross-document preview lost inherited source styles')
   visible.dispose();previewHost.remove()
   const isolated=await window.spool.preparePrint({target,styles:{mode:'isolated',css:'.report{color:rgb(90, 20, 40)}'}})
   const isolatedColor=isolated.frame.contentWindow.getComputedStyle(isolated.target).color
   const inheritedApp=isolated.document.querySelector('.app')!==null
   isolated.dispose();isolated.dispose()
   const observer = new MutationObserver(records => {
     for(const record of records) for(const node of record.addedNodes) {
       if(node instanceof HTMLIFrameElement) node.contentWindow.print=()=>node.contentWindow.dispatchEvent(new Event('afterprint'))
     }
   })
   observer.observe(document.body,{childList:true})
   await window.spool.print({target})
   observer.disconnect()
   if(document.querySelectorAll('iframe').length!==1) throw new Error('Public print did not clean up after afterprint')
   let missingFailed=false
   try {await window.spool.preparePrint({target,styles:{mode:'isolated',stylesheets:['/missing.css']},timeoutMs:1000})} catch {missingFailed=true}
   let fontFailed=false
   try {await window.spool.preparePrint({target,styles:{mode:'isolated',css:'@font-face{font-family:MissingRequired;src:url(/missing-required-font.woff2)}.report{font-family:MissingRequired}'}})} catch {fontFailed=true}
   if(!fontFailed) throw new Error('Failed export font was silently replaced')
   return {previewPopulated,fontFailed,color,isolatedColor,inheritedApp,missingFailed,frameCount:document.querySelectorAll('iframe').length}
 })
 assert.equal(result.color,'rgb(20, 60, 90)');assert.equal(result.isolatedColor,'rgb(90, 20, 40)');assert.equal(result.inheritedApp,false);assert.equal(result.missingFailed,true);assert.equal(result.frameCount,1)
 const frame=page.frames().find(f=>f!==page.mainFrame())
 await frame.evaluate(()=>{const style=document.createElement('style');style.textContent='@page{size:A4;margin:14mm}';document.head.append(style)})
 const html=await frame.content()
 const pdfPage=await browser.newPage()
 await pdfPage.setContent(html)
 await pdfPage.pdf({path:path.join(__dirname,'report.pdf'),preferCSSPageSize:true,printBackground:true})
 await page.evaluate(()=>window.prepared.dispose())
 assert.equal(await page.locator('iframe').count(),0)
 const orientations = {}
 for (const orientation of ['portrait', 'landscape']) {
   const html = await page.evaluate(async orientation => {
     const prepared = await window.spool.preparePrint({ target: document.querySelector('#report'), orientation })
     try { return prepared.document.documentElement.outerHTML } finally { prepared.dispose() }
   }, orientation)
   const oriented = await browser.newPage()
   await oriented.setContent(html)
   const bytes = await oriented.pdf({ preferCSSPageSize: true })
   const box = /\/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)/.exec(bytes.toString('latin1'))
   assert.ok(box, 'Native PDF page dimensions must be available')
   const width = Number(box[1]), height = Number(box[2])
   assert.equal(width > height, orientation === 'landscape')
   orientations[orientation] = { width, height }
   await oriented.close()
 }
 console.log(JSON.stringify({...result, orientations}))
 await browser.close()
})().catch(e=>{console.error(e);process.exit(1)})
