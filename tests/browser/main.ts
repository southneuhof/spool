import { createApp, defineComponent, h } from 'vue'
import { PrintHeader, PrintFooter, PrintKeepTogether, PrintPageBreak } from '../../src/vue/print'
import { prepareDocument, preparePrint, print } from '../../src/print'
const Report = defineComponent({ setup() { return () => h('article', {class:'report',id:'report'}, [
  h(PrintHeader,{repeat:'every-page'},()=> 'Spool repeated header'),
  h('h1','Native selectable report'),
  h('table',Array.from({length:90},(_,i)=>h('tr',[h('td',`Row ${i+1}`),h('td','Selectable export content')]))),
  h(PrintPageBreak),h(PrintKeepTogether,{},()=>[h('h2','Together heading'),h('p','Together content')]),
  h(PrintFooter,{},()=> 'Spool repeated footer'),
]) } })
createApp(Report).mount('#app')
Object.assign(window,{spool:{prepareDocument,preparePrint,print}})
